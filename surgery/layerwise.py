"""逐层（Layer-by-Layer）流式执行引擎：7B 不整体加载，常驻内存只有 embedding + final norm + lm_head + 当前一层。

7B（bf16）常驻约 2.2GB（embed 1.09GB + lm_head 1.09GB），每层约 0.47GB 用完即释放 →
进程峰值 ≈ 3–4GB，远低于 16–20GB 红线；代价是每次“全模型前向”要从磁盘读一遍 15GB
（Windows 页缓存会接住大部分，第二遍起很快）。

所有探针一起逐层推进：每层只加载一次，所有序列都过完这一层再换下一层。因此“测一次全部探针”
只需读一遍权重。

功能
  forward_last(prompts)          → 每条 prompt 末位置 logits（用于 pair_acc 指标）
  collect_mlp(seqs, layers)      → 指定层 MLP 神经元激活（down_proj 输入），用于复读神经元打分
  scale_neurons_on_disk(...)     → 逐分片改写 down_proj 列（neuron_scale 手术），每次只读写一个分片
  selftest                       → 用 1.5B 对比逐层结果与整模型结果（必须一致才可信）

    python surgery/layerwise.py selftest --model models/Qwen2.5-Coder-1.5B-Instruct
"""
from __future__ import annotations

import argparse
import gc
import json
import shutil
from pathlib import Path

import torch
from safetensors import safe_open
from safetensors.torch import save_file
from transformers import AutoConfig, AutoTokenizer
from transformers.models.qwen2.modeling_qwen2 import Qwen2DecoderLayer, Qwen2RMSNorm, Qwen2RotaryEmbedding


def rss_gb():
    try:
        import psutil
        return round(psutil.Process().memory_info().rss / 2**30, 2)
    except ImportError:
        return None


class Layerwise:
    def __init__(self, path, dtype=torch.bfloat16):
        self.path = Path(path)
        self.cfg = AutoConfig.from_pretrained(path)
        self.cfg._attn_implementation = "sdpa"
        self.tok = AutoTokenizer.from_pretrained(path)
        self.dtype = dtype
        idx = self.path / "model.safetensors.index.json"
        if idx.exists():
            self.wmap = json.loads(idx.read_text())["weight_map"]
        else:
            self.wmap = {k: "model.safetensors" for k in safe_open(self.path / "model.safetensors", "pt").keys()}
        self.embed = self._get("model.embed_tokens.weight")
        self.norm = Qwen2RMSNorm(self.cfg.hidden_size, eps=self.cfg.rms_norm_eps).to(dtype)
        self.norm.weight.data = self._get("model.norm.weight")
        self.lm_head = self.embed if "lm_head.weight" not in self.wmap else self._get("lm_head.weight")
        self.rope = Qwen2RotaryEmbedding(config=self.cfg)
        self.n_layers = self.cfg.num_hidden_layers
        self.peak_rss = rss_gb()

    def _get(self, key):
        with safe_open(self.path / self.wmap[key], "pt") as f:
            return f.get_tensor(key).to(self.dtype)

    def layer(self, i):
        L = Qwen2DecoderLayer(self.cfg, layer_idx=i).to(self.dtype)
        pre = f"model.layers.{i}."
        sd = {}
        files = {}
        for k, fn in self.wmap.items():
            if k.startswith(pre):
                files.setdefault(fn, []).append(k)
        for fn, keys in files.items():
            with safe_open(self.path / fn, "pt") as f:
                for k in keys:
                    sd[k[len(pre):]] = f.get_tensor(k).to(self.dtype)
        L.load_state_dict(sd, strict=True)
        return L.eval()

    def _run_layer(self, L, h, pos_emb):
        out = L(h, attention_mask=None, position_ids=None, position_embeddings=pos_emb)
        return out[0] if isinstance(out, tuple) else out

    @torch.no_grad()
    def forward(self, seqs, hooks=None, upto=None, keep_hidden=False):
        """seqs: list[list[int]]。逐层推进全部序列。hooks: {layer: fn(layer_module, i_seq, h_in)} 可注册前置钩子。
        返回每条序列最终隐状态（未过 norm），keep_hidden=True 时同时返回每层输出。"""
        hs = [self.embed[torch.tensor(s)].unsqueeze(0) for s in seqs]
        pes = [self.rope(h, torch.arange(h.shape[1]).unsqueeze(0)) for h in hs]
        per_layer = [] if keep_hidden else None
        last = self.n_layers if upto is None else upto
        for i in range(last):
            L = self.layer(i)
            handles = hooks(i, L) if hooks else []
            for j in range(len(hs)):
                self._cur_seq = j
                hs[j] = self._run_layer(L, hs[j], pes[j])
            for hd in handles:
                hd.remove()
            if keep_hidden:
                per_layer.append([h.clone() for h in hs])
            del L
            gc.collect()
            r = rss_gb()
            if r and (self.peak_rss is None or r > self.peak_rss):
                self.peak_rss = r
        return hs, per_layer

    @torch.no_grad()
    def forward_last(self, prompts):
        seqs = [self.tok(p, add_special_tokens=False).input_ids for p in prompts]
        hs, _ = self.forward(seqs)
        return torch.stack([(self.norm(h[0, -1:]) @ self.lm_head.T)[0].float() for h in hs])

    @torch.no_grad()
    def collect_mlp(self, seqs, layers=None):
        """返回 {layer: [每条序列的 down_proj 输入 [T, d_ff] fp32]}。"""
        want = set(range(self.n_layers)) if layers is None else set(layers)
        acts = {i: [] for i in want}

        def hooks(i, L):
            if i not in want:
                return []
            return [L.mlp.down_proj.register_forward_pre_hook(lambda m, inp, i=i: acts[i].append(inp[0][0].float()))]
        self.forward(seqs, hooks=hooks, upto=max(want) + 1)
        return acts


def scale_neurons_on_disk(src, dst, neurons, alpha):
    """neuron_scale 手术：逐分片读入→改 down_proj 列→写出。峰值内存 ≈ 单个分片大小（7B 约 3.9GB）。"""
    src, dst = Path(src), Path(dst)
    dst.mkdir(parents=True, exist_ok=True)
    by_layer = {}
    for l, j in neurons:
        by_layer.setdefault(l, []).append(j)
    for p in src.iterdir():
        if p.suffix != ".safetensors" and p.is_file():
            shutil.copy2(p, dst / p.name)
    touched = []
    for shard in sorted(src.glob("*.safetensors")):
        with safe_open(shard, "pt") as f:
            meta = f.metadata()
            tensors = {k: f.get_tensor(k) for k in f.keys()}
        for l, idx in by_layer.items():
            k = f"model.layers.{l}.mlp.down_proj.weight"
            if k in tensors:
                w = tensors[k]
                w[:, idx] = (w[:, idx].float() * alpha).to(w.dtype)
                touched.append(k)
        save_file(tensors, dst / shard.name, metadata=meta or {"format": "pt"})
        del tensors
        gc.collect()
        print("  wrote", shard.name, "rss", rss_gb(), "GB", flush=True)
    return touched


def selftest(path):
    from transformers import AutoModelForCausalLM
    lw = Layerwise(path, dtype=torch.float32)
    prompts = ["=SUMIFS(", "按部门汇总金额，公式写作 ="]
    a = lw.forward_last(prompts)
    full = AutoModelForCausalLM.from_pretrained(path, torch_dtype=torch.float32)
    tok = lw.tok
    b = torch.stack([full(tok(p, return_tensors="pt", add_special_tokens=False).input_ids).logits[0, -1].float() for p in prompts])
    diff = (a - b).abs().max().item()
    same_top = bool((a.argmax(-1) == b.argmax(-1)).all())
    print(json.dumps({"max_abs_diff": diff, "same_argmax": same_top, "peak_rss_gb": lw.peak_rss}))
    assert same_top and diff < 1e-2, "layerwise != full model"


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["selftest"])
    ap.add_argument("--model", required=True)
    a = ap.parse_args()
    selftest(a.model)
