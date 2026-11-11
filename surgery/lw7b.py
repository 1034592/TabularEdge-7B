"""7B 逐层手术工作台：测量 + ROME 秩一编辑，进程峰值内存控制在 16GB 以内（红线 16–20GB）。

不整体加载 7B：
  measure  所有探针一起逐层推进（每层只从磁盘读一次），得到 pair_acc（sumifs/vlookup/xlookup）与回归集 ppl。
           repeat 需要逐 token 生成，逐层模式下每个 token 都要读一遍 15GB 权重，不可行 → 7B 不测 repeat
           （1.5B 上真循环率已是 0，见 DEV_LOG）。
  rome     ① 层 0..L-1 逐层流式：得到各提示在层 L 的输入、以及协方差 C=E[kkᵀ]（按序列累加，不存全部 key）
           ② 只让层 L..N-1 常驻（L=21 时 7 层 fp32 ≈ 6.6GB），在末 token 处优化 MLP 输出偏移 δ（ROME 的 v*）
           ③ ΔW = δ (C⁻¹k*)ᵀ / (k*ᵀC⁻¹k*)，只改写含该层 down_proj 的一个分片；其余分片用硬链接（不复制 15GB）

    python surgery/lw7b.py measure --model models/Qwen2.5-Coder-7B-Instruct --out outputs/measure/base-7b.json
    python surgery/lw7b.py rome --model ... --layer 21 --request surgery/probes/rome_xlookup.jsonl --dst models/out/edited-7b/hf
"""
from __future__ import annotations

import argparse
import gc
import json
import math
import os
import shutil
import time
from pathlib import Path

import torch
import torch.nn.functional as F
from safetensors import safe_open
from safetensors.torch import save_file

from common import read_jsonl, target_id
from layerwise import Layerwise, rss_gb

HERE = Path(__file__).resolve().parent
PROBES = HERE / "probes"
torch.set_num_threads(os.cpu_count() or 8)


# ---------------------------------------------------------------- measure
@torch.no_grad()
def measure(path, split="heldout", families=("sumifs", "vlookup", "xlookup")):
    lw = Layerwise(path, dtype=torch.float32)
    tok = lw.tok
    rows = {f: read_jsonl(PROBES / f"{f}.{split}.jsonl") for f in families}
    regress = [r["text"] for r in read_jsonl(PROBES / "regress_texts.jsonl")]
    probe_seqs = [tok(r["prompt"], add_special_tokens=False).input_ids for f in families for r in rows[f]]
    reg_seqs = [tok(t).input_ids for t in regress]
    t0 = time.time()
    hs, _ = lw.forward(probe_seqs + reg_seqs)
    out, k = {}, 0
    for f in families:
        ok = top = 0
        margins = []
        sub = {"SUMIF": [0, 0], "SUMIFS": [0, 0]}
        for r in rows[f]:
            lg = (lw.norm(hs[k][0, -1:]) @ lw.lm_head.T)[0]
            k += 1
            t, d = target_id(tok, r["target"]), target_id(tok, r["distractor"])
            good = bool(lg[t] > lg[d])
            ok += good
            top += int(lg.argmax().item() == t)
            margins.append((lg[t] - lg[d]).item())
            if f == "sumifs":
                sub[r["variant"]][0] += good
                sub[r["variant"]][1] += 1
        n = max(len(rows[f]), 1)
        out[f] = {"pair_acc": round(ok / n, 3), "top1": round(top / n, 3), "margin": round(sum(margins) / n, 3), "n": len(rows[f])}
        if f == "sumifs":
            for v, (g, m) in sub.items():
                out[f"sumifs[{v}]"] = {"pair_acc": round(g / max(m, 1), 3), "n": m}
    nll = ntok = 0
    for s, h in zip(reg_seqs, hs[k:]):
        lg = lw.norm(h[0, :-1]) @ lw.lm_head.T
        nll += F.cross_entropy(lg, torch.tensor(s[1:]), reduction="sum").item()
        ntok += len(s) - 1
    out["ppl"] = round(math.exp(nll / ntok), 4)
    out["_meta"] = {"sec": round(time.time() - t0), "peak_rss_gb": lw.peak_rss, "model": str(path)}
    return out


# ---------------------------------------------------------------- rome
def _cov(lw, seqs_cov, L):
    """层 0..L 逐层推进协方差语料；在层 L 的 down_proj 输入上累加 C=Σ kkᵀ（逐序列累加，不保存 key）。"""
    d_ff = lw.cfg.intermediate_size
    C = torch.zeros(d_ff, d_ff)
    n = [0]

    def hooks(i, layer):
        if i != L:
            return []

        def acc(m, inp):
            k = inp[0][0].float()
            C.addmm_(k.T, k)
            n[0] += k.shape[0]
        return [layer.mlp.down_proj.register_forward_pre_hook(acc)]
    lw.forward(seqs_cov, hooks=hooks, upto=L + 1)
    return C, n[0]


def rome(path, L, requests, dst, steps=25, lr=0.5, kl_w=0.0625, corpus_n=300):
    lw = Layerwise(path, dtype=torch.float32)
    tok = lw.tok
    reqs = read_jsonl(requests)
    corpus = [r["text"] for r in read_jsonl(PROBES / "regress_texts.jsonl")]
    corpus += [r["text"] for r in read_jsonl(HERE.parent / "data" / "synth" / "train.jsonl")[:corpus_n]]
    corpus += [r["prompt"] for f in ("sumifs", "vlookup") for r in read_jsonl(PROBES / f"{f}.edit.jsonl")]
    seqs_cov = [tok(t, add_special_tokens=False).input_ids for t in corpus]
    groups = [[req["prompt"]] + req.get("paraphrases", []) for req in reqs]
    seqs_req = [tok(p, add_special_tokens=False).input_ids for g in groups for p in g]
    t0 = time.time()

    # ① 层 0..L-1 + 层 L 上的协方差。forward(upto=L+1) 返回的是层 L 的输出；我们需要层 L 的输入 → 跑 upto=L
    hs_in, _ = lw.forward(seqs_req, upto=L)
    C, ntok = _cov(lw, seqs_cov, L)
    C = C / ntok + 1e-4 * torch.eye(C.shape[0])
    print(f"cov: {ntok} tokens, d_ff={C.shape[0]}, {time.time() - t0:.0f}s, rss {rss_gb()}GB", flush=True)

    # ② 层 L..N-1 常驻
    resident = [lw.layer(i) for i in range(L, lw.n_layers)]
    for m in resident:
        for p in m.parameters():
            p.requires_grad_(False)
    print(f"resident layers {L}..{lw.n_layers - 1}, rss {rss_gb()}GB", flush=True)
    pes = [lw.rope(h, torch.arange(h.shape[1]).unsqueeze(0)) for h in hs_in]
    down = resident[0].mlp.down_proj
    infos = []
    k_off = 0
    for req, g in zip(reqs, groups):
        idx = list(range(k_off, k_off + len(g)))
        k_off += len(g)
        tgt = target_id(tok, req["target"])
        keys = []

        def run(i_seq, delta=None):
            h = hs_in[i_seq]
            cap = {}

            def pre(m, inp):
                cap["k"] = inp[0][0, -1].detach().clone()

            def post(m, inp, out):
                if delta is None:
                    return out
                out = out.clone()
                out[0, -1] = out[0, -1] + delta
                return out
            hk = [resident[0].mlp.down_proj.register_forward_pre_hook(pre), resident[0].mlp.register_forward_hook(post)]
            try:
                for layer in resident:
                    h = lw._run_layer(layer, h, pes[i_seq])
            finally:
                for x in hk:
                    x.remove()
            return (lw.norm(h[0, -1:]) @ lw.lm_head.T)[0].log_softmax(-1), cap["k"]

        with torch.no_grad():
            base_lp = []
            for i in idx:
                lp, k = run(i)
                base_lp.append(lp)
                keys.append(k)
        k_star = torch.stack(keys).mean(0)
        delta = torch.zeros(lw.cfg.hidden_size, requires_grad=True)
        opt = torch.optim.Adam([delta], lr=lr)
        for s in range(steps):
            loss = 0.0
            with torch.enable_grad():
                for j, i in enumerate(idx):
                    lp, _ = run(i, delta)
                    loss = loss - lp[tgt] + kl_w * F.kl_div(lp, base_lp[j], log_target=True, reduction="sum")
                loss = loss / len(idx) + 1e-3 * delta.norm() ** 2
                opt.zero_grad()
                loss.backward()
                opt.step()
            if s % 5 == 0 or s == steps - 1:
                print(f"  req {req['prompt'][:24]!r} step {s} loss {loss.item():.3f}", flush=True)
        with torch.no_grad():
            Cinv_k = torch.linalg.solve(C, k_star)
            dW = torch.outer(delta.detach(), Cinv_k) / (k_star @ Cinv_k)
            down.weight += dW
        infos.append({"prompt": req["prompt"], "n_prompts": len(idx), "dW_norm": dW.norm().item(),
                      "delta_norm": delta.norm().item()})
        print(" ", infos[-1], flush=True)

    # ③ 只改写含 down_proj 的分片；其它文件硬链接
    key = f"model.layers.{L}.mlp.down_proj.weight"
    new_w = down.weight.detach().to(torch.bfloat16).contiguous()
    del resident, hs_in, C
    gc.collect()
    src, dst = Path(path), Path(dst)
    dst.mkdir(parents=True, exist_ok=True)
    shard = lw.wmap[key]
    for p in src.iterdir():
        if not p.is_file():
            continue
        q = dst / p.name
        if q.exists():
            q.unlink()
        if p.name == shard:
            continue
        try:
            os.link(p, q)
        except OSError:
            shutil.copy2(p, q)
    with safe_open(src / shard, "pt") as f:
        meta = f.metadata()
        tensors = {k: f.get_tensor(k) for k in f.keys()}
    tensors[key] = new_w
    save_file(tensors, dst / shard, metadata=meta or {"format": "pt"})
    print(f"wrote {dst / shard} (rss {rss_gb()}GB)", flush=True)
    return {"layer": L, "requests": infos, "sec": round(time.time() - t0), "peak_rss_gb": max(lw.peak_rss or 0, rss_gb() or 0),
            "dst": str(dst)}


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["measure", "rome"])
    ap.add_argument("--model", required=True)
    ap.add_argument("--split", default="heldout")
    ap.add_argument("--out")
    ap.add_argument("--layer", type=int)
    ap.add_argument("--request")
    ap.add_argument("--dst")
    a = ap.parse_args()
    if a.cmd == "measure":
        m = measure(a.model, a.split)
    else:
        m = rome(a.model, a.layer, a.request, a.dst)
    if a.out:
        Path(a.out).parent.mkdir(parents=True, exist_ok=True)
        Path(a.out).write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(m, ensure_ascii=False))


if __name__ == "__main__":
    main()
