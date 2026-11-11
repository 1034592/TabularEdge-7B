"""权重手术：只做能“烘焙”进权重的编辑，保证最终能导出单文件 GGUF、在 llama.cpp 中无需任何运行时钩子。

为什么限定为可烘焙编辑：
- 推理期 hook（ITI 激活偏移、ReFT 表征干预）在 llama.cpp 中没有对应机制；
  Qwen2 的 o_proj/down_proj 无 bias，常量残差偏移也无法直接写进 GGUF。
- 下面三种编辑都等价于修改已有张量的数值，转换 GGUF 后效果不变：
  1. neuron_scale : down_proj[:, j] *= α        ≡ 神经元 j 的激活乘 α（repetition_neurons.py 的结论）
  2. rank_one     : ROME 闭式秩一更新 down_proj += Δ，改写某层 MLP 的 “key→value” 关联（如 SUMIF↔SUMIFS 参数顺序混淆）
  3. lora_merge   : 外部训练的 LoRA 合并（因果定位后只在病灶层训练 LoRA，见 ARCHITECTURE.md §6.4）

每次编辑必须过 gate：目标探针改善 且 回归集困惑度相对上升 ≤ --max-ppl-regress（默认 1%），否则不保存。
产出：outputs/edits/<name>/ (safetensors + manifest.json)，再用 llama.cpp convert_hf_to_gguf.py 转换。

    python surgery/edit.py neuron_scale --neurons outputs/repetition_neurons.json --top 20 --alpha 0.3 --name rep20
    python surgery/edit.py rank_one --layer 18 --request surgery/probes/rome_requests.jsonl --name sumifs_order
"""
from __future__ import annotations

import argparse
import json
import math
import time
from pathlib import Path

import torch

from common import hooks, layers, load, read_jsonl, subject_span, target_id


# ---------------------------------------------------------------- gate
from measure import measure  # noqa: E402  统一指标，手术前后、量化前后同一套


def gate(before, after, family, max_ppl_regress, max_other_drop=0.02):
    """通过条件：目标顽疾在 heldout 上改善；困惑度上升 ≤ 阈值；其它顽疾的 pair_acc 不掉超过 2 个点。"""
    ok_ppl = after["ppl"] / before["ppl"] - 1 <= max_ppl_regress
    if family == "repeat":
        improved = (after["repeat"]["loop_frac"], after["repeat"]["rep_rate"]) < (before["repeat"]["loop_frac"], before["repeat"]["rep_rate"])
    else:
        improved = after[family]["pair_acc"] > before[family]["pair_acc"]
    others = [f for f, v in before.items() if f != family and f in after and isinstance(v, dict) and "pair_acc" in v
              and "[" not in f]  # 所有在场的顽疾族（叠加手术时每一刀都要保护之前所有族）
    no_harm = all(after[f]["pair_acc"] >= before[f]["pair_acc"] - max_other_drop for f in others)
    return {"improved": improved, "ppl_ok": ok_ppl, "no_collateral": no_harm, "passed": improved and ok_ppl and no_harm}


# ---------------------------------------------------------------- edits
@torch.no_grad()
def neuron_scale(model, neurons, alpha):
    for l, j in neurons:
        W = layers(model)[l].mlp.down_proj.weight  # [d_model, d_ff]
        W[:, j] *= alpha
    return {"neurons": neurons, "alpha": alpha}


def mlp_keys(model, tok, layer, texts, pos_fn):
    """收集 down_proj 输入（ROME 中的 key k）。"""
    ks = []
    h = layers(model)[layer].mlp.down_proj.register_forward_pre_hook(lambda m, inp: ks.append(inp[0][0].float()))
    with hooks([h]), torch.no_grad():
        for t in texts:
            ids = tok(t, return_tensors="pt", add_special_tokens=False).input_ids.to(model.device)
            model(ids)
    return [k[pos_fn(i)] if pos_fn else k for i, k in enumerate(ks)]


def key_covariance(model, tok, layer, corpus, max_tokens=30000):
    """C = E[k kᵀ]（ROME 的二阶矩）。d_ff=18944(7B) 时 C 为 18944² fp32 ≈ 1.4GB，32GB 内存可承受。"""
    d = layers(model)[layer].mlp.down_proj.in_features
    dev = model.device  # 大显存卡上 C 放 GPU（7B: 18944² fp32 ≈ 1.4GB）
    C, n = torch.zeros(d, d, device=dev), 0
    for k in mlp_keys(model, tok, layer, corpus, None):
        k = k.to(dev)
        C += k.T @ k
        n += k.shape[0]
        if n >= max_tokens:
            break
    return C / n + 1e-4 * torch.eye(d, device=dev)  # 轻微岭正则保证可逆


@torch.enable_grad()  # rank_one 外层是 no_grad，这里需要重新打开
def optimize_value(model, tok, layer, req, steps=25, lr=0.5, kl_w=0.0625):
    """求 v*：使 subject 末 token 处 MLP 输出 = v* 时，模型对 prompt 输出 target。含 KL 约束防止破坏其它预测。"""
    ids = tok(req["prompt"], return_tensors="pt", add_special_tokens=False).input_ids.to(model.device)
    pos = ids.shape[1] - 1 if req.get("subject_at_end") else subject_span(tok, req["prompt"], req["subject"])[-1]
    tgt = target_id(tok, req["target"])
    mlp = layers(model)[layer].mlp
    base = {}
    def grab(m, i, o):  # 注意：forward hook 返回非 None 会替换模块输出 → 这里必须显式返回 None
        base.setdefault("v", o[0, pos].detach().float())
    h = mlp.register_forward_hook(grab)
    with hooks([h]), torch.no_grad():
        p0 = model(ids).logits[0, -1].float().log_softmax(-1)
    delta = torch.zeros_like(base["v"], requires_grad=True)

    def add(m, i, o):
        o = o.clone()
        o[0, pos] = o[0, pos] + delta.to(o.dtype)
        return o
    opt = torch.optim.Adam([delta], lr=lr)
    for _ in range(steps):
        with hooks([mlp.register_forward_hook(add)]):
            lp = model(ids).logits[0, -1].float().log_softmax(-1)
        loss = -lp[tgt] + kl_w * torch.nn.functional.kl_div(lp, p0, log_target=True, reduction="sum") \
            + 1e-3 * delta.norm() ** 2
        opt.zero_grad()
        loss.backward()
        opt.step()
    return base["v"] + delta.detach(), pos


@torch.no_grad()
def rank_one(model, tok, layer, req, corpus):
    v_star, pos = optimize_value(model, tok, layer, req)
    k_star = torch.stack(mlp_keys(model, tok, layer, [req["prompt"]] + req.get("paraphrases", []),
                                  lambda i: pos if i == 0 else -1)).mean(0).cpu()
    C = key_covariance(model, tok, layer, corpus)
    k_star = k_star.to(C.device)
    Cinv_k = torch.linalg.solve(C, k_star)
    W = layers(model)[layer].mlp.down_proj.weight
    resid = v_star.to(C.device) - W.float().to(C.device) @ k_star
    dW = torch.outer(resid, Cinv_k) / (k_star @ Cinv_k)
    W += dW.to(W.dtype).to(W.device)
    return {"layer": layer, "request": req, "dW_norm": dW.norm().item()}


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("kind", choices=["neuron_scale", "rank_one"])
    ap.add_argument("--model", default=None)
    ap.add_argument("--name", required=True)
    ap.add_argument("--family", required=True, help="目标顽疾：repeat / sumifs / vlookup / xlookup")
    ap.add_argument("--max-ppl-regress", type=float, default=0.01)
    ap.add_argument("--neurons")
    ap.add_argument("--top", type=int, default=20)
    ap.add_argument("--alpha", type=float, default=0.3)
    ap.add_argument("--layer", type=int)
    ap.add_argument("--request")
    ap.add_argument("--before", help="复用已测好的 heldout 基线 JSON，省一次测量")
    ap.add_argument("--save", action="store_true", help="通过闸门后保存权重")
    ap.add_argument("--families", nargs="+", default=["sumifs", "vlookup", "xlookup", "repeat"],
                    help="闸门测量的顽疾（repeat 需要生成，最慢）")
    a = ap.parse_args()
    model, tok = load(a.model) if a.model else load()
    before = json.loads(Path(a.before).read_text(encoding="utf-8")) if a.before else measure(model, tok, "heldout", tuple(a.families))
    t0 = time.time()
    if a.kind == "neuron_scale":
        ns = [(n["layer"], n["idx"]) for n in json.loads(Path(a.neurons).read_text(encoding="utf-8"))["neurons"][:a.top]]
        info = neuron_scale(model, ns, a.alpha)
    else:
        here = Path(__file__).resolve().parent
        corpus = [r["text"] for r in read_jsonl(here / "probes" / "regress_texts.jsonl")]
        # 协方差 C=E[kkᵀ] 需要足够多 token（d_ff=8960/18944）：再加训练集指令文本与其它顽疾的 edit 提示
        corpus += [r["text"] for r in read_jsonl(here.parent / "data" / "synth" / "train.jsonl")[:300]]
        corpus += [r["prompt"] for f in ("sumifs", "vlookup") for r in read_jsonl(here / "probes" / f"{f}.edit.jsonl")]
        info = [rank_one(model, tok, a.layer, r, corpus) for r in read_jsonl(a.request)]
    edit_s = time.time() - t0
    after = measure(model, tok, "heldout", tuple(a.families))
    g = gate(before, after, a.family, a.max_ppl_regress)
    manifest = {"name": a.name, "kind": a.kind, "family": a.family, "info": info, "edit_sec": round(edit_s, 1),
                "before": before, "after": after, "gate": g, "time": time.strftime("%Y-%m-%d %H:%M:%S")}
    out = Path("outputs/edits") / a.name
    out.mkdir(parents=True, exist_ok=True)
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    brief = lambda m: {k: (v.get("pair_acc", v.get("rep_rate")) if isinstance(v, dict) else v) for k, v in m.items()}  # noqa: E731
    print(json.dumps({"before": brief(before), "after": brief(after), "gate": g}, ensure_ascii=False))
    if g["passed"] and a.save:
        # 权重一律放 models\out（不进仓库目录）；bf16 存盘，CPU 上 fp32 编辑的精度对 bf16 足够
        import os
        hf = Path(os.environ.get("TABULAREDGE_OUT", r"models\out" if os.name == "nt" else "models/out")) / a.name / "hf"
        model.to(torch.bfloat16).save_pretrained(hf, safe_serialization=True)
        tok.save_pretrained(hf)
        print("saved ->", hf)
    elif not g["passed"]:
        print("GATE FAILED: weights not saved")


if __name__ == "__main__":
    main()
