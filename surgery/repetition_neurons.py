"""定位“复读/死循环”神经元（Hiraoka & Inui, NAACL 2025, arXiv:2410.13497 的方法，适配到公式生成）。

目标症状：生成嵌套 IF / 长 SUMIFS / LAMBDA 递归时陷入重复片段，例如
    =IF(A2>90,"A",IF(A2>90,"A",IF(A2>90,"A",...
步骤
 1. 采集：对 repeat 类探针贪心生成，检测重复 n-gram，切分为“重复前”与“重复中”两段；
    另加合成对照（随机 token 序列重复两遍）作为与任务无关的基线。
 2. 打分：MLP 中间激活（down_proj 的输入，维度 = intermediate_size）在两段上的均值差 → 每层每神经元一个分数。
 3. 验证：把 top-k 神经元激活乘以 α（0=消融），测 (a) 重复率下降 (b) 正常探针准确率/困惑度不退化。
    只有通过 (b) 的神经元才进入 edit.py 的手术候选。

    python surgery/repetition_neurons.py --probes surgery/probes/formula_probes.jsonl --topk 50
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import torch

from common import hooks, layers, load, read_jsonl


def repeat_onset(ids: list[int], n: int = 6) -> int | None:
    """首次出现“已出现过的 n-gram”的位置；None = 无重复。"""
    seen = {}
    for i in range(len(ids) - n + 1):
        g = tuple(ids[i:i + n])
        if g in seen:
            return i
        seen[g] = i
    return None


def repetition_rate(ids: list[int], n: int = 4) -> float:
    grams = [tuple(ids[i:i + n]) for i in range(len(ids) - n + 1)]
    return 0.0 if not grams else 1 - len(set(grams)) / len(grams)


def collect_acts(model, ids):
    acts = []
    hs = [L.mlp.down_proj.register_forward_pre_hook(lambda m, inp: acts.append(inp[0][0].float().cpu()))
          for L in layers(model)]
    with hooks(hs), torch.no_grad():
        model(ids)
    return torch.stack(acts)  # [layers, T, d_ff]


def score(model, tok, probes, max_new=160, synthetic=8, seed=0):
    dev = model.device
    diffs, used = [], []
    for pr in probes:
        ids = tok(pr["prompt"], return_tensors="pt", add_special_tokens=False).input_ids.to(dev)
        with torch.no_grad():
            gen = model.generate(ids, max_new_tokens=max_new, do_sample=False, pad_token_id=tok.eos_token_id)
        new = gen[0, ids.shape[1]:].tolist()
        on = repeat_onset(new)
        if on is None or on < 4:
            continue
        a = collect_acts(model, gen)
        P = ids.shape[1]
        diffs.append(a[:, P + on:].mean(1) - a[:, P:P + on].mean(1))
        used.append({"id": pr["id"], "onset": on, "rep_rate": repetition_rate(new)})
    g = torch.Generator().manual_seed(seed)
    for _ in range(synthetic):  # 合成对照：随机序列重复一遍
        r = torch.randint(1000, tok.vocab_size - 1000, (1, 40), generator=g)
        a = collect_acts(model, torch.cat([r, r], 1).to(dev))
        diffs.append(a[:, 40:].mean(1) - a[:, :40].mean(1))
    return torch.stack(diffs).mean(0), used  # [layers, d_ff]


def scaled(model, neurons, alpha):
    """推理期把指定神经元激活乘 α（验证用）；验证通过后 edit.py 会把同样的缩放烘焙进 down_proj 权重。"""
    by_layer = {}
    for l, j in neurons:
        by_layer.setdefault(l, []).append(j)

    def mk(idx):
        def pre(m, inp):
            x = inp[0].clone()
            x[..., idx] *= alpha
            return (x,)
        return pre
    return hooks([layers(model)[l].mlp.down_proj.register_forward_pre_hook(mk(idx)) for l, idx in by_layer.items()])


def eval_rep(model, tok, probes, max_new=160):
    rates = []
    for pr in probes:
        ids = tok(pr["prompt"], return_tensors="pt", add_special_tokens=False).input_ids.to(model.device)
        with torch.no_grad():
            gen = model.generate(ids, max_new_tokens=max_new, do_sample=False, pad_token_id=tok.eos_token_id)
        rates.append(repetition_rate(gen[0, ids.shape[1]:].tolist()))
    return sum(rates) / max(len(rates), 1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default=None)
    ap.add_argument("--probes", default="surgery/probes/repeat.edit.jsonl")
    ap.add_argument("--topk", type=int, default=50)
    ap.add_argument("--alphas", type=float, nargs="+", default=[0.0, 0.5])
    ap.add_argument("--out", default="outputs/repetition_neurons.json")
    a = ap.parse_args()
    model, tok = load(a.model) if a.model else load()
    probes = read_jsonl(a.probes)
    rep = [p for p in probes if p.get("kind") == "repeat"]
    s, used = score(model, tok, rep)
    flat = s.flatten().topk(a.topk)
    d = s.shape[1]
    neurons = [(int(i) // d, int(i) % d) for i in flat.indices]
    report = {"used": used, "neurons": [{"layer": l, "idx": j, "score": round(float(v), 4)}
                                        for (l, j), v in zip(neurons, flat.values)],
              "baseline_rep": eval_rep(model, tok, rep), "ablations": []}
    for al in a.alphas:
        with scaled(model, neurons, al):
            report["ablations"].append({"alpha": al, "rep": eval_rep(model, tok, rep)})
        print(report["ablations"][-1])
    # 注意：这里还缺“正常任务不退化”验证，由 edit.py 的 gate 统一执行，避免两处标准不一致
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    print("baseline", report["baseline_rep"], "->", a.out)


if __name__ == "__main__":
    main()
