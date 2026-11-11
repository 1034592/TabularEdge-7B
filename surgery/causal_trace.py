"""因果追踪：定位“哪一层、哪个位置、哪种组件”承载了表格/公式任务的关键计算。

方法（Meng et al. 2022 ROME 的 causal tracing，适配到公式任务）：
  clean     : 原始提示 → P(正确 token)
  corrupted : 给 subject（如“如果找不到返回默认值”“条件求和多条件”）的 embedding 加噪 → P 下降
  restore   : 在 corrupted 运行中，把第 L 层位置 i 的激活恢复为 clean 值 → P 回升多少 = 该状态的间接效应(IE)

两种模式：
  --method attrib : 归因修补（attribution patching, 一阶近似），1 次前向+1 次反向得到全部 (层×位置) 估计，
                    单位是 Δlog P（与 patch 模式的 ΔP 不同量纲，只用于排序）。
                    7B 在 CPU 上也是分钟级，用来初筛。
  --method patch  : 精确激活修补，按位置批量；只对 attrib 初筛出的 top 层做，用来确认。

输出 outputs/trace/<probe_id>.<kind>.json：{tokens, ie[layer][pos], p_clean, p_corrupt}

    python surgery/causal_trace.py --probes surgery/probes/formula_probes.jsonl --method attrib --kinds hidden mlp attn
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import torch

from common import _main, _replace, hooks, layers, load, module, read_jsonl, subject_span, target_id

NOISE_MULT = 3.0  # ROME 设定：噪声标准差 = 3 × embedding 标准差


def _noise_hook(span, std, seed=0):
    g = torch.Generator().manual_seed(seed)
    cache = {}

    def fn(mod, inp, out):
        if "n" not in cache:  # 同一次实验所有批次共享同一份噪声
            cache["n"] = (torch.randn(out.shape[-2:], generator=g) * std).to(out.dtype).to(out.device)
        out = out.clone()
        out[:, span] += cache["n"][span]
        if torch.is_grad_enabled():
            out.requires_grad_(True)  # 计算图从加噪后的 embedding 起建立，不需要对 embedding 权重求梯度
        return out
    return fn


def run_probe(model, tok, probe, kinds, method, top_layers=None):
    ids = tok(probe["prompt"], return_tensors="pt", add_special_tokens=False).input_ids.to(model.device)
    T, nL = ids.shape[1], len(layers(model))
    tgt = target_id(tok, probe["target"])
    dis = target_id(tok, probe["distractor"]) if probe.get("distractor") else None

    def objective(logits):  # logit 差（target − distractor）：目标 token 绝对概率很小时依然稳定
        lp = logits[0, -1].float().log_softmax(-1)
        return lp[tgt] - (lp[dis] if dis is not None else 0)
    span = subject_span(tok, probe["prompt"], probe["subject"])
    emb = model.model.embed_tokens
    std = NOISE_MULT * emb.weight.float().std().item()

    # clean 运行：缓存所有层/组件输出
    clean = {}
    hs = [module(model, l, k).register_forward_hook(
        lambda m, i, o, key=(l, k): clean.__setitem__(key, _main(o).detach()))
        for l in range(nL) for k in kinds]
    with hooks(hs), torch.no_grad():
        lg = model(ids).logits
        p_clean = lg[0, -1].float().softmax(-1)[tgt].item()
        obj_clean = objective(lg).item()

    res = {"id": probe["id"], "tokens": tok.convert_ids_to_tokens(ids[0]), "span": span,
           "p_clean": p_clean, "obj_clean": obj_clean, "target": probe["target"], "ie": {}}

    if method == "attrib":
        acts = {}

        def keep(m, i, o, key):
            t = _main(o)
            t.retain_grad()
            acts[key] = t
        hs = [emb.register_forward_hook(_noise_hook(span, std))]
        hs += [module(model, l, k).register_forward_hook(lambda m, i, o, key=(l, k): keep(m, i, o, key))
               for l in range(nL) for k in kinds]
        with hooks(hs), torch.enable_grad():
            lg = model(ids).logits
            obj = objective(lg)
            obj.backward()
        res["p_corrupt"] = lg[0, -1].float().softmax(-1)[tgt].item()
        res["obj_corrupt"] = obj.item()
        for k in kinds:
            res["ie"][k] = [((clean[(l, k)] - acts[(l, k)].detach()) * acts[(l, k)].grad)[0].float().sum(-1).tolist()
                            for l in range(nL)]
        return res

    # 精确修补：每层一次批量前向，batch 维 = 位置
    with hooks([emb.register_forward_hook(_noise_hook(span, std))]), torch.no_grad():
        res["p_corrupt"] = model(ids).logits[0, -1].float().softmax(-1)[tgt].item()
    for k in kinds:
        grid = []
        for l in range(nL):
            if top_layers and l not in top_layers:
                grid.append([None] * T)
                continue

            def patch(m, i, o, l=l):
                h = _main(o).clone()
                for b in range(T):
                    h[b, b] = clean[(l, k)][0, b]
                return _replace(o, h)
            hs = [emb.register_forward_hook(_noise_hook(span, std)), module(model, l, k).register_forward_hook(patch)]
            with hooks(hs), torch.no_grad():
                p = model(ids.repeat(T, 1)).logits[:, -1].float().softmax(-1)[:, tgt]
            grid.append((p - res["p_corrupt"]).tolist())
        res["ie"][k] = grid
    return res


def summarize(res, k=5):
    """返回 IE 最大的 (kind, layer, pos, token) —— 即候选“病灶”。"""
    cells = [(v, kind, l, p) for kind, g in res["ie"].items() for l, row in enumerate(g)
             for p, v in enumerate(row) if v is not None]
    return [{"kind": kd, "layer": l, "pos": p, "token": res["tokens"][p], "ie": round(v, 5)}
            for v, kd, l, p in sorted(cells, reverse=True)[:k]]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default=None)
    ap.add_argument("--probes", default="surgery/probes/formula_probes.jsonl")
    ap.add_argument("--kinds", nargs="+", default=["hidden", "mlp", "attn"])
    ap.add_argument("--method", choices=["attrib", "patch"], default="attrib")
    ap.add_argument("--layers", type=int, nargs="*", help="patch 模式只修补这些层（来自 attrib 初筛）")
    ap.add_argument("--out", default="outputs/trace")
    a = ap.parse_args()
    model, tok = load(a.model) if a.model else load()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    for pr in read_jsonl(a.probes):
        if pr.get("kind", "trace") != "trace":
            continue
        r = run_probe(model, tok, pr, a.kinds, a.method, set(a.layers or []))
        (out / f"{pr['id']}.{a.method}.json").write_text(json.dumps(r, ensure_ascii=False), encoding="utf-8")
        print(pr["id"], f"p_clean={r['p_clean']:.3f} p_corrupt={r['p_corrupt']:.3f} obj={r['obj_clean']:.2f}->{r.get('obj_corrupt', float('nan')):.2f}",
              json.dumps(summarize(r), ensure_ascii=False))


if __name__ == "__main__":
    main()
