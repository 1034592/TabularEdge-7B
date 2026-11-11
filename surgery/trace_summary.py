"""汇总 causal_trace 输出：按层平均“subject 末 token 处”的间接效应，挑出 rank_one 手术的目标层。

    python surgery/trace_summary.py --dir outputs/trace/sumifs --method attrib --out outputs/trace/sumifs_summary.json
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    ap.add_argument("--method", default="attrib")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    files = sorted(Path(a.dir).glob(f"*.{a.method}.json"))
    agg, n = {}, 0
    for f in files:
        r = json.loads(f.read_text(encoding="utf-8"))
        s_last, last = r["span"][-1], len(r["tokens"]) - 1
        for kind, grid in r["ie"].items():
            a_ = agg.setdefault(kind, {"subject_last": [0.0] * len(grid), "last": [0.0] * len(grid)})
            for l, row in enumerate(grid):
                a_["subject_last"][l] += row[s_last] or 0.0
                a_["last"][l] += row[last] or 0.0
        n += 1
    for kind in agg:
        for pos in agg[kind]:
            agg[kind][pos] = [round(v / max(n, 1), 5) for v in agg[kind][pos]]
    mlp = agg.get("mlp", {}).get("subject_last", [])
    best = max(range(len(mlp)), key=lambda l: mlp[l]) if mlp else None
    rep = {"n_probes": n, "method": a.method, "per_layer": agg, "best_mlp_layer_subject_last": best}
    Path(a.out).write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    for kind, d in agg.items():
        top = sorted(range(len(d["subject_last"])), key=lambda l: -d["subject_last"][l])[:5]
        print(kind, "subject_last top layers:", [(l, d["subject_last"][l]) for l in top])
    print("best mlp layer:", best)


if __name__ == "__main__":
    main()
