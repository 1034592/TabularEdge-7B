"""特异性检查：编辑是否把“以 = 结尾、答案本不是 XLOOKUP”的提示也带偏了。

  hit         首 token 等于正确函数首 token 的比例
  x_rate      首 token 是 XLOOKUP 首 token（"X"）的比例 —— 编辑后不应上升
  same_top1   与 --ref 模型首 token 一致的比例（只在传 --ref 时）

    python surgery/specificity.py --model A [--ref B] --out outputs/measure/spec.json
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

import torch

from common import load, read_jsonl, target_id

P = Path(__file__).resolve().parent / "probes" / "eq_specificity.jsonl"


@torch.no_grad()
def tops(model, tok, rows):
    out = []
    for r in rows:
        ids = tok(r["prompt"], return_tensors="pt", add_special_tokens=False).input_ids.to(model.device)
        out.append(model(ids).logits[0, -1].argmax().item())
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", required=True)
    ap.add_argument("--ref")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    rows = read_jsonl(P)
    m, tok = load(a.model)
    t = tops(m, tok, rows)
    x = target_id(tok, "XLOOKUP")
    res = {"n": len(rows), "hit": round(sum(ti == target_id(tok, r["target"]) for ti, r in zip(t, rows)) / len(rows), 3),
           "x_rate": round(sum(ti == x for ti in t) / len(rows), 3),
           "top1": [tok.decode([ti]) for ti in t]}
    if a.ref:
        del m
        mr, _ = load(a.ref)
        tr = tops(mr, tok, rows)
        res["same_top1"] = round(sum(p == q for p, q in zip(t, tr)) / len(rows), 3)
        res["ref_top1"] = [tok.decode([ti]) for ti in tr]
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps(res, ensure_ascii=False))


if __name__ == "__main__":
    main()
