"""通用漂移困惑度：在留出的自蒸馏对话（原版 7B 自己的回答，未截断）上，算模型对回答 token 的困惑度。
越接近原版 = 通用能力保持得越好。（8 条 regress_texts 的 ppl 另由 measure.py 给出）

    python surgery/ppl_chat.py --model <hf> --data data/distill/distill.jsonl --out outputs/measure/pplchat-<tag>.json
"""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path

import torch
import torch.nn.functional as F

from common import load


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", required=True)
    ap.add_argument("--data", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    rows = [json.loads(l) for l in open(a.data, encoding="utf-8")]
    rows = [r for r in rows if r["split"] == "heldout" and r.get("finish") != "length"]
    model, tok = load(a.model)
    nll = n = 0
    with torch.no_grad():
        for r in rows:
            p = tok.apply_chat_template(r["messages"], add_generation_prompt=True, tokenize=True)
            p = p if isinstance(p, list) else p["input_ids"]
            y = tok(r["response"] + "<|im_end|>", add_special_tokens=False).input_ids
            ids = torch.tensor([p + y], device=model.device)
            lg = model(ids).logits[0, len(p) - 1:-1].float()
            nll += F.cross_entropy(lg, ids[0, len(p):], reduction="sum").item()
            n += len(y)
    res = {"n_rows": len(rows), "n_tokens": n, "ppl": round(math.exp(nll / n), 4), "model": a.model}
    Path(a.out).write_text(json.dumps(res, indent=1), encoding="utf-8")
    print(json.dumps(res))


if __name__ == "__main__":
    main()
