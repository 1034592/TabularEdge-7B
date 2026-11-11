"""顽疾指标（手术前后、量化前后统一用这一套）。

  pair_acc[family]   P(target) > P(distractor) 的比例（sumifs / vlookup / xlookup）
  top1[family]       target 首 token 就是 argmax 的比例
  rep_rate           repeat 探针贪心生成的 4-gram 重复率（参考值：枚举月份/星期等正常输出也会有结构性重复）
  loop_frac          真·死循环：生成在预算内没结束，且末尾 token 序列以周期 p≤40 连续重复 ≥3 次
  unfinished         预算内没有自然结束（没生成 EOS）的比例
  ppl                回归文本困惑度（越低越好；手术不能让它明显上升）

    python surgery/measure.py --split heldout --out outputs/measure/base.json
"""
from __future__ import annotations

import argparse
import json
import math
from collections import Counter
from pathlib import Path

import torch

from common import load, read_jsonl, target_id

PROBES = Path(__file__).resolve().parent / "probes"


def first_logits(model, tok, prompt):
    ids = tok(prompt, return_tensors="pt", add_special_tokens=False).input_ids.to(model.device)
    with torch.no_grad():
        return model(ids).logits[0, -1].float()


def trace_metrics(model, tok, rows):
    ok = top = 0
    margins = []
    for r in rows:
        lg = first_logits(model, tok, r["prompt"])
        t, d = target_id(tok, r["target"]), target_id(tok, r["distractor"])
        ok += int(lg[t] > lg[d])
        top += int(lg.argmax().item() == t)
        margins.append((lg[t] - lg[d]).item())
    n = max(len(rows), 1)
    return {"pair_acc": round(ok / n, 3), "top1": round(top / n, 3), "margin": round(sum(margins) / n, 3), "n": len(rows)}


def ngram_stats(ids):
    g4 = [tuple(ids[i:i + 4]) for i in range(len(ids) - 3)]
    rep = 0.0 if not g4 else 1 - len(set(g4)) / len(g4)
    g8 = Counter(tuple(ids[i:i + 8]) for i in range(len(ids) - 7))
    return rep, (max(g8.values()) if g8 else 0) >= 3


def tail_loop(ids, max_p=40, reps=3):
    """末尾是否为周期 p 的重复（=死循环的指纹）。返回周期或 0。"""
    for p in range(1, max_p + 1):
        if len(ids) >= p * reps and all(ids[-k] == ids[-k - p] for k in range(1, p * (reps - 1) + 1)):
            return p
    return 0


def repeat_metrics(model, tok, rows, max_new=256):
    reps, loops, unfinished, samples = [], 0, 0, []
    eos = {tok.eos_token_id, tok.convert_tokens_to_ids("<|im_end|>"), tok.convert_tokens_to_ids("<|endoftext|>")}
    for r in rows:
        ids = tok(r["prompt"], return_tensors="pt", add_special_tokens=False).input_ids.to(model.device)
        with torch.no_grad():
            g = model.generate(ids, max_new_tokens=max_new, do_sample=False, pad_token_id=tok.eos_token_id)
        new = g[0, ids.shape[1]:].tolist()
        rep, _ = ngram_stats(new)
        done = any(t in eos for t in new)
        per = 0 if done else tail_loop(new)
        reps.append(rep)
        loops += per > 0
        unfinished += not done
        samples.append({"id": r["id"], "finished": done, "loop_period": per, "tail": tok.decode(new[-80:])})
    n = max(len(rows), 1)
    return {"rep_rate": round(sum(reps) / n, 4), "loop_frac": round(loops / n, 3), "unfinished": round(unfinished / n, 3),
            "n": len(rows), "samples": samples}


def ppl(model, tok, texts):
    nll = n = 0
    for t in texts:
        ids = tok(t, return_tensors="pt").input_ids.to(model.device)
        with torch.no_grad():
            loss = model(ids, labels=ids).loss.item()
        nll += loss * (ids.shape[1] - 1)
        n += ids.shape[1] - 1
    return round(math.exp(nll / n), 4)


def measure(model, tok, split="heldout", families=("sumifs", "vlookup", "xlookup", "repeat"), max_new=256):
    out = {}
    for fam in families:
        rows = read_jsonl(PROBES / f"{fam}.{split}.jsonl")
        out[fam] = repeat_metrics(model, tok, rows, max_new) if fam == "repeat" else trace_metrics(model, tok, rows)
        if fam == "sumifs":  # 按 SUMIF / SUMIFS 分开看，混淆通常是单侧的
            for v in ("SUMIF", "SUMIFS"):
                out[f"sumifs[{v}]"] = trace_metrics(model, tok, [r for r in rows if r["variant"] == v])
    out["ppl"] = ppl(model, tok, [r["text"] for r in read_jsonl(PROBES / "regress_texts.jsonl")])
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", default=None)
    ap.add_argument("--split", default="heldout")
    ap.add_argument("--families", nargs="+", default=["sumifs", "vlookup", "xlookup", "repeat"])
    ap.add_argument("--device", default=None)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    model, tok = load(a.model, device=a.device) if a.model else load(device=a.device)
    m = measure(model, tok, a.split, a.families)
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: {kk: vv for kk, vv in v.items() if kk != "samples"} if isinstance(v, dict) else v for k, v in m.items()}, ensure_ascii=False))


if __name__ == "__main__":
    main()
