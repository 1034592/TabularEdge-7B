"""自蒸馏：对 distill_prompts 的每条提示，用原版 7B（llama-server，greedy）生成回答，写成通用对话训练行。

    python synth/distill.py --prompts data/distill/prompts.jsonl --url http://127.0.0.1:8085 --out data/distill/distill.jsonl --workers 8
"""
from __future__ import annotations

import argparse
import json
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import requests


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--prompts", required=True)
    ap.add_argument("--url", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--max-tokens", type=int, default=512)
    a = ap.parse_args()
    rows = [json.loads(l) for l in open(a.prompts, encoding="utf-8")]

    def one(r):
        j = requests.post(a.url.rstrip("/") + "/v1/chat/completions", timeout=600, json={
            "messages": [{"role": "user", "content": r["prompt"]}], "temperature": 0, "max_tokens": a.max_tokens}).json()
        c = j["choices"][0]
        return {**r, "chat": True, "messages": [{"role": "user", "content": r["prompt"]}], "response": c["message"]["content"],
                "finish": c.get("finish_reason")}

    with ThreadPoolExecutor(a.workers) as ex:
        out = list(ex.map(one, rows))
    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        for r in out:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    print(len(out), "rows;", sum(r["finish"] == "length" for r in out), "truncated at max_tokens")


if __name__ == "__main__":
    main()
