"""把某个顽疾的 edit 切分探针转成 ROME rank_one 请求（heldout 切分只用于闸门，绝不进入编辑）。

    python surgery/probes/rome_from_probes.py --family xlookup --n 6 --out surgery/probes/rome_xlookup.jsonl

每个 subject（如“无此人”“N/A”）只取一条主提示，其余同 subject 的提示作为 paraphrases，用来平均 key。
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

HERE = Path(__file__).resolve().parent


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--family", required=True)
    ap.add_argument("--n", type=int, default=6, help="最多几条请求（按 subject 去重）")
    ap.add_argument("--out", required=True)
    ap.add_argument("--key", choices=["subject", "last"], default="subject",
                    help="last = 在提示末 token 处改写（因果追踪显示模型不使用 subject 线索时用这个）")
    a = ap.parse_args()
    rows = [json.loads(l) for l in (HERE / f"{a.family}.edit.jsonl").read_text(encoding="utf-8").splitlines() if l.strip()]
    if a.key == "last":  # 按语言分组：每组 1 条请求，key = 该组所有提示末 token 的平均
        groups = {}
        for r in rows:
            groups.setdefault("zh" if any("\u4e00" <= c <= "\u9fff" for c in r["prompt"]) else "en", []).append(r)
        reqs = [{"prompt": g[0]["prompt"], "subject": g[0]["prompt"][-1], "subject_at_end": True, "target": g[0]["target"],
                 "paraphrases": [r["prompt"] for r in g[1:]]} for g in groups.values()]
        Path(a.out).write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in reqs) + "\n", encoding="utf-8")
        print(len(reqs), "requests ->", a.out)
        return
    by_subj = {}
    for r in rows:
        by_subj.setdefault(r["subject"], []).append(r)
    reqs = []
    for subj, rs in list(by_subj.items())[:a.n]:
        main_, rest = rs[0], rs[1:4]
        # paraphrase 只取到 subject 结束处（mlp_keys 对 paraphrase 取末 token 的 key）
        paras = [p["prompt"][:p["prompt"].index(subj) + len(subj)] for p in rest]
        reqs.append({"prompt": main_["prompt"], "subject": subj, "target": main_["target"], "paraphrases": paras})
    Path(a.out).write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in reqs) + "\n", encoding="utf-8")
    print(len(reqs), "requests ->", a.out)


if __name__ == "__main__":
    main()
