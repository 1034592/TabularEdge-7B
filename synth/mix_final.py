"""收官淬火训练集 = v3 靶向补强（全量） + v2 回放（按反事实对整组抽样，防遗忘） + 自蒸馏通用对话（未截断，×2 加权）。
另输出污染核查：训练文本与人工测试集（v1/v2）完全相同或相似度 > 0.9 的题目清单（评测时报“剔除污染题”的成绩）。

    python synth/mix_final.py --v3 data/synth_v3/train.jsonl --v2 data/synth_v2/train.jsonl --distill data/distill/distill.jsonl \
        --replay 3000 --out data/final_mix.jsonl --contam outputs/eval/contamination.json
"""
from __future__ import annotations

import argparse
import difflib
import json
import random
from collections import defaultdict
from pathlib import Path


def rd(p):
    return [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--v3", required=True)
    ap.add_argument("--v2", required=True)
    ap.add_argument("--distill", required=True)
    ap.add_argument("--replay", type=int, default=3000)
    ap.add_argument("--distill-repeat", type=int, default=2)
    ap.add_argument("--raw-repeat", type=int, default=0, help="把蒸馏回答作为原始文本（不套模板）加入 N 次")
    ap.add_argument("--out", required=True)
    ap.add_argument("--contam", required=True)
    ap.add_argument("--human", nargs="+", default=["data/human/human_v1.jsonl", "data/human/human_v2.jsonl"])
    a = ap.parse_args()
    rng = random.Random(0)
    v3, v2 = rd(a.v3), rd(a.v2)
    groups = defaultdict(list)
    for i, r in enumerate(v2):
        groups[r["pair"] or f"solo{i}"].append(r)
    keys = sorted(groups)
    rng.shuffle(keys)
    replay = []
    for k in keys:
        if len(replay) >= a.replay:
            break
        replay += groups[k]
    dist = [r for r in rd(a.distill) if r["split"] == "train" and r.get("finish") != "length"]
    raw = [{"raw": r["response"], "id": f"raw{r['id']}"} for r in dist] * a.raw_repeat
    mix = v3 + replay + dist * a.distill_repeat + raw
    rng.shuffle(mix)
    with open(a.out, "w", encoding="utf-8") as f:
        for r in mix:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    print(f"mix: v3 {len(v3)} + replay {len(replay)} + distill {len(dist)}x{a.distill_repeat} + raw {len(raw)} = {len(mix)}")

    # 污染核查：人工题 vs 全部 TabularEdge 训练文本（v2 全量 + v3）
    train_txt = {r["text"] for r in v2 + v3 if "text" in r}
    out = {}
    for hp in a.human:
        hum = rd(hp)
        flagged = []
        for h in hum:
            if h["text"] in train_txt:
                flagged.append({"id": h["id"], "text": h["text"], "match": "exact"})
                continue
            sm = difflib.SequenceMatcher()
            sm.set_seq2(h["text"])  # seq2 的预处理只做一次
            best = (0.0, "")
            for t in train_txt:
                if abs(len(t) - len(h["text"])) > max(4, len(h["text"]) // 5):
                    continue
                sm.set_seq1(t)
                if sm.real_quick_ratio() > 0.9 and sm.quick_ratio() > 0.9:
                    r = sm.ratio()
                    if r > best[0]:
                        best = (r, t)
            if best[0] > 0.9:
                flagged.append({"id": h["id"], "text": h["text"], "match": f"{best[0]:.2f}", "train": best[1]})
        out[Path(hp).stem] = {"n": len(hum), "flagged": flagged}
        print(Path(hp).stem, "flagged", len(flagged), "/", len(hum))
    Path(a.contam).parent.mkdir(parents=True, exist_ok=True)
    Path(a.contam).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
