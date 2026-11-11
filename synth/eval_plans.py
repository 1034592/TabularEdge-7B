"""Action Plan 评测：任意后端（规则解析器 / llama-server 上的任意 GGUF）在合成测试集上的表现。

指标
  json_ok     输出是合法 JSON
  valid       通过完整 schema + 语义校验（列存在等）
  exact       steps 与金标逐字段一致（忽略 summary）
  exec_eq     在同一份随机 .xlsx 上分别执行金标与预测，结果摘要一致（最终用户真正关心的）
  amb_recall  金标有歧义 → 预测也提出了歧义
  amb_fp      金标无歧义 → 预测却提出歧义（多余追问）
  cf_both     反事实对的两条都 exec_eq（=模型确实随因果变量改变了输出）
  latency     单条平均耗时（秒），tok/s 来自 llama-server timings

    python synth/eval_plans.py --backend rule --data data/synth/test_ood.jsonl
    python synth/eval_plans.py --backend server --url http://127.0.0.1:8080 --prompt sft --data data/synth/test_ood.jsonl
"""
from __future__ import annotations

import argparse
import copy
import json
import random
import shutil
import sys
import tempfile
import time
from pathlib import Path

import requests
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "synth"))
from gen import Table, digest  # noqa: E402
from benchmarks import rule_baseline as rule_parser  # noqa: E402
from tabularedge.executor import Executor, Workspace  # noqa: E402
from tabularedge.llm_parser import SKELETON, SYSTEM as FEWSHOT_SYSTEM  # noqa: E402
from tabularedge.prompting import complete, messages  # noqa: E402
from tabularedge.protocol import profile_sheet, validate  # noqa: E402


def table_for(row):
    if row.get("gen") == 2:  # v2 样本：带维度表的 Table2
        from gen2 import Table2
        t = Table2(row["theme"], row["hlang"], random.Random(0))
    else:
        t = Table(row["theme"], row["hlang"], random.Random(0))
    t.file = row["file"]
    return t


def run_exec(row, plan):
    d = Path(tempfile.mkdtemp(prefix="eval_"))
    try:
        t = table_for(row)
        t.make_xlsx(d / t.file, random.Random(row["seed"]))
        p = copy.deepcopy(plan)
        p.pop("ambiguities", None)
        Executor(Workspace(d)).run(p)
        return digest(d)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def profile_of(row):
    d = Path(tempfile.mkdtemp(prefix="prof_"))
    try:
        t = table_for(row)
        t.make_xlsx(d / t.file, random.Random(row["seed"]))
        return profile_sheet(load_workbook(d / t.file)[t.sheet])
    finally:
        shutil.rmtree(d, ignore_errors=True)


def fewshot_msgs(row):
    """未微调基座用：长系统提示（op 速查）+ 一个示例。"""
    demo_prof = "[sheet=流水 rows=5] columns: 日期:text, 部门:text, 金额:num, 备注:text"
    from tabularedge.protocol import TableProfile
    p = TableProfile("流水", ["日期", "部门", "金额", "备注"], {"日期": "text", "部门": "text", "金额": "num", "备注": "text"}, 5)
    demo = rule_parser.parse("把带括号的备注删了，金额负数的标红，按部门汇总新表", p, "ledger.xlsx")
    demo.pop("inputs")
    return [{"role": "system", "content": FEWSHOT_SYSTEM},
            {"role": "user", "content": f"文件：ledger.xlsx\n{demo_prof}\n指令：把带括号的备注删了，金额负数的标红，按部门汇总新表"},
            {"role": "assistant", "content": json.dumps(demo, ensure_ascii=False, separators=(",", ":"))},
            {"role": "user", "content": f"文件：{row['file']}\n{row['profile']}\n指令：{row['text']}"}]


def predict_server(row, url, prompt, grammar, max_tokens):
    msgs = messages(row["file"], row["profile"], row["text"]) if prompt == "sft" else fewshot_msgs(row)
    body = {"messages": msgs, "temperature": 0, "max_tokens": max_tokens, "cache_prompt": True}
    if grammar:
        body["response_format"] = {"type": "json_schema", "json_schema": {"name": "plan", "schema": SKELETON}}
    r = requests.post(url.rstrip("/") + "/v1/chat/completions", json=body, timeout=600)
    r.raise_for_status()
    j = r.json()
    return j["choices"][0]["message"]["content"], j.get("timings", {})


def canon(steps):
    return json.dumps([{k: s[k] for k in sorted(s) if k not in ("note",)} for s in steps], ensure_ascii=False, sort_keys=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--backend", choices=["rule", "server"], required=True)
    ap.add_argument("--url", default="http://127.0.0.1:8080")
    ap.add_argument("--prompt", choices=["sft", "fewshot"], default="sft")
    ap.add_argument("--grammar", action="store_true", help="解码期加骨架 JSON Schema 约束")
    ap.add_argument("--data", nargs="+", default=["data/synth/test_id.jsonl", "data/synth/test_ood.jsonl"])
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--max-tokens", type=int, default=900)
    ap.add_argument("--tag", default="")
    ap.add_argument("--out", default="outputs/eval")
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    summary = {}
    for path in a.data:
        rows = [json.loads(l) for l in open(path, encoding="utf-8")]
        if a.limit:
            rows = rows[:a.limit]
        res = []
        for row in rows:
            t0 = time.time()
            rec = {"text": row["text"], "pair": row.get("pair"), "gold_amb": bool(row["plan"].get("ambiguities")),
                   "amb_kind": row.get("amb_kind"), "noise": row.get("noise") or [],
                   "risky": any(st["op"] in ("delete_rows", "dedupe") for st in row["plan"]["steps"])}
            try:
                if a.backend == "rule":
                    plan = rule_parser.parse(row["text"], profile_of(row), row["file"])
                    raw, timings = json.dumps(plan, ensure_ascii=False), {}
                else:
                    raw, timings = predict_server(row, a.url, a.prompt, a.grammar, a.max_tokens)
                rec["raw"] = raw
                rec["tok_s"] = timings.get("predicted_per_second")
                rec["n_tok"] = timings.get("predicted_n")
                obj = json.loads(raw)
                rec["json_ok"] = True
                plan = complete({k: v for k, v in obj.items() if k not in ("version", "inputs")}, row["file"], profile_of(row).sheet)
                errs = validate(plan, profile_of(row))
                rec["valid"] = not errs
                rec["errors"] = errs[:3]
                rec["exact"] = rec["valid"] and canon(plan["steps"]) == canon(row["plan"]["steps"])
                rec["pred_amb"] = bool(plan.get("ambiguities"))
                rec["exec_eq"] = bool(rec["valid"] and run_exec(row, plan) == run_exec(row, row["plan"]))
            except Exception as e:  # noqa: BLE001
                rec.setdefault("json_ok", False)
                rec.setdefault("valid", False)
                rec.setdefault("exact", False)
                rec.setdefault("exec_eq", False)
                rec.setdefault("pred_amb", False)
                rec["exc"] = f"{type(e).__name__}: {e}"[:200]
            rec["sec"] = round(time.time() - t0, 3)
            res.append(rec)
        n = len(res)
        pairs = {}
        for r in res:
            if r["pair"]:
                pairs.setdefault(r["pair"], []).append(r["exec_eq"])
        full = [v for v in pairs.values() if len(v) == 2]
        gold_amb = [r for r in res if r["gold_amb"]]
        no_amb = [r for r in res if not r["gold_amb"]]
        toks = [r["tok_s"] for r in res if r.get("tok_s")]
        m = {"n": n, **{k: round(sum(bool(r.get(k)) for r in res) / n, 3) for k in ("json_ok", "valid", "exact", "exec_eq")},
             "amb_recall": round(sum(r["pred_amb"] for r in gold_amb) / max(len(gold_amb), 1), 3), "n_amb": len(gold_amb),
             "amb_fp": round(sum(r["pred_amb"] for r in no_amb) / max(len(no_amb), 1), 3),
             "cf_both": round(sum(all(v) for v in full) / max(len(full), 1), 3), "n_pairs": len(full),
             "sec_avg": round(sum(r["sec"] for r in res) / n, 2),
             "tok_s": round(sum(toks) / len(toks), 1) if toks else None}
        # 分组：按歧义类型的召回、按扰动族的 exec_eq、含破坏性操作样本的 exec_eq
        kinds = sorted({r["amb_kind"] for r in gold_amb if r.get("amb_kind")})
        if kinds:
            m["amb_recall_by_kind"] = {k: round(sum(r["pred_amb"] for r in gold_amb if r["amb_kind"] == k) /
                                                max(1, sum(r["amb_kind"] == k for r in gold_amb)), 3) for k in kinds}
        fams = sorted({f for r in res for f in r["noise"]})
        if fams:
            m["exec_eq_by_noise"] = {f: round(sum(r["exec_eq"] for r in res if f in r["noise"]) /
                                              max(1, sum(f in r["noise"] for r in res)), 3) for f in fams}
            clean = [r for r in res if not r["noise"]]
            m["exec_eq_noised"] = round(sum(r["exec_eq"] for r in res if r["noise"]) / max(1, n - len(clean)), 3)
        risky = [r for r in res if r["risky"] and not r["gold_amb"]]
        if risky:
            m["exec_eq_risky"] = round(sum(r["exec_eq"] for r in risky) / len(risky), 3)
        name = Path(path).stem
        summary[name] = m
        tag = a.tag or a.backend
        (out / f"{tag}.{name}.jsonl").write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in res), encoding="utf-8")
        print(tag, name, json.dumps(m, ensure_ascii=False), flush=True)
    (out / f"{a.tag or a.backend}.summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
