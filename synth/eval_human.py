"""人工编写的真实口语测试集（data/human/*.jsonl，非模板生成）在 Demo 工作簿上的端到端评测。

评分只看“结果是否等价”：执行模型计划与金标计划，比较宽松摘要——
忽略默认输出文件名（*_处理后 / *_processed）与新建工作表的名字（模型可自行命名），其余单元格值/公式/格式/条件格式都要一致。
另报 expect_amb 样本的追问召回、非歧义样本的多余追问。

    python synth/eval_human.py --url http://127.0.0.1:8080 --tag tabularedge-7b-v2
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path

import requests
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "demo"))
from run_demo import make_files  # noqa: E402
sys.path.insert(0, str(ROOT / "data" / "human"))
from books import make_books  # noqa: E402
from tabularedge.executor import Executor, Workspace  # noqa: E402
from tabularedge.prompting import complete, messages  # noqa: E402
from tabularedge.protocol import profile_sheet, validate  # noqa: E402

OUT_RX = re.compile(r"_(处理后|processed)\.xlsx$")


def lenient_digest(root: Path, src_sheets: dict[str, list[str]]) -> str:
    h = hashlib.sha1()
    for p in sorted(root.rglob("*.xlsx")):
        if "_backup" in p.parts:
            continue
        rel = str(p.relative_to(root))
        name = "OUT" if OUT_RX.search(rel) else (rel if "/" not in rel else "SPLIT/" + p.name)  # 拆分文件只看文件名，不看目录名
        h.update(name.encode())
        wb = load_workbook(p)
        orig = src_sheets.get(rel) or src_sheets.get(OUT_RX.sub(".xlsx", rel)) or []
        for k, ws in enumerate(wb.worksheets):
            h.update((ws.title if ws.title in orig else f"NEW{k}").encode())
            for row in ws.iter_rows():
                for c in row:
                    v = (c.value.ref, c.value.text) if isinstance(c.value, ArrayFormula) else c.value
                    if ws.title not in orig and c.row == 1 and v in ("计数", "Count"):
                        v = "COUNT"  # 新表计数列表头的中英文视为等同
                    h.update(repr((v, c.number_format, c.font.color.rgb if c.font and c.font.color else None)).encode())
            for cf in ws.conditional_formatting:
                for r in cf.rules:
                    h.update(repr((str(cf.sqref), r.formula, r.dxf.font.color.rgb if r.dxf and r.dxf.font and r.dxf.font.color else None,
                                   r.dxf.fill.fgColor.rgb if r.dxf and r.dxf.fill else None)).encode())
    return h.hexdigest()


def gold_plan(item):
    steps = []
    for i, s in enumerate(item["steps"], 1):
        st = {"id": f"s{i}", **copy.deepcopy(s)}
        if st["op"] not in ("save", "fs_backup"):
            st["sheet"] = item["sheet"]
        steps.append(st)
    if not any(s["op"] == "save" for s in steps) and any(s["op"] not in ("split_to_files",) for s in steps):
        steps.append({"id": f"s{len(steps) + 1}", "op": "save", "args": {"path": item["file"].replace(".xlsx", "_处理后.xlsx")}})
    return {"version": "0.2", "lang": "zh", "inputs": [{"alias": "src", "path": item["file"], "sheet": item["sheet"]}], "steps": steps}


MAKE = {"v1": make_files, "v2": make_books}
RISKY = ("delete_rows", "dedupe")


def run(plan, workdir_files, books="v1"):
    d = Path(tempfile.mkdtemp(prefix="hum_"))
    try:
        MAKE[books](d)
        p = copy.deepcopy(plan)
        p.pop("ambiguities", None)
        Executor(Workspace(d)).run(p)
        return lenient_digest(d, workdir_files)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", default="http://127.0.0.1:8080")
    ap.add_argument("--data", default=str(ROOT / "data/human/human_v1.jsonl"))
    ap.add_argument("--tag", required=True)
    ap.add_argument("--books", choices=["v1", "v2"], default="v1")
    ap.add_argument("--out", default="outputs/eval")
    a = ap.parse_args()
    items = [json.loads(l) for l in open(a.data, encoding="utf-8") if l.strip()]
    d0 = Path(tempfile.mkdtemp(prefix="hum0_"))
    MAKE[a.books](d0)
    src_sheets = {p.name: load_workbook(p).sheetnames for p in d0.glob("*.xlsx")}
    res = []
    for it in items:
        wb = load_workbook(d0 / it["file"])
        prof = profile_sheet(wb[it["sheet"]])
        ptxt = prof.to_prompt() + "\n" + "\n".join(profile_sheet(x).to_prompt() for x in wb.worksheets if x.title != it["sheet"])
        rec = {"id": it["id"], "text": it["text"], "expect_amb": it["expect_amb"], "cat": it.get("cat"), "scored": it.get("steps") is not None}
        t0 = time.time()
        try:
            r = requests.post(a.url.rstrip("/") + "/v1/chat/completions", timeout=300, json={
                "messages": messages(it["file"], ptxt, it["text"]), "temperature": 0, "max_tokens": 900})
            raw = r.json()["choices"][0]["message"]["content"]
            rec["raw"] = raw
            obj = json.loads(raw)
            plan = complete({k: v for k, v in obj.items() if k not in ("version", "inputs")}, it["file"], it["sheet"])
            errs = validate(plan, prof)
            rec["valid"] = not errs
            rec["errors"] = errs[:3]
            rec["pred_amb"] = bool(plan.get("ambiguities"))
            risky = [s for s in plan["steps"] if s["op"] in RISKY or (s["op"] == "save" and s["args"].get("path") == it["file"])]
            rec["pred_risky_ops"] = len(risky)  # 这些步骤执行前一律经 UI 确认（执行器硬规则）
            rec["exec_eq"] = bool(rec["scored"] and not errs and run(plan, src_sheets, a.books) == run(gold_plan(it), src_sheets, a.books))
        except Exception as e:  # noqa: BLE001
            rec.update(valid=False, exec_eq=False, pred_amb=False, exc=f"{type(e).__name__}: {e}"[:200])
        rec["sec"] = round(time.time() - t0, 2)
        res.append(rec)
        print(it["id"], "OK " if rec["exec_eq"] else "XX ", "AMB" if rec["pred_amb"] else "   ", it["text"][:50], flush=True)
    shutil.rmtree(d0, ignore_errors=True)
    amb = [r for r in res if r["expect_amb"]]
    clr = [r for r in res if not r["expect_amb"]]
    sc = [r for r in res if r["scored"]]
    cats = sorted({r["cat"] for r in res if r["cat"]})
    by_cat = {c: {"n": sum(r["cat"] == c for r in res),
                  "exec_eq": round(sum(r["exec_eq"] for r in sc if r["cat"] == c) / max(1, sum(r["cat"] == c for r in sc)), 3),
                  "ask": round(sum(r["pred_amb"] for r in res if r["cat"] == c) / max(1, sum(r["cat"] == c for r in res)), 3)} for c in cats}
    # 歧义题“拦截”：模型主动追问，或计划中的破坏性步骤被执行器确认拦下（两者分开报告）
    m_extra = {"n_scored": len(sc), "exec_eq_scored": round(sum(r["exec_eq"] for r in sc) / max(len(sc), 1), 3),
               "amb_ask_only_items": sum(1 for r in amb if not r["scored"]),
               "amb_intercept_ask_or_confirm": round(sum(r["pred_amb"] or r.get("pred_risky_ops", 0) > 0 for r in amb) / max(len(amb), 1), 3),
               "risky_ops_total": sum(r.get("pred_risky_ops", 0) for r in res), "by_cat": by_cat}
    m = {"n": len(res), **m_extra, "exec_eq": round(sum(r["exec_eq"] for r in res) / len(res), 3),
         "exec_eq_clear": round(sum(r["exec_eq"] for r in clr) / max(len(clr), 1), 3),
         "valid": round(sum(r["valid"] for r in res) / len(res), 3),
         "amb_recall": round(sum(r["pred_amb"] for r in amb) / max(len(amb), 1), 3), "n_amb": len(amb),
         "amb_fp": round(sum(r["pred_amb"] for r in clr) / max(len(clr), 1), 3),
         "sec_avg": round(sum(r["sec"] for r in res) / len(res), 2)}
    Path(a.out).mkdir(parents=True, exist_ok=True)
    stem = Path(a.data).stem
    (Path(a.out) / f"{a.tag}.{stem}.jsonl").write_text("\n".join(json.dumps(r, ensure_ascii=False) for r in res), encoding="utf-8")
    (Path(a.out) / f"{a.tag}.{stem}.summary.json").write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    print(a.tag, stem, json.dumps(m, ensure_ascii=False))


if __name__ == "__main__":
    main()
