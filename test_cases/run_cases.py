"""10 个实测场景的自动化验收：金标计划 vs 本地模型，逐步执行并比较“算出来的结果”。

  python test_cases/run_cases.py --mode gold                       # 只跑金标（验证执行器 + 公式求值）
  python test_cases/run_cases.py --mode model --url http://127.0.0.1:18080 --out outputs/cases

比较口径（宽松，只看结果）：
  - 主表与新表：每列（去掉表头）的求值结果作为一个向量，比较列向量多重集合 → 不在乎模型给新列/新表起的名字和列顺序
  - 条件格式：用求值器逐格计算规则，比较“哪些格子被标了什么颜色”
  - 拆分文件：文件名集合 + 每个文件的列向量
  - 高危步骤（删行 / 去重 / 覆盖原文件）自动确认，并记录确认次数；歧义按默认方案执行
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import tempfile
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from openpyxl import load_workbook  # noqa: E402

from client.core import RefExecutor, prompt_context  # noqa: E402
from client.describe import is_risky  # noqa: E402
from tabularedge import llm_parser  # noqa: E402
from tabularedge.executor import Workspace  # noqa: E402
from tabularedge.formula_eval import Evaluator, Unsupported, XLError  # noqa: E402

SAVE = lambda p: {"op": "save", "args": {"path": p}}  # noqa: E731
LK = lambda name, ret, nf: {"op": "lookup_column", "args": {"name": name, "key": "工号", "from_sheet": "花名册", "from_key": "工号",  # noqa: E731
                                                           "return": ret, "if_not_found": nf}}
ADD = lambda n, f: {"op": "add_column", "args": {"name": n, "formula": f}}  # noqa: E731
CF = lambda where, st, to="cell": {"op": "conditional_format", "args": {"where": where, "style": st, "apply_to": to}}  # noqa: E731
GA = lambda by, aggs, out: {"op": "group_aggregate", "args": {"by": by, "aggs": [{"col": c, "fn": f} for c, f in aggs], "output_sheet": out}}  # noqa: E731

GOLD = {
    1: [[LK("部门", "部门", "未登记"), LK("职级", "职级", "未登记"), SAVE("01_销售业绩_处理后.xlsx")],
        [ADD("净销售额", "=[销售额]-[退货额]"), SAVE("x.xlsx")],
        [ADD("提成比例", "=XLOOKUP([净销售额],提成档位!$A:$A,提成档位!$B:$B,0,-1)"), SAVE("x.xlsx")],
        [ADD("应发提成", '=IF([回款状态]="未回款",0,[净销售额]*[提成比例])'), SAVE("x.xlsx")]],
    2: [[ADD("订单号", r'=IFERROR(REGEXEXTRACT([备注],"SO\d{8}-\d{2}"),"")'), SAVE("02_银行流水_处理后.xlsx")],
        [{"op": "regex_replace", "args": {"col": "备注", "pattern": r"[(（\[【〔{｛][^)）\]】〕}｝]*[)）\]】〕}｝]", "repl": ""}}, {"op": "trim", "args": {"cols": ["摘要"]}},
         SAVE("02_银行流水_处理后_处理后.xlsx")],
        [{"op": "delete_rows", "args": {"where": {"col": "摘要", "cmp": "contains", "value": "退款"}}}, SAVE("02_银行流水_处理后_处理后_处理后.xlsx")]],
    3: [[ADD("达标率", "=[完成额]/[目标额]"), ADD("核算奖金", "=IF([达标率]>=0.8,[完成额]*[权重系数]*0.05,0)"), SAVE("03_季度绩效_处理后.xlsx")],
        [GA(["部门", "季度"], [("核算奖金", "sum"), ("完成额", "sum")], "部门季度汇总"), SAVE("03_季度绩效_处理后_处理后.xlsx")]],
    4: [[{"op": "dedupe", "args": {}}, {"op": "split_to_files", "args": {"by": "大区", "path_template": "按大区拆分/{value}.xlsx"}},
         SAVE("04_客户订单_处理后.xlsx")]],
    5: [[CF({"col": "预算余额", "cmp": "<", "value": 0}, {"font_color": "FF0000"}),
         CF({"col": "超期天数", "cmp": ">", "value": 30}, {"fill_color": "FFFF00"}, "row"), SAVE("05_项目进度_处理后.xlsx")],
        [ADD("完成度百分位", "=PERCENTRANK.INC($F:$F,[完成度])"),
         CF({"col": "完成度百分位", "cmp": ">=", "value": 0.9}, {"fill_color": "C6EFCE", "bold": True}), SAVE("05_项目进度_处理后_处理后.xlsx")]],
    6: [[{"op": "regex_replace", "args": {"col": "金额", "pattern": "[¥￥,元]", "repl": ""}}, SAVE("06_报销明细_处理后.xlsx")],
        [{"op": "regex_replace", "args": {"col": "报销日期", "pattern": "[年月./]", "repl": "-"}},
         {"op": "regex_replace", "args": {"col": "报销日期", "pattern": "日", "repl": ""}}, SAVE("06_报销明细_处理后.xlsx")],
        [ADD("金额数值", "=VALUE([金额])"), ADD("标准日期", "=DATEVALUE([报销日期])"), ADD("年月", '=TEXT([标准日期],"yyyy-mm")'),
         ADD("周次", "=ISOWEEKNUM([标准日期])"), SAVE("06_报销明细_处理后_处理后.xlsx")],
        [GA(["周次"], [("金额数值", "sum")], "按周汇总"), SAVE("06_报销明细_处理后_处理后_处理后.xlsx")]],
    7: [[ADD("是否最新", '=IF([更新时间]=MAXIFS($E:$E,$A:$A,[订单号]),"保留","归档")'), SAVE("07_订单变更记录_处理后.xlsx")],
        [{"op": "filter_to_sheet", "args": {"where": {"col": "是否最新", "cmp": "==", "value": "归档"}, "output_sheet": "归档记录", "mode": "value"}},
         SAVE("07_订单变更记录_处理后_处理后.xlsx")],
        [{"op": "delete_rows", "args": {"where": {"col": "是否最新", "cmp": "==", "value": "归档"}}}, SAVE("07_订单变更记录_处理后_处理后_处理后.xlsx")]],
    8: [[{"op": "unique_to_sheet", "args": {"col": "供应商", "output_sheet": "供应商清单", "with_count": True}}, SAVE("08_采购明细_处理后.xlsx")],
        [ADD("采购金额", "=[数量]*[单价]"), SAVE("08_采购明细_处理后_处理后.xlsx")],
        [GA(["供应商"], [("采购金额", "sum")], "供应商金额"), SAVE("08_采购明细_处理后_处理后.xlsx")]],
    9: [[ADD("上半年合计", "=SUM(D{r}:I{r})"), SAVE("09_门店月度销售_处理后.xlsx")],
        [GA(["大区", "品类"], [("上半年合计", "sum")], "大区品类汇总"), SAVE("09_门店月度销售_处理后_处理后.xlsx")],
        [GA(["大区"], [(f"{m}月", "sum") for m in range(1, 7)], "大区月度矩阵"), SAVE("09_门店月度销售_处理后_处理后_处理后.xlsx")]],
    10: [[{"op": "delete_rows", "args": {"where": {"all": [{"col": c, "cmp": "is_empty"} for c in
                                                         ["工号", "姓名", "部门", "基本工资", "绩效奖金", "实发工资", "备注"]]}}},
          SAVE("10_工资表_处理后.xlsx")],
         [CF({"col": "实发工资", "cmp": "<", "value": 8000}, {"font_color": "FF0000"}),
          GA(["部门"], [("实发工资", "sum")], "部门汇总"), SAVE("10_工资表_处理后_处理后.xlsx")],
         [{"op": "delete_rows", "args": {"where": {"col": "备注", "cmp": "contains", "value": "离职"}}}, SAVE("10_工资表_处理后.xlsx")]],
}


def gold_plan(steps, file, sheet):
    out = []
    for i, s in enumerate(steps, 1):
        st = {"id": f"s{i}", **json.loads(json.dumps(s))}
        if st["op"] not in ("save",):
            st["sheet"] = sheet
        out.append(st)
    return {"version": "0.2", "lang": "zh", "inputs": [{"alias": "src", "path": file, "sheet": sheet}], "steps": out}


# ------------------------------------------------------------------ 结果摘要（宽松）
def _norm(v):
    if isinstance(v, XLError):
        return v.code
    if isinstance(v, bool):
        return v
    if isinstance(v, float):
        return round(v, 6) if not v.is_integer() else int(v)
    if hasattr(v, "year"):
        from tabularedge.formula_eval import to_serial
        return round(to_serial(v), 6)
    if v == "":
        return None
    return v


def _cols(ws, ev):
    out = []
    for c in range(1, ws.max_column + 1):
        col = []
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, c).value
            if hasattr(v, "text") and hasattr(v, "ref"):  # ArrayFormula（FILTER/UNIQUE）：记公式范围大小
                col.append(("ARRAY", v.ref.split(":")[-1].rstrip("0123456789") if ":" in v.ref else v.ref))
                continue
            try:
                col.append(_norm(ev.cell(ws, r, c)))
            except (Unsupported, Exception):  # noqa: BLE001
                col.append(("FORMULA", str(v)))
        while col and col[-1] is None:
            col.pop()
        out.append(tuple(col))
    return Counter(out)


def _cf(ws, ev):
    marks = Counter()
    for cf in ws.conditional_formatting:
        for rule in cf.rules:
            f = rule.formula[0] if rule.formula else None
            st = (_hue(rule.dxf.font.color.rgb[-6:]) if rule.dxf and rule.dxf.font and rule.dxf.font.color else None,
                  _hue(rule.dxf.fill.fgColor.rgb[-6:]) if rule.dxf and rule.dxf.fill and rule.dxf.fill.fgColor is not None and rule.dxf.fill.fgColor.rgb else None,
                  bool(rule.dxf.font.b) if rule.dxf and rule.dxf.font else False)
            if st == (None, None, False):
                continue  # 没有任何样式的规则 = 无效果
            first = min(rg.min_row for rg in cf.sqref.ranges)
            for rg in cf.sqref.ranges:
                for r in range(rg.min_row, rg.max_row + 1):
                    shifted = _shift(f, r - first)
                    try:
                        hit = ev.eval("=" + shifted, ws, r)
                    except Exception:  # noqa: BLE001
                        hit = ("ERR", shifted)
                    if hit is True:
                        for c in range(rg.min_col, rg.max_col + 1):
                            marks[(r, c, st)] += 1
    return frozenset(marks)


def _hue(rgb):
    """颜色归类（红/橙/黄/绿/蓝/紫/灰）：浅绿 C6EFCE 与 90EE90 视为同一种“绿色”。"""
    import colorsys
    try:
        r, g, b = (int(rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except (ValueError, TypeError):
        return rgb
    h, l, sat = colorsys.rgb_to_hls(r, g, b)
    if sat < 0.15:
        return "grey"
    deg = h * 360
    for lim, name in ((15, "red"), (40, "orange"), (70, "yellow"), (170, "green"), (260, "blue"), (330, "purple"), (361, "red")):
        if deg < lim:
            return name


def _shift(f, dr):
    import re
    return re.sub(r"(\$?[A-Z]{1,3})(\$?)(\d+)", lambda m: m.group(1) + m.group(2) + (m.group(3) if m.group(2) else str(int(m.group(3)) + dr)), f)


def digest(wd: Path, main_file: str, main_sheet: str, skip_sheets=()):
    p = wd / main_file
    res = {"exists": p.exists()}
    if p.exists():
        wb = load_workbook(p)
        ev = Evaluator(wb)
        ws = wb[main_sheet] if main_sheet in wb.sheetnames else wb.worksheets[0]
        res["main"] = _cols(ws, ev)
        res["cf"] = _cf(ws, ev)
        res["new"] = sorted((_cols(x, ev) for x in wb.worksheets if x.title != ws.title and x.title not in skip_sheets),
                            key=lambda c: repr(sorted(c.items(), key=repr)))
    res["split"] = {str(f.relative_to(wd)).split("/")[-1]: _cols(load_workbook(f).worksheets[0], Evaluator(load_workbook(f)))
                    for f in wd.rglob("*.xlsx") if f.parent != wd and "_backup" not in f.parts}
    return res


def same(a, b):
    diffs = []
    for k in ("exists", "main", "cf", "new", "split"):
        if a.get(k) != b.get(k):
            diffs.append(k)
    return not diffs, diffs


# ------------------------------------------------------------------ run
def newest_output(wd: Path, before: dict, cur: str):
    changed = [p for p in wd.glob("*.xlsx") if before.get(p) != p.stat().st_mtime]
    changed.sort(key=lambda p: -p.stat().st_mtime)
    return str(changed[0].relative_to(wd)) if changed else cur


def run_chain(scn, src_dir: Path, mode, url=None, log=print, compat=None):
    wd = Path(tempfile.mkdtemp(prefix=f"case{scn['id']}_"))
    for f in src_dir.glob("*.xlsx"):
        shutil.copy2(f, wd / f.name)
    cur, sheet = scn["file"], scn["sheet"]
    steps_out = []
    for k, st in enumerate(scn["steps"]):
        text = st["text"]
        rec = {"step": k + 1, "text": text, "input": cur}
        before = {p: p.stat().st_mtime for p in wd.glob("*.xlsx")}
        t0 = time.time()
        prof, extra, refs = prompt_context(wd, cur, sheet, text)
        rec["refs"] = [r[0] for r in refs]
        if mode == "gold":
            plan = gold_plan(GOLD[scn["id"]][k], cur, sheet)
            for st_ in plan["steps"]:  # 金标另存名统一为 <输入>_处理后.xlsx（与模型习惯一致）
                if st_["op"] == "save":
                    st_["args"]["path"] = Path(cur).stem + "_处理后.xlsx"
            if scn["id"] == 10 and k == 2:
                plan["steps"][-1]["args"]["path"] = cur
            info = {}
        else:
            times = {}
            plan, info = llm_parser.parse(text, prof, cur, extra_profile=extra, url=url,
                                          on_summary=lambda s: times.setdefault("summary", (s, time.time() - t0)))
            plan = plan or {"steps": []}
            rec["summary"] = (times.get("summary") or (plan.get("summary"), None))
            rec["ambiguities"] = plan.get("ambiguities")
            rec["failed"] = info.get("explain") if info.get("failed") else None
            rec["fixes"] = info.get("fixes")
            rec["dropped"] = info.get("dropped")
            rec["attempts"] = info.get("attempts")
            if info.get("failed"):
                plan = {**plan, "steps": []}  # 无法修复的计划不执行（界面同样不允许执行）
        rec["plan_steps"] = [{k2: v for k2, v in s.items() if k2 != "id"} for s in plan["steps"]]
        rec["parse_s"] = round(time.time() - t0, 2)
        risky = []
        ex = RefExecutor(Workspace(wd), lambda s: (risky.append(s["op"]) if is_risky(s, cur) else None) or True, refs, compat=compat)
        p2 = json.loads(json.dumps(plan))
        p2.pop("ambiguities", None)
        try:
            rec["log"] = ex.run(p2)
            rec["ok_exec"] = True
        except Exception as e:  # noqa: BLE001
            rec["log"] = [f"{type(e).__name__}: {e}"]
            rec["ok_exec"] = False
        rec["risky_confirms"] = risky
        nxt = newest_output(wd, before, cur)
        rec["output"] = nxt
        rec["digest"] = digest(wd, nxt, sheet, skip_sheets=())
        steps_out.append(rec)
        log(f"  [{scn['id']}.{k + 1}] {mode} {rec['parse_s']}s -> {nxt} risky={risky} {'' if rec['ok_exec'] else rec['log']}")
        cur = nxt
    return wd, steps_out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["gold", "model", "both"], default="both")
    ap.add_argument("--url", default="http://127.0.0.1:18080")
    ap.add_argument("--cases", default=str(Path(__file__).resolve().parent))
    ap.add_argument("--only", type=int, nargs="*")
    ap.add_argument("--out", default="outputs/cases")
    ap.add_argument("--excel", choices=["365", "2019"], default="2019", help="目标 Excel 版本（宿主机是 Excel 2019）")
    a = ap.parse_args()
    src = Path(a.cases)
    scen = json.loads((src / "scenarios.json").read_text(encoding="utf-8"))
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    report = []
    for s in scen:
        if a.only and s["id"] not in a.only:
            continue
        print(f"== {s['id']}. {s['title']}", flush=True)
        compat = "2019" if a.excel == "2019" else None
        g_wd, g = run_chain(s, src, "gold", compat=compat)
        row = {"id": s["id"], "title": s["title"], "gold": [{k: v for k, v in x.items() if k != "digest"} for x in g]}
        if a.mode in ("model", "both"):
            m_wd, m = run_chain(s, src, "model", a.url, compat=compat)
            row["model"] = []
            for gx, mx in zip(g, m):
                ok, diffs = same(gx["digest"], mx["digest"])
                row["model"].append({**{k: v for k, v in mx.items() if k != "digest"}, "match_gold": ok, "diff": diffs})
                print(f"    step {mx['step']}: {'PASS' if ok else 'DIFF ' + ','.join(diffs)}  | {mx.get('summary')}", flush=True)
            shutil.copytree(m_wd, out / f"case{s['id']:02d}_model", dirs_exist_ok=True)
        shutil.copytree(g_wd, out / f"case{s['id']:02d}_gold", dirs_exist_ok=True)
        report.append(row)
    (out / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    if a.mode != "gold":
        n = sum(len(r["model"]) for r in report)
        k = sum(x["match_gold"] for r in report for x in r["model"])
        print(f"TOTAL steps {k}/{n} match gold; scenarios fully passed "
              f"{sum(all(x['match_gold'] for x in r['model']) for r in report)}/{len(report)}")


if __name__ == "__main__":
    main()
