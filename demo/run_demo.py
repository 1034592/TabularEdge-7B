"""TabularEdge 端到端 Demo（宿主机本地、完全离线）。

  自然语言（中/英）→ 本地 GGUF 模型（llama-server）流式编译 Action Plan
  → 1 秒级先显示“我理解的是……”摘要 → 校验 / 歧义确认 → 执行器在真实 .xlsx 上修改 → 读回验证

    python demo/run_demo.py                                  # 跑内置示例
    python demo/run_demo.py -i "把带括号的备注删了，金额负数的标红，按部门汇总新表"
    python demo/run_demo.py --interactive

需要 llama-server 已在 127.0.0.1:8080 运行（启动命令见 README 的 Quickstart）。
"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
import time
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from tabularedge import llm_parser  # noqa: E402
from tabularedge.executor import Executor, Workspace  # noqa: E402
from tabularedge.protocol import profile_sheet  # noqa: E402

LEDGER = [
    ["日期", "部门", "金额", "备注"],
    ["2026-10-01", "销售部", 12000, "客户A（首付）"],
    ["2026-10-02", "财务部", -300, "退款(重复扣款)"],
    ["2026-10-03", "销售部", 8000, "客户B"],
    ["2026-10-05", "研发部", -50.5, "  报销（打车） "],
    ["2026-10-06", "财务部", 1000, None],
    ["2026-10-08", "研发部", 23000, "服务器采购（一期）"],
    ["2026-10-09", "市场部", -1200, "活动退款"],
    ["2026-10-09", "市场部", 4500, None],
]
SALES = [
    ["Order ID", "Sales Rep", "Region", "Product", "Qty", "Unit Price", "Remarks"],
    ["SO-1001", "Alice", "East", "Laptop", 3, 5999.0, "urgent (VIP)"],
    ["SO-1002", "Bob", "South", "Monitor", 10, 1299.0, ""],
    ["SO-1003", "Alice", "East", "Mouse", 50, 59.9, None],
    ["SO-1004", "Carol", "North", "Laptop", 1, 6299.0, "demo unit (return)"],
    ["SO-1004", "Carol", "North", "Laptop", 1, 6299.0, "demo unit (return)"],
    ["SO-1005", "Dave", "West", "Keyboard", 20, 199.0, None],
]
DIMS = {"ledger.xlsx": ("部门预算", [["部门", "预算", "负责人"], ["销售部", 30000, "张伟"], ["财务部", 5000, "王芳"],
                                    ["研发部", 50000, "李娜"]]),
        "sales.xlsx": ("Catalog", [["Product", "Category", "List Price"], ["Laptop", "PC", 6299], ["Monitor", "Display", 1399],
                                   ["Mouse", "Accessory", 69]])}
EXAMPLES_V2 = [  # v2：烂输入 / 歧义 / 跨表 / 破坏性操作
    ("ledger.xlsx", "流水", "帮我把被注里带扩号的内容去掉，金额负的表红，按不门hz一下放新标"),
    ("ledger.xlsx", "流水", "把带括号的备注删了"),
    ("ledger.xlsx", "流水", "按部门从部门预算表把预算查过来，查不到写“未找到”，再把金额大于1000的筛出来放新表"),
    ("sales.xlsx", "Orders", "pls dedupe by order id, add col revenue = qty*unit price, then list distinct regions w/ counts"),
    ("ledger.xlsx", "流水", "把金额小于0的行删了，改完直接覆盖原文件"),
    ("sales.xlsx", "Orders", "按单价降序排一下，弄完保存"),
]
EXAMPLES = [
    ("ledger.xlsx", "流水", "把带括号的备注删了，金额负数的标红，按部门汇总新表"),
    ("ledger.xlsx", "流水", "remove the stuff in brackets from the notes, make negative amounts red and total amount by department"),
    ("sales.xlsx", "Orders", "按订单号去重，加一列销售额等于数量乘单价，然后按区域汇总销售额"),
    ("sales.xlsx", "Orders", "sort by unit price high to low and split into one file per region"),
]


def make_files(ws: Path):
    for name, sheet, rows in (("ledger.xlsx", "流水", LEDGER), ("sales.xlsx", "Orders", SALES)):
        wb = Workbook()
        wb.active.title = sheet
        for r in rows:
            wb.active.append(r)
        if name in DIMS:
            dn, drows = DIMS[name]
            d = wb.create_sheet(dn)
            for r in drows:
                d.append(r)
        wb.save(ws / name)


def show_result(ws: Path, plan: dict):
    for s in plan["steps"]:
        if s["op"] == "save":
            p = ws / s["args"]["path"]
            if p.suffix == ".xlsx" and p.exists():
                wb = load_workbook(p)
                for sh in wb.worksheets:
                    print(f"    ── {p.name} / {sh.title}")
                    for row in sh.iter_rows(values_only=True, max_row=10):
                        print("      " + " | ".join("" if v is None else str(v) for v in row))
                    for cf in sh.conditional_formatting:
                        for r in cf.rules:
                            print(f"      [条件格式] {cf.sqref}: ={r.formula[0]}")
    for p in sorted(ws.rglob("*.xlsx")):
        if "_backup" not in p.parts:
            print("    file:", p.relative_to(ws))


def run_one(ws: Path, file: str, sheet: str, text: str, style: str, auto_yes: bool):
    print("\n" + "=" * 100)
    print("指令:", text)
    wbk = load_workbook(ws / file)
    prof = profile_sheet(wbk[sheet])
    extra = "\n".join(profile_sheet(x).to_prompt() for x in wbk.worksheets if x.title != sheet)
    t0 = time.time()
    plan, info = llm_parser.parse(text, prof, file, style=style, extra_profile=extra,
                                  on_summary=lambda s: print(f"  [{time.time() - t0:5.2f}s] 我理解的是：{s}", flush=True))
    if info.get("failed"):
        print("  无法得到可执行的计划：\n   - " + "\n   - ".join(info["explain"]))
        return
    print(f"  [{time.time() - t0:5.2f}s] 完整计划（{len(plan['steps'])} 步，尝试 {info['attempts']} 次"
          f"{'，自动修正 ' + str(len(info['fixes'])) + ' 处' if info.get('fixes') else ''}）:")
    for s in plan["steps"]:
        print(f"     {s['id']} {s['op']:<18} {json.dumps(s['args'], ensure_ascii=False)}")
    for amb in plan.get("ambiguities", []):
        print(f"  ❓ {amb['question']}")
        alts = amb.get("alternatives") or []
        choice = "1"
        if alts and not auto_yes:
            choice = input("     1=按当前计划  2=改用备选  > ").strip() or "1"
        if choice == "2" and alts:
            plan["steps"] = [alts[0] if s["id"] == amb["step_id"] else s for s in plan["steps"]]
            print("     → 已切换为备选:", alts[0]["op"])
        else:
            print("     → 按当前计划执行")
    plan.pop("ambiguities", None)
    def confirm(st):
        risky = st["op"] in ("delete_rows", "dedupe") or (st["op"] == "save" and st["args"]["path"] == file)
        tag = "⚠ 破坏性操作确认" if risky else "写盘确认"
        if auto_yes:
            print(f"  [{tag}] {st['op']} {json.dumps(st['args'], ensure_ascii=False)} → 自动确认（--yes）")
            return True
        return input(f"  [{tag}] {st['op']} {st['args']} [Y/n] ").strip().lower() != "n"
    log = Executor(Workspace(ws), confirm=confirm).run(plan)
    print("  执行日志:")
    for line in log:
        print("    ", line)
    show_result(ws, plan)
    return {"text": text, "latency": info, "steps": [s["op"] for s in plan["steps"]], "log": log}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--instruction")
    ap.add_argument("--file", default="ledger.xlsx")
    ap.add_argument("--sheet", default="流水")
    ap.add_argument("--interactive", action="store_true")
    ap.add_argument("--style", default="sft", choices=["sft", "fewshot"])
    ap.add_argument("--workspace", default=str(ROOT / "outputs" / "demo_ws"))
    ap.add_argument("--yes", action="store_true", help="自动确认（录制/测试用）")
    ap.add_argument("--v2", action="store_true", help="跑 v2 示例（烂输入 / 歧义 / 跨表 / 破坏性操作）")
    a = ap.parse_args()
    ws = Path(a.workspace)
    shutil.rmtree(ws, ignore_errors=True)
    ws.mkdir(parents=True)
    make_files(ws)
    print("工作区:", ws, "（示例文件 ledger.xlsx / sales.xlsx 已生成）")
    results = []
    if a.instruction:
        results.append(run_one(ws, a.file, a.sheet, a.instruction, a.style, a.yes))
    elif a.interactive:
        while True:
            t = input("\n请输入指令（回车退出）> ").strip()
            if not t:
                break
            results.append(run_one(ws, a.file, a.sheet, t, a.style, a.yes))
    else:
        for k, (f, sh, t) in enumerate(EXAMPLES_V2 if a.v2 else EXAMPLES, 1):
            sub = ws / f"ex{k}"  # 每个示例独立子目录，互不影响
            sub.mkdir(parents=True, exist_ok=True)
            make_files(sub)
            results.append(run_one(sub, f, sh, t, a.style, a.yes))
    (ws / "demo_results.json").write_text(json.dumps(results, ensure_ascii=False, indent=1, default=str), encoding="utf-8")


if __name__ == "__main__":
    main()
