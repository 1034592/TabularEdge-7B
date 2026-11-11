"""反事实因果语料 v2（Action Plan 0.2）：在 v1 之上扩展函数族 / 跨表 / 歧义体系 / 人类扰动，多进程执行验证。

新增
  主题   invoice（训练，含税金额/不含税金额两列“金额”）、project（只测试，预算金额/实际金额）
         每个主题附一张维度表（部门预算 / 产品目录 / 员工信息 / 供应商 / 客户信息 …），用于跨表 XLOOKUP
  意图   lookup_column(XLOOKUP) · filter_to_sheet(FILTER) · unique_to_sheet(UNIQUE+COUNTIF)
         add_column 公式族：TEXTJOIN / IFS 分档 / ROUND / YEAR·MONTH / LEFT · 覆盖原文件保存
  歧义   A1 带括号的X删了（删字 vs 删行）        A2 “金额”指哪一列（两列都叫金额）
         A3 重复的订单删了（按单号 vs 整行相同）  A4 把空的删了（整行全空 vs 某列为空）
         A5 改完保存（另存 vs 覆盖原文件）
         每条歧义样本以 70% 概率配一条“已澄清”的反事实兄弟（单槽干预：补上那个缺失的信息 → 不再追问、执行备选解释）
  扰动   synth/noise.py；训练集 45% 加训练族扰动；test_v2_noisy = test_v2_ood 的指令加“留出族”扰动（没见过的错法）

    python synth/gen2.py --n 24000 --workers 18 --out data/synth_v2
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import multiprocessing as mp
import random
import shutil
import sys
import tempfile
import time
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "synth"))
import gen as G  # noqa: E402
from gen import C, NAMES, PAREN_RX, JOIN, say, build_plan, utter, digest  # noqa: E402
from noise import perturb  # noqa: E402
from tabularedge.executor import Executor, Workspace  # noqa: E402
from tabularedge.protocol import profile_sheet, validate  # noqa: E402

# ------------------------------------------------------------------ 新主题（注入 v1 的 THEMES，Table 直接可用）
STATUS = (["进行中", "已完成", "暂停"], ["Active", "Done", "On hold"])
G.THEMES["invoice"] = [C("发票号", "Invoice No", "text", "id", ("票号",), ("invoice number",)),
                       C("客户", "Client", "text", "cat", ("客户名称",), ("customer",), NAMES),
                       C("含税金额", "Gross Amount", "num", "amount", ("含税额",), ("gross",)),
                       C("不含税金额", "Net Amount", "num", "amount", ("净额",), ("net",)),
                       C("开票日期", "Invoice Date", "date", "date", ("日期",), ("date",)),
                       C("备注", "Notes", "text", "note", (), ("remarks",))]
G.THEMES["project"] = [C("项目编号", "Project ID", "text", "id", ("编号",), ("project code",)),
                       C("负责人", "Owner", "text", "cat", ("经办人",), ("lead",), NAMES),
                       C("状态", "Status", "text", "cat", (), ("state",), STATUS),
                       C("预算金额", "Budget Amount", "num", "amount", ("预算",), ("budget",)),
                       C("实际金额", "Actual Amount", "num", "amount", ("实际",), ("actual", "spend")),
                       C("备注", "Remarks", "text", "note", (), ("notes",))]
V3 = False  # --v3：收官补强（新意图、更多说法、扩充缩写词表）；默认 False 以保证 v2 语料可复现
G.THEMES["attendance"] = [C("工号", "Emp ID", "text", "id", ("员工号",), ("employee id",)),
                          C("姓名", "Name", "text", "cat", ("员工",), ("employee",), NAMES),
                          C("班组", "Team", "text", "cat", ("组别",), ("group",), (["一组", "二组", "三组"], ["Team A", "Team B", "Team C"])),
                          C("加班时长", "Overtime Hours", "num", "amount", ("加班",), ("overtime",)),
                          C("请假时长", "Leave Hours", "num", "amount", ("请假",), ("leave",)),
                          C("备注", "Notes", "text", "note", (), ("remarks",))]
TRAIN_THEMES = ["ledger", "sales", "expense", "inventory", "invoice"]
TRAIN_THEMES_V3 = TRAIN_THEMES + ["attendance"]
OOD_THEMES = ["payroll", "receivable", "project"]
SHEET_ZH = {"attendance": "考勤", "ledger": "流水", "sales": "销售明细", "expense": "报销", "inventory": "库存", "payroll": "工资表",
            "receivable": "应收", "invoice": "发票", "project": "项目台账"}
SHARED = {"invoice": ("金额", "amount"), "project": ("金额", "amount"), "attendance": ("时长", "hours")}  # A2：两列都能叫“金额”

# 维度表：主表键列（按表头 zh 名）→ 维度表 [键, 两个属性列]
DIM = {
    "ledger": ("部门", ("部门预算", "Budget"), [("部门", "Department"), ("预算", "Budget", "num"), ("负责人", "Owner", "name")]),
    "sales": ("产品", ("产品目录", "Catalog"), [("产品", "Product"), ("品类", "Category", "text"), ("标准价", "List Price", "num")]),
    "expense": ("员工", ("员工信息", "Staff"), [("员工", "Employee"), ("工号", "Staff ID", "code"), ("职级", "Grade", "text")]),
    "inventory": ("物料编码", ("供应商", "Suppliers"), [("物料编码", "SKU"), ("供应商", "Supplier", "name"), ("安全库存", "Safety Stock", "num")]),
    "invoice": ("客户", ("客户信息", "Clients"), [("客户", "Client"), ("信用等级", "Rating", "text"), ("客户经理", "Account Manager", "name")]),
    "payroll": ("部门", ("部门信息", "Divisions"), [("部门", "Division"), ("经理", "Manager", "name"), ("编制", "Headcount", "num")]),
    "receivable": ("客户", ("客户档案", "Customers"), [("客户", "Customer"), ("信用额度", "Credit Limit", "num"), ("联系人", "Contact", "name")]),
    "attendance": ("班组", ("班组信息", "Teams"), [("班组", "Team"), ("组长", "Leader", "name"), ("人数", "Headcount", "num")]),
    "project": ("负责人", ("人员信息", "People"), [("负责人", "Owner"), ("部门", "Department", "text"), ("工号", "Staff ID", "code")]),
}


class Table2(G.Table):
    def __init__(self, theme, hlang, rng):
        super().__init__(theme, hlang, rng)
        if hlang == "zh":
            self.sheet = SHEET_ZH[theme]
        key_zh, names, cols = DIM[theme]
        self.key = next(c for c in self.cols if c["zh"] == key_zh)
        self.dim_sheet = names[0] if hlang == "zh" else names[1]
        self.dim_cols = cols
        self.dim_heads = [c[0] if hlang == "zh" else c[1] for c in cols]

    def make_xlsx(self, path, rng, n=24):
        super().make_xlsx(path, rng, n)
        wb = load_workbook(path)
        ws = wb[self.sheet]
        idc = [i for i, c in enumerate(self.cols, 1) if c["role"] == "id"]
        if idc and rng.random() < 0.6:  # A3 需要：同单号但内容不同的行（按单号去重 ≠ 整行去重）
            row = [c.value for c in ws[2]]
            j = next((i for i, c in enumerate(self.cols) if c["type"] == "num"), None)
            if j is not None and isinstance(row[j], (int, float)):
                row[j] = row[j] + 1
            ws.append(row)
        if rng.random() < 0.4:  # A4 需要：整行全空
            ws.append([None] * len(self.cols))
        kc = self.cols.index(self.key) + 1
        keys = []
        for r in range(2, ws.max_row + 1):
            v = ws.cell(r, kc).value
            if v not in (None, "") and v not in keys:
                keys.append(v)
        rng.shuffle(keys)
        keys = keys[:max(1, len(keys) - rng.randint(1, 2))]  # 维度表故意缺 1–2 个键 → if_not_found 有效果
        dim = wb.create_sheet(self.dim_sheet)
        dim.append(self.dim_heads)
        for k in sorted(keys, key=str):
            row = [k]
            for c in self.dim_cols[1:]:
                kind = c[2]
                row.append(rng.choice([1000, 3000, 5000, 8000, 12000]) if kind == "num" else
                           rng.choice(["A", "B", "C", "P1", "P2"]) if kind == "text" else
                           f"E{rng.randint(100, 999)}" if kind == "code" else rng.choice(NAMES[0] + NAMES[1]))
            dim.append(row)
        wb.save(path)


def say2(c, lang, rng, t):
    return say(c, lang, rng, t)


# ------------------------------------------------------------------ 新意图
class LookupCol(G.Intent):
    name = "lookup_column"
    stage = 2
    NF = ["未找到", "N/A", "无", "-"]
    NF_V3 = NF + [0, "空", "not found"]

    def slots(self, t, rng):
        return {"ret": rng.choice([1, 2]), "nf": rng.choice(self.NF_V3 if V3 else self.NF), "mode": "formula" if rng.random() < 0.85 else "value"}

    def cf(self, s, t, rng):
        k = rng.choice(["ret", "nf", "mode"])
        if k == "ret":
            return {**s, "ret": 3 - s["ret"]}, "返回列"
        if k == "nf":
            return {**s, "nf": rng.choice([x for x in (self.NF_V3 if V3 else self.NF) if x != s["nf"]])}, "缺省值"
        return {**s, "mode": "value" if s["mode"] == "formula" else "formula"}, "公式/值"

    def step(self, s, t, lang):
        ret = t.dim_heads[s["ret"]]
        name = ret if ret not in t.headers() else ret + ("(查)" if t.hlang == "zh" else " (lookup)")
        a = {"name": name, "key": t.h(t.key), "from_sheet": t.dim_sheet, "from_key": t.dim_heads[0], "return": ret,
             "if_not_found": s["nf"]}
        if s["mode"] == "value":
            a["mode"] = "value"
        return {"op": "lookup_column", "args": a}

    def text(self, s, t, lang, rng):
        k, ret, d, nf = say(t.key, lang, rng, t), t.dim_heads[s["ret"]], t.dim_sheet, s["nf"]
        num = isinstance(nf, (int, float))
        if lang == "zh":
            q = str(nf) if num else f"“{nf}”"
            base = rng.choice([f"按{k}从{d}表把{ret}查过来，查不到写{q}", f"用XLOOKUP把{d}里的{ret}带过来（按{k}匹配），没有的填{q}",
                               f"新增一列{ret}，从“{d}”按{k}匹配，找不到显示{q}"])
            return base + (rng.choice(["，只要结果值不要公式", "，直接贴成值"]) if s["mode"] == "value" else "")
        q = str(nf) if num else f'"{nf}"'
        base = rng.choice([f'look up {ret} from the {d} sheet by {k}, show {q} if missing',
                           f'pull {ret} in from {d} matching on {k}; use {q} when not found',
                           f'add a {ret} column via XLOOKUP against {d} on {k}, {q} if no match'])
        return base + (rng.choice([", values only, no formulas", ", paste as values"]) if s["mode"] == "value" else "")

    def summ(self, s, t, lang):
        r = t.dim_heads[s["ret"]]
        return f"按{t.h(t.key)}从{t.dim_sheet}查{r}" if lang == "zh" else f"lookup {r} from {t.dim_sheet} by {t.h(t.key)}"


class FilterSheet(G.Intent):
    name = "filter_to_sheet"
    stage = 4

    def ok(self, t): return bool(t.num_cols() or [c for c in t.by_role("cat") if c["values"]])

    def slots(self, t, rng):
        cats = [c for c in t.by_role("cat") if c["values"]]
        if t.num_cols() and (not cats or rng.random() < 0.6):
            return {"c": rng.choice(t.num_cols()), "cmp": rng.choice([">", "<", ">="]), "v": rng.choice([0, 100, 1000, 5000])}
        c = rng.choice(cats)
        return {"c": c, "cmp": "==", "v": rng.choice(c["values"][0 if t.hlang == "zh" else 1])}

    def cf(self, s, t, rng):
        if s["cmp"] == "==":
            vals = [v for v in s["c"]["values"][0 if t.hlang == "zh" else 1] if v != s["v"]]
            return {**s, "v": rng.choice(vals)}, "筛选值"
        if rng.random() < 0.5:
            return {**s, "cmp": "<" if s["cmp"] != "<" else ">"}, "比较方向"
        return {**s, "v": rng.choice([v for v in [0, 100, 1000, 5000] if v != s["v"]])}, "阈值"

    def step(self, s, t, lang):
        h = t.h(s["c"])
        return {"op": "filter_to_sheet", "args": {"where": {"col": h, "cmp": s["cmp"], "value": s["v"]},
                                                   "output_sheet": (f"筛选_{h}" if lang == "zh" else f"Filtered {h}")[:31]}}

    def text(self, s, t, lang, rng):
        c, v = say(s["c"], lang, rng, t), s["v"]
        if s["cmp"] == "==":
            return rng.choice([f"把{c}是“{v}”的记录筛出来放到新表", f"筛选{c}为“{v}”的行，单独放一张表"] if lang == "zh" else
                              [f'filter the rows where {c} is "{v}" into a new sheet', f'put the "{v}" {c} rows on their own sheet'])
        if lang == "zh":
            w = {">": rng.choice(["大于", "超过"]), "<": rng.choice(["小于", "低于"]), ">=": "不低于"}[s["cmp"]]
            return rng.choice([f"把{c}{w}{v}的记录筛出来放新表", f"用FILTER把{c}{w}{v}的行提取到新sheet", f"{c}{w}{v}的单独筛一张表出来"])
        w = {">": rng.choice(["over", "above"]), "<": rng.choice(["under", "below"]), ">=": "at least"}[s["cmp"]]
        return rng.choice([f"filter rows with {c} {w} {v} into a new sheet", f"extract everything where {c} is {w} {v} to its own tab"])

    def summ(self, s, t, lang):
        return f"筛选{t.h(s['c'])}{s['cmp']}{s['v']}到新表" if lang == "zh" else f"filter {t.h(s['c'])} {s['cmp']} {s['v']} -> new sheet"


class UniqueSheet(G.Intent):
    name = "unique_to_sheet"
    stage = 4

    def ok(self, t): return bool(t.by_role("cat"))
    def slots(self, t, rng): return {"c": rng.choice(t.by_role("cat")), "cnt": rng.random() < 0.5}
    def cf(self, s, t, rng): return {**s, "cnt": not s["cnt"]}, "是否计数"

    def step(self, s, t, lang):
        h = t.h(s["c"])
        a = {"col": h, "output_sheet": (f"{h}清单" if lang == "zh" else f"{h} list")[:31]}
        if s["cnt"]:
            a["with_count"] = True
            if lang == "en" and t.hlang == "zh":
                a["count_header"] = "Count"
            if lang == "zh" and t.hlang == "en":
                a["count_header"] = "计数"
        return {"op": "unique_to_sheet", "args": a}

    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        if lang == "zh":
            base = rng.choice([f"列出所有不重复的{c}，放新表", f"把{c}去重后的清单单独放一张表", f"用UNIQUE提取{c}的唯一值"])
            return base + (rng.choice(["，并统计每个出现几次", "，顺便数一下各有几条"]) if s["cnt"] else "")
        base = rng.choice([f"list the distinct {c} values in a new sheet", f"give me a unique list of {c} on a separate tab"])
        return base + (rng.choice([" with counts", ", and count how many of each"]) if s["cnt"] else "")

    def summ(self, s, t, lang):
        return f"{t.h(s['c'])}唯一值清单{'+计数' if s['cnt'] else ''}" if lang == "zh" else f"unique {t.h(s['c'])}{' + counts' if s['cnt'] else ''}"


class FormulaCol(G.Intent):
    """add_column 的函数族：TEXTJOIN / IFS / ROUND / YEAR·MONTH / LEFT。"""
    name = "formula_col"
    stage = 2

    def kinds(self, t):
        k = []
        if len([c for c in t.cols if c["type"] == "text" and c["role"] in ("cat", "id", "name")]) >= 2:
            k.append("join")
        if t.by_role("amount"):
            k.append("ifs")
        if t.num_cols():
            k.append("round")
        if t.by_role("date"):
            k.append("datepart")
        if t.by_role("id"):
            k.append("left")
        return k

    def ok(self, t): return bool(self.kinds(t))

    def slots(self, t, rng):
        k = rng.choice(self.kinds(t))
        if k == "join":
            a, b = rng.sample([c for c in t.cols if c["type"] == "text" and c["role"] in ("cat", "id", "name")], 2)
            return {"k": k, "a": a, "b": b, "sep": rng.choice(["-", "/", "_"])}
        if k == "ifs":
            hi = rng.choice([5000, 10000])
            return {"k": k, "c": rng.choice(t.by_role("amount")), "hi": hi, "lo": rng.choice([100, 1000])}
        if k == "round":
            return {"k": k, "c": rng.choice(t.num_cols()), "d": rng.choice([-2, -1, 0, 1, 2] if V3 else [0, 1, 2])}
        if k == "datepart":
            return {"k": k, "c": t.by_role("date")[0], "fn": rng.choice(["YEAR", "MONTH"])}
        return {"k": k, "c": t.by_role("id")[0], "n": rng.choice([1, 2, 3, 4] if V3 else [2, 3])}

    def cf(self, s, t, rng):
        k = s["k"]
        if k == "join":
            return {**s, "sep": rng.choice([x for x in "-/_" if x != s["sep"]])}, "分隔符"
        if k == "ifs":
            return ({**s, "hi": 15000 - s["hi"]}, "上档阈值") if rng.random() < 0.5 else ({**s, "lo": 1100 - s["lo"]}, "下档阈值")
        if k == "round":
            return {**s, "d": rng.choice([d for d in ((-2, -1, 0, 1, 2) if V3 else (0, 1, 2)) if d != s["d"]])}, "小数位"
        if k == "datepart":
            return {**s, "fn": "MONTH" if s["fn"] == "YEAR" else "YEAR"}, "年/月"
        return {**s, "n": 5 - s["n"] if s["n"] in (2, 3) else (3 if s["n"] != 3 else 2)}, "位数"

    def _name(self, s, t):
        zh = t.hlang == "zh"
        k = s["k"]
        if k == "join":
            return f"{t.h(s['a'])}{t.h(s['b'])}" if zh else f"{t.h(s['a'])} {t.h(s['b'])}"
        if k == "ifs":
            return "金额档位" if zh else "Tier"
        if k == "round":
            if s["d"] < 0:
                return f"{t.h(s['c'])}({'百' if s['d'] == -2 else '十'}位)" if zh else f"{t.h(s['c'])} rounded{s['d']}"
            return f"{t.h(s['c'])}(取整)" if zh and s["d"] == 0 else f"{t.h(s['c'])}({s['d']}位)" if zh else f"{t.h(s['c'])} rounded"
        if k == "datepart":
            return ("年份" if s["fn"] == "YEAR" else "月份") if zh else ("Year" if s["fn"] == "YEAR" else "Month")
        return "前缀" if zh else "Prefix"

    def step(self, s, t, lang):
        k = s["k"]
        if k == "join":
            f = f'=TEXTJOIN("{s["sep"]}",TRUE,[{t.h(s["a"])}],[{t.h(s["b"])}])'
        elif k == "ifs":
            hi, mid, lo = ("高", "中", "低") if t.hlang == "zh" else ("High", "Medium", "Low")
            c = t.h(s["c"])
            f = f'=IFS([{c}]>={s["hi"]},"{hi}",[{c}]>={s["lo"]},"{mid}",TRUE,"{lo}")'
        elif k == "round":
            f = f"=ROUND([{t.h(s['c'])}],{s['d']})"
        elif k == "datepart":
            f = f"={s['fn']}([{t.h(s['c'])}])"
        else:
            f = f"=LEFT([{t.h(s['c'])}],{s['n']})"
        return {"op": "add_column", "args": {"name": self._name(s, t), "formula": f}}

    def text(self, s, t, lang, rng):
        k, zh = s["k"], lang == "zh"
        if k == "join":
            a, b = say(s["a"], lang, rng, t), say(s["b"], lang, rng, t)
            return rng.choice([f"加一列把{a}和{b}用“{s['sep']}”连起来", f"新增一列：{a}“{s['sep']}”{b}拼接"] if zh else
                              [f'add a column joining {a} and {b} with "{s["sep"]}"', f'concatenate {a} and {b} separated by "{s["sep"]}" in a new column'])
        if k == "ifs":
            c = say(s["c"], lang, rng, t)
            return (f"加一列档位：{c}≥{s['hi']}为高，≥{s['lo']}为中，其余为低" if zh else
                    f"add a tier column: {c} >= {s['hi']} is High, >= {s['lo']} is Medium, otherwise Low")
        if k == "round":
            c = say(s["c"], lang, rng, t)
            if s["d"] < 0:
                unit = "百位" if s["d"] == -2 else "十位"
                return f"加一列把{c}四舍五入到{unit}" if zh else f"add a column with {c} rounded to the nearest {'hundred' if s['d'] == -2 else 'ten'}"
            if zh:
                return f"加一列把{c}取整" if s["d"] == 0 else f"加一列把{c}四舍五入保留{s['d']}位小数"
            return f"add a column with {c} rounded to {s['d']} decimals" if s["d"] else f"add a column with {c} rounded to a whole number"
        if k == "datepart":
            c = say(s["c"], lang, rng, t)
            part = ("年份" if s["fn"] == "YEAR" else "月份") if zh else ("year" if s["fn"] == "YEAR" else "month")
            return rng.choice([f"加一列提取{c}的{part}", f"从{c}里取出{part}放新列"] if zh else [f"add a column with the {part} of {c}", f"extract the {part} from {c} into a new column"])
        c = say(s["c"], lang, rng, t)
        return f"加一列取{c}的前{s['n']}位" if zh else f"add a column with the first {s['n']} characters of {c}"

    def summ(self, s, t, lang):
        return (f"新增公式列{self._name(s, t)}" if lang == "zh" else f"add formula column {self._name(s, t)}")


class SaveInPlace(G.Intent):
    name = "save_in_place"
    stage = 5
    def step(self, s, t, lang): return {"op": "save", "args": {"path": t.file}}
    def text(self, s, t, lang, rng):
        return rng.choice(["改完直接覆盖原文件", "直接保存回原文件（覆盖）", "存回原表，覆盖掉"] if lang == "zh" else
                          ["overwrite the original file", "save it back over the same file"])
    def summ(self, s, t, lang): return "覆盖保存原文件" if lang == "zh" else "overwrite original file"


NEW = [LookupCol(), FilterSheet(), UniqueSheet(), FormulaCol(), SaveInPlace()]
INTENTS2 = G.INTENTS + NEW
WEIGHTS2 = {**G.WEIGHTS, "lookup_column": 3, "filter_to_sheet": 3, "unique_to_sheet": 2, "formula_col": 4, "save_in_place": 1}


# ------------------------------------------------------------------ v3 补强意图（只在 --v3 下进入意图池）
class TextSort(G.Intent):
    """文本列排序（v2 只覆盖了数值/日期列）。"""
    name = "text_sort"
    stage = 3
    def ok(self, t): return bool(t.by_role("cat", "name"))
    def slots(self, t, rng): return {"c": rng.choice(t.by_role("cat", "name")), "desc": rng.random() < 0.3}
    def cf(self, s, t, rng): return {**s, "desc": not s["desc"]}, "排序方向"
    def step(self, s, t, lang): return {"op": "sort", "args": {"by": [{"col": t.h(s["c"]), "desc": s["desc"]}]}}
    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        if lang == "zh":
            return rng.choice([f"按{c}倒序排列", f"{c}反向排一下"] if s["desc"] else [f"按{c}排个序", f"按{c}顺序排列", f"{c}按字母顺序排"])
        return rng.choice([f"sort by {c} Z to A", f"order rows by {c} in reverse"] if s["desc"] else [f"sort by {c}", f"order by {c} alphabetically"])
    def summ(self, s, t, lang): return f"按{t.h(s['c'])}{'降序' if s['desc'] else '升序'}排序" if lang == "zh" else f"sort by {t.h(s['c'])}"


class DeleteEq(G.Intent):
    """按文本取值删除整行（如某个部门/某个人的记录）。"""
    name = "delete_eq"
    def ok(self, t): return bool([c for c in t.by_role("cat") if c["values"]])
    def slots(self, t, rng):
        c = rng.choice([c for c in t.by_role("cat") if c["values"]])
        return {"c": c, "v": rng.choice(c["values"][0 if t.hlang == "zh" else 1])}
    def cf(self, s, t, rng):
        vals = [v for v in s["c"]["values"][0 if t.hlang == "zh" else 1] if v != s["v"]]
        return {**s, "v": rng.choice(vals)}, "删除值"
    def step(self, s, t, lang): return {"op": "delete_rows", "args": {"where": {"col": t.h(s["c"]), "cmp": "==", "value": s["v"]}}}
    def text(self, s, t, lang, rng):
        c, v = say(s["c"], lang, rng, t), s["v"]
        return rng.choice([f"{c}是{v}的记录都删除", f"删除{c}等于“{v}”的行", f"{v}那些行不要了（看{c}）"] if lang == "zh" else
                          [f'delete rows where {c} is "{v}"', f"drop all {v} rows ({c})"])
    def summ(self, s, t, lang): return f"删除{t.h(s['c'])}={s['v']}的行" if lang == "zh" else f"delete {t.h(s['c'])} = {s['v']}"


class CFText(G.Intent):
    """文本条件格式：等于某值 / 包含某词。"""
    name = "cf_text"
    stage = 3
    def ok(self, t): return bool([c for c in t.by_role("cat") if c["values"]] or t.by_role("note"))
    def slots(self, t, rng):
        cats = [c for c in t.by_role("cat") if c["values"]]
        if cats and (not t.by_role("note") or rng.random() < 0.5):
            c = rng.choice(cats)
            return {"c": c, "cmp": "==", "v": rng.choice(c["values"][0 if t.hlang == "zh" else 1]), "color": rng.choice(list(COLORS3)), "row": rng.random() < 0.4}
        return {"c": rng.choice(t.by_role("note")), "cmp": "contains", "v": rng.choice(["退款", "加急", "已核对", "pending", "ok", "首付"]),
                "color": rng.choice(list(COLORS3)), "row": rng.random() < 0.4}
    def cf(self, s, t, rng):
        return ({**s, "row": not s["row"]}, "作用范围") if rng.random() < 0.5 else ({**s, "color": rng.choice([c for c in COLORS3 if c != s["color"]])}, "颜色")
    def step(self, s, t, lang):
        style = {"fill_color": s["color"]} if s["color"] == "FFFF00" else {"font_color": s["color"]}
        return {"op": "conditional_format", "args": {"where": {"col": t.h(s["c"]), "cmp": s["cmp"], "value": s["v"]},
                                                     "apply_to": "row" if s["row"] else "cell", "style": style}}
    def text(self, s, t, lang, rng):
        c, v = say(s["c"], lang, rng, t), s["v"]
        zc, ec = COLORS3[s["color"]]
        if lang == "zh":
            cond = f"{c}是“{v}”的" if s["cmp"] == "==" else f"{c}里含“{v}”的"
            return cond + ("整行" if s["row"] else "") + (f"标{zc}" if s["color"] != "FFFF00" else "标黄")
        cond = f'{c} equal to "{v}"' if s["cmp"] == "==" else f'{c} containing "{v}"'
        return f"highlight {'whole rows with ' if s['row'] else ''}{cond} in {ec}"
    def summ(self, s, t, lang): return f"{t.h(s['c'])}{s['cmp']}{s['v']}标色" if lang == "zh" else f"highlight {t.h(s['c'])} {s['cmp']} {s['v']}"


COLORS3 = {"FF0000": ("红", "red"), "FFFF00": ("黄", "yellow"), "00B050": ("绿", "green")}


class CFRange(G.Intent):
    """区间条件格式：a ≤ x ≤ b。"""
    name = "cf_range"
    stage = 3
    def ok(self, t): return bool(t.num_cols())
    def slots(self, t, rng):
        lo = rng.choice([0, 100, 500, 1000])
        return {"c": rng.choice(t.num_cols()), "lo": lo, "hi": lo + rng.choice([500, 1000, 5000]), "color": rng.choice(list(COLORS3))}
    def cf(self, s, t, rng): return {**s, "hi": s["hi"] + 1000}, "上限"
    def step(self, s, t, lang):
        h = t.h(s["c"])
        style = {"fill_color": s["color"]} if s["color"] == "FFFF00" else {"font_color": s["color"]}
        return {"op": "conditional_format", "args": {"where": {"all": [{"col": h, "cmp": ">=", "value": s["lo"]}, {"col": h, "cmp": "<=", "value": s["hi"]}]},
                                                     "apply_to": "cell", "style": style}}
    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        zc, ec = COLORS3[s["color"]]
        return (rng.choice([f"{c}介于{s['lo']}和{s['hi']}之间的标{zc}", f"{c}在{s['lo']}~{s['hi']}范围内的标{zc}"]) if lang == "zh" else
                f"highlight {c} between {s['lo']} and {s['hi']} in {ec}")
    def summ(self, s, t, lang): return f"{t.h(s['c'])}∈[{s['lo']},{s['hi']}]标色" if lang == "zh" else f"highlight {t.h(s['c'])} in range"


class Arith(G.Intent):
    """两列四则运算的新列（列名由用户给出）。"""
    name = "arith"
    stage = 2
    NAMES = (["合计", "差额", "比率", "结果", "调整后"], ["Total", "Diff", "Ratio", "Result", "Adjusted"])
    def ok(self, t): return len(t.num_cols()) >= 2
    def slots(self, t, rng):
        a, b = rng.sample(t.num_cols(), 2)
        return {"a": a, "b": b, "op": rng.choice("+-*/"), "name": rng.choice(self.NAMES[0 if t.hlang == "zh" else 1])}
    def cf(self, s, t, rng): return {**s, "op": rng.choice([o for o in "+-*/" if o != s["op"]])}, "运算符"
    def step(self, s, t, lang): return {"op": "add_column", "args": {"name": s["name"], "formula": f"=[{t.h(s['a'])}]{s['op']}[{t.h(s['b'])}]"}}
    def text(self, s, t, lang, rng):
        a, b = say(s["a"], lang, rng, t), say(s["b"], lang, rng, t)
        zo = {"+": "加", "-": "减", "*": "乘", "/": "除以"}[s["op"]]
        eo = {"+": "plus", "-": "minus", "*": "times", "/": "divided by"}[s["op"]]
        return (rng.choice([f"加一列{s['name']}，等于{a}{zo}{b}", f"新增列“{s['name']}”= {a} {s['op']} {b}"]) if lang == "zh" else
                rng.choice([f"add a column {s['name']} = {a} {eo} {b}", f"new column {s['name']}: {a} {s['op']} {b}"]))
    def summ(self, s, t, lang): return f"新增列{s['name']}" if lang == "zh" else f"add column {s['name']}"


class GroupMulti(G.Intent):
    """多聚合（同时求和与平均）或多键汇总。"""
    name = "group_multi"
    stage = 4
    def ok(self, t): return bool(t.by_role("cat") and t.num_cols())
    def slots(self, t, rng):
        cats = t.by_role("cat")
        if len(cats) >= 2 and rng.random() < 0.5:
            return {"kind": "keys", "by": rng.sample(cats, 2), "v": rng.choice(t.num_cols())}
        return {"kind": "aggs", "by": [rng.choice(cats)], "v": rng.choice(t.num_cols()), "fns": rng.choice([["sum", "avg"], ["sum", "max"], ["min", "max"]])}
    def step(self, s, t, lang):
        by = [t.h(c) for c in s["by"]]
        aggs = [{"col": t.h(s["v"]), "fn": f} for f in (s.get("fns") or ["sum"])]
        return {"op": "group_aggregate", "args": {"by": by, "aggs": aggs, "output_sheet": ("按" + "和".join(by) + "汇总" if lang == "zh" else " & ".join(by) + " summary")[:31],
                                                   "mode": "formula"}}
    def text(self, s, t, lang, rng):
        v = say(s["v"], lang, rng, t)
        if s["kind"] == "keys":
            a, b = (say(c, lang, rng, t) for c in s["by"])
            return f"按{a}和{b}两个维度汇总{v}，放新表" if lang == "zh" else f"total {v} by {a} and {b} in a new sheet"
        by = say(s["by"][0], lang, rng, t)
        zf = {"sum": "合计", "avg": "平均值", "max": "最大值", "min": "最小值"}
        return (f"按{by}汇总{v}，同时给出" + "和".join(zf[f] for f in s["fns"]) if lang == "zh" else
                f"by {by}, show the " + " and ".join({"sum": "total", "avg": "average", "max": "max", "min": "min"}[f] for f in s["fns"]) + f" of {v}")
    def summ(self, s, t, lang): return "多维汇总" if lang == "zh" else "multi summary"


V3_INTENTS = [TextSort(), DeleteEq(), CFText(), CFRange(), Arith(), GroupMulti()]
V3_WEIGHTS = {"text_sort": 2, "delete_eq": 2, "cf_text": 2, "cf_range": 1, "arith": 2, "group_multi": 2}


# ------------------------------------------------------------------ 歧义（A1–A5）
# 措辞按 split 分开：测试用的说法训练时没见过
PH = {
    "A1": {"train": (["把带括号的{c}删了", "带括号的{c}不要了", "删掉有括号的{c}"], ["delete the {c} with parentheses", "get rid of bracketed {c}"]),
           "test": (["带括号的{c}都去掉", "{c}有括号的全删"], ["remove {c} that have brackets", "drop the parenthesized {c}"])},
    "A2": {"train": None, "test": None},  # 由意图本身的说法生成，只把列称呼换成共享词“金额/amount”
    "A3": {"train": (["把重复的{e}删了", "重复的{e}去掉", "{e}有重复的清理掉"], ["remove the duplicate {e}s", "get rid of repeated {e}s"]),
           "test": (["{e}重复的只留一条", "去掉重复{e}"], ["dedupe the {e}s", "drop duplicated {e}s"])},
    "A4": {"train": (["把空的删了", "空着的行删掉", "没填的都删了"], ["delete the empty ones", "remove the blank rows"]),
           "test": (["空白的去掉", "删除空记录"], ["drop the empty entries", "clear out blank records"])},
    "A5": {"train": (["改完保存一下", "弄完存一下"], ["save it when done", "then save"]),
           "test": (["处理完保存", "最后存起来"], ["save afterwards", "and save the file"])},
}
ENTITY = {"attendance": ("考勤记录", "attendance record"), "invoice": ("发票", "invoice"), "sales": ("订单", "order"), "inventory": ("物料", "item"), "project": ("项目", "project"),
          "payroll": ("员工记录", "employee record"), "receivable": ("客户", "customer"), "ledger": ("记录", "record"),
          "expense": ("报销单", "claim")}


def _clear_extra(t, lang, rng, exclude=()):
    it = rng.choice([x for x in INTENTS2 if x.name not in ("strip_parens", "backup", "save_as", "save_in_place", *exclude) and x.ok(t)])
    s = it.slots(t, rng)
    return it, s


PH_V3 = {  # 收官补强：训练用的更多说法（与人工测试集措辞不重合，生成后另做全文去重核查）
    "A1": (["括号的{c}全部删掉", "{c}带括号的给我去掉", "有括号的{c}清一下", "把含括号的{c}处理掉", "{c}里面带括号的那些删了"],
           ["remove the {c} containing brackets", "clear bracketed {c}", "delete {c} that has parentheses", "lose the {c} with brackets"]),
    "A3": (["把重复的{e}去一下", "{e}重了的删掉", "删除重复{e}", "重复{e}只要一个"],
           ["delete duplicated {e}s", "clean out duplicate {e}s", "dump the repeated {e}s"]),
    "A4": (["空白行删了", "把空着的去掉", "没内容的删除", "删掉空数据", "空的那几条不要了"],
           ["get rid of the blanks", "remove empty stuff", "delete the empty records", "trash the blank lines"]),
    "A5": (["做完存一下", "改好保存", "最后保存一下", "搞定后存盘", "完了存个档"],
           ["save it after", "then save it", "save once finished", "and save at the end"]),
}


def make_amb(t, lang, rng, split):
    """返回 (text, plan, items, sibling_or_None)。sibling = (text2, plan2) 已澄清 → 执行备选解释、无 ambiguities。"""
    kinds = ["A1"] * (bool(t.by_role("note"))) + ["A2"] * 2 * (t.theme in SHARED) + ["A3"] * bool(t.by_role("id")) + ["A4", "A5"]
    kind = rng.choice(kinds)
    ph = PH[kind]["train" if split == "train" else "test"]
    if V3 and split == "train" and kind in PH_V3 and rng.random() < 0.6:
        ph = PH_V3[kind]
    zh = lang == "zh"
    extra = [_clear_extra(t, lang, rng, exclude=("dedupe", "delete_empty")) for _ in range(rng.random() < 0.6)]
    ex_text = "".join(rng.choice(JOIN[lang]) + it.text(s, t, lang, rng) for it, s in extra)

    if kind == "A1":
        c = rng.choice(t.by_role("note"))
        cs = say(c, lang, rng, t)
        text = rng.choice(ph[0] if zh else ph[1]).format(c=cs)
        items = [(G.StripParens(), {"c": c})] + extra
        plan = build_plan(t, items, lang)
        sid = next(s["id"] for s in plan["steps"] if s["op"] == "regex_replace")
        h = t.h(c)
        alt = {"id": sid, "op": "delete_rows", "sheet": t.sheet, "args": {"where": {"col": h, "cmp": "regex", "value": PAREN_RX}}}
        q = (f"“删掉带括号的{h}”是指删除{h}里括号及括号中的文字，还是删除{h}含括号的整行？" if zh else
             f"Remove only the bracketed text inside {h}, or delete the whole rows whose {h} contains brackets?")
        sib_text = rng.choice([f"把{cs}里含括号的整行删掉", f"{cs}带括号的整条记录删除"] if zh else
                              [f"delete the whole rows whose {cs} contains brackets", f"drop every row with brackets in {cs}"])
        sib_items = [(DelRegexRows(), {"c": c})] + extra
    elif kind == "A2":
        a_col, b_col = t.by_role("amount")[:2]
        word = SHARED[t.theme][0 if zh else 1]
        it = rng.choice([G.CFNegative(), G.CFThreshold(), G.DeleteThreshold(), G.Sort()])
        s = it.slots(t, rng)
        s["c"] = a_col
        fake = dict(a_col, zh=word, en=word, say_zh=[word], say_en=[word])
        text = it.text({**s, "c": fake}, t, lang, rng)
        items = [(it, s)] + extra
        plan = build_plan(t, items, lang)
        st = next(x for x in plan["steps"] if x["op"] == it.step(s, t, lang)["op"])
        alt = copy.deepcopy(st)
        _swap_col(alt, t.h(a_col), t.h(b_col))
        q = (f"“{word}”是指{t.h(a_col)}还是{t.h(b_col)}？" if zh else f'By "{word}" do you mean {t.h(a_col)} or {t.h(b_col)}?')
        s2 = {**s, "c": b_col}
        sib_text = it.text({**s2, "c": dict(b_col, say_zh=[t.h(b_col)], say_en=[t.h(b_col)], zh=t.h(b_col), en=t.h(b_col))}, t, lang, rng)
        sib_items = [(it, s2)] + extra
        sid = st["id"]
    elif kind == "A3":
        idc = t.by_role("id")[0]
        e = ENTITY[t.theme][0 if zh else 1]
        text = rng.choice(ph[0] if zh else ph[1]).format(e=e)
        items = [(G.Dedupe(), {"key": idc})] + extra
        plan = build_plan(t, items, lang)
        st = next(x for x in plan["steps"] if x["op"] == "dedupe")
        alt = {"id": st["id"], "op": "dedupe", "sheet": t.sheet, "args": {}}
        h = t.h(idc)
        q = (f"按{h}判断重复（同一{h}只留第一条），还是整行完全相同才算重复？" if zh else
             f"Treat rows as duplicates when {h} matches (keep the first), or only when the entire row is identical?")
        sib_text = rng.choice(["只删除整行完全一样的重复行", "完全相同的行才算重复，删掉"] if zh else
                              ["remove only rows that are completely identical", "dedupe on the whole row"])
        sib_items = [(G.Dedupe(), {"key": None})] + extra
        sid = st["id"]
    elif kind == "A4":
        text = rng.choice(ph[0] if zh else ph[1])
        items = [(DelAllEmpty(), {})] + extra
        plan = build_plan(t, items, lang)
        st = next(x for x in plan["steps"] if x["op"] == "delete_rows")
        c = rng.choice(t.by_role("note") or t.cols)
        alt = {"id": st["id"], "op": "delete_rows", "sheet": t.sheet, "args": {"where": {"col": t.h(c), "cmp": "is_empty"}}}
        q = (f"是删除整行全空的行，还是删除{t.h(c)}为空的行？" if zh else f"Delete rows that are entirely blank, or rows where {t.h(c)} is empty?")
        sib_items = [(G.DeleteEmpty(), {"c": c})] + extra
        sib_text = G.DeleteEmpty().text({"c": c}, t, lang, rng)
        sid = st["id"]
    else:  # A5
        extra = [x for x in extra if x[0].name != "split"]
        if not extra:
            extra = [_clear_extra(t, lang, rng, exclude=("split",))]
            ex_text = rng.choice(JOIN[lang]) + extra[0][0].text(extra[0][1], t, lang, rng)
        lead = ex_text.lstrip("，, ;；").removeprefix("然后").removeprefix("再").removeprefix("then ").removeprefix("and ")
        text = lead + rng.choice(JOIN[lang]) + rng.choice(ph[0] if zh else ph[1])
        ex_text = ""
        items = list(extra)
        plan = build_plan(t, items, lang)
        st = next((x for x in plan["steps"] if x["op"] == "save"), None)
        if st is None:  # 例如只有拆分文件，没有默认另存 → 这类组合不构成 A5
            raise ValueError("A5 without save step")
        alt = {"id": st["id"], "op": "save", "args": {"path": t.file}}
        q = (f"要直接覆盖原文件 {t.file} 吗？默认另存为 {st['args']['path']}，原文件不动。" if zh else
             f"Overwrite the original {t.file}? By default it is saved as {st['args']['path']} and the original is untouched.")
        sib_text = lead + rng.choice(JOIN[lang]) + SaveInPlace().text({}, t, lang, rng)
        sib_items = items + [(SaveInPlace(), {})]
        sid = st["id"]
    if kind != "A5":
        text += ex_text
        sib_text += ex_text
    plan["ambiguities"] = [{"step_id": sid, "question": q, "alternatives": [alt]}]
    sib = (sib_text, build_plan(t, sib_items, lang)) if rng.random() < 0.7 else None
    return text, plan, kind, sib


def _swap_col(step, a, b):
    def rec(o):
        if isinstance(o, dict):
            for k, v in o.items():
                if k == "col" and v == a:
                    o[k] = b
                else:
                    rec(v)
        elif isinstance(o, list):
            for x in o:
                rec(x)
    rec(step["args"])


class DelRegexRows(G.Intent):
    name = "del_regex_rows"
    def step(self, s, t, lang): return {"op": "delete_rows", "args": {"where": {"col": t.h(s["c"]), "cmp": "regex", "value": PAREN_RX}}}
    def summ(self, s, t, lang): return f"删除{t.h(s['c'])}含括号的整行" if lang == "zh" else f"delete rows whose {t.h(s['c'])} has brackets"


class DelAllEmpty(G.Intent):
    name = "del_all_empty"
    def step(self, s, t, lang):
        return {"op": "delete_rows", "args": {"where": {"all": [{"col": t.h(c), "cmp": "is_empty"} for c in t.cols]}}}
    def summ(self, s, t, lang): return "删除整行全空的行" if lang == "zh" else "delete fully blank rows"


# ------------------------------------------------------------------ 执行校验
def execute2(t, plan, seed):
    d = Path(tempfile.mkdtemp(prefix="synth2_"))
    try:
        t.make_xlsx(d / t.file, random.Random(seed))
        wb = load_workbook(d / t.file)
        prof = profile_sheet(wb[t.sheet])
        errs = validate(plan, prof)
        if errs:
            raise ValueError(errs)
        p2 = copy.deepcopy(plan)
        p2.pop("ambiguities", None)
        Executor(Workspace(d)).run(p2)
        text = prof.to_prompt() + "\n" + profile_sheet(wb[t.dim_sheet]).to_prompt()
        return text, digest(d)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def _row(split, t, lang, seed, text, plan, pair, **kw):
    return {"split": split, "gen": 2, "theme": t.theme, "lang": lang, "file": t.file, "seed": seed, "hlang": t.hlang,
            "text": text, "plan": plan, "pair": pair, **kw}


def worker(job):
    split, n, themes, seed, noise_p, amb_p = job
    rng = random.Random(seed)
    rows, st = [], {"ok": 0, "exec_fail": 0, "cf_pairs": 0, "cf_no_effect": 0, "ambiguous": 0, "amb_pairs": 0, "noised": 0}
    fails = []
    while st["ok"] < n:
        theme = rng.choice(themes)
        t = Table2(theme, "zh" if rng.random() < 0.65 else "en", rng)
        lang = "zh" if rng.random() < 0.6 else "en"
        xseed = rng.randint(0, 10**9)
        nrng = random.Random(xseed ^ 0x5EED)
        do_noise = rng.random() < noise_p

        def noisy(text):
            if not do_noise:
                return text, []
            return perturb(text, lang, random.Random(nrng.random()), "train", v3=V3)

        if rng.random() < amb_p:
            kind = "?"
            try:  # 注意：StopIteration 若逃出 worker 会被 Pool.map 当成“迭代结束”而静默返回 None
                text, plan, kind, sib = make_amb(t, lang, rng, split)
                plan["version"] = "0.2"
                prof, dg = execute2(t, plan, xseed)
                pid = None
                srow = None
                if sib:
                    sib[1]["version"] = "0.2"
                    _, dg2 = execute2(t, sib[1], xseed)
                    if dg2 != dg:
                        pid = hashlib.sha1(f"{seed}{xseed}{text}".encode()).hexdigest()[:12]
                        t2, nz2 = noisy(sib[0])
                        srow = _row(split, t, lang, xseed, t2, sib[1], pid, profile=prof, intervened=f"amb.{kind}.clarified", noise=nz2)
                        st["amb_pairs"] += 1
                    else:
                        st["cf_no_effect"] += 1
                t1, nz = noisy(text)
                rows.append(_row(split, t, lang, xseed, t1, plan, pid, profile=prof, amb_kind=kind, noise=nz))
                if srow:
                    rows.append(srow)
                st["ambiguous"] += 1
                st["ok"] += 1
                st["noised"] += bool(nz)
            except Exception as e:  # noqa: BLE001
                st["exec_fail"] += 1
                if len(fails) < 3:
                    fails.append(f"amb {kind}: {e}"[:300])
            continue
        pool = [it for it in (INTENTS2 + V3_INTENTS if V3 else INTENTS2) if it.ok(t)]
        k = rng.choice([1, 1, 2, 2, 2, 3, 3, 4])
        chosen = []
        for _ in range(k):
            cand = [it for it in pool if it not in chosen and not (it.name in ("save_as", "save_in_place") and
                                                                    any(c.name in ("save_as", "save_in_place") for c in chosen))]
            chosen.append(rng.choices(cand, weights=[{**WEIGHTS2, **V3_WEIGHTS}[c.name] for c in cand])[0])
        items = [(it, it.slots(t, rng)) for it in chosen]
        plan = build_plan(t, items, lang)
        plan["version"] = "0.2"
        try:
            prof, dg = execute2(t, plan, xseed)
        except Exception as e:  # noqa: BLE001
            st["exec_fail"] += 1
            if len(fails) < 3:
                fails.append(f"{[i.name for i in chosen]}: {e}"[:300])
            continue
        text = utter(t, items, lang, random.Random(xseed))
        pid, sib = None, None
        cfable = [i for i, (it, s) in enumerate(items) if it.cf(s, t, random.Random(0)) is not None]
        if cfable and rng.random() < 0.5:
            i = rng.choice(cfable)
            it, s = items[i]
            ns, var = it.cf(s, t, rng)
            items2 = list(items)
            items2[i] = (it, ns)
            plan2 = build_plan(t, items2, lang)
            plan2["version"] = "0.2"
            text2 = utter(t, items2, lang, random.Random(xseed))
            try:
                _, dg2 = execute2(t, plan2, xseed)
                if dg2 == dg:
                    st["cf_no_effect"] += 1
                else:
                    pid = hashlib.sha1(f"{seed}{xseed}{text}".encode()).hexdigest()[:12]
                    t2, nz2 = noisy(text2)
                    sib = _row(split, t, lang, xseed, t2, plan2, pid, profile=prof, intervened=f"{it.name}.{var}", noise=nz2)
                    st["cf_pairs"] += 1
            except Exception:  # noqa: BLE001
                st["exec_fail"] += 1
        t1, nz = noisy(text)
        rows.append(_row(split, t, lang, xseed, t1, plan, pid, profile=prof, noise=nz,
                         **({"intervened": sib["intervened"]} if sib else {})))
        if sib:
            rows.append(sib)
        st["ok"] += 1
        st["noised"] += bool(nz)
    return rows, st, fails


def run(split, n, themes, seed, workers, noise_p, amb_p):
    per = -(-n // workers)
    jobs = [(split, min(per, n - i * per), themes, seed * 1000 + i, noise_p, amb_p) for i in range(workers) if n - i * per > 0]
    with mp.Pool(len(jobs)) as pool:
        res = pool.map(worker, jobs)
    rows = [r for rs, _, _ in res for r in rs]
    st = {}
    for _, s, f in res:
        for k, v in s.items():
            st[k] = st.get(k, 0) + v
    fails = [x for _, _, f in res for x in f][:5]
    return rows, st, fails


def noisy_copy(rows, seed):
    """把测试行的指令换成“留出族”扰动版本（金标不变）。"""
    out = []
    rng = random.Random(seed)
    for r in rows:
        t, nz = perturb(r["text"], r["lang"], random.Random(rng.random()), "heldout", strength=1.5)
        out.append({**r, "split": "test_v2_noisy", "text": t, "noise": nz})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=24000)
    ap.add_argument("--n-test", type=int, default=400)
    ap.add_argument("--workers", type=int, default=max(1, (mp.cpu_count() or 2) - 2))
    ap.add_argument("--noise", type=float, default=0.45)
    ap.add_argument("--amb", type=float, default=0.2)
    ap.add_argument("--seed", type=int, default=11)
    ap.add_argument("--out", default=str(ROOT / "data/synth_v2"))
    ap.add_argument("--v3", action="store_true", help="收官补强模式（只生成训练集）")
    a = ap.parse_args()
    global V3
    V3 = a.v3
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    stats = {}
    sets = {}
    jobs = [("train", a.n, TRAIN_THEMES_V3 if V3 else TRAIN_THEMES, a.seed, a.noise)]
    if not V3:
        jobs += [("test_v2_id", a.n_test, TRAIN_THEMES, a.seed + 1, 0.0), ("test_v2_ood", a.n_test, OOD_THEMES, a.seed + 2, 0.0)]
    for name, n, themes, seed, nz in jobs:
        rows, st, fails = run(name, n, themes, seed, a.workers, nz, a.amb)
        sets[name] = rows
        stats[name] = {**st, "rows": len(rows), "fails": fails}
        print(name, json.dumps(stats[name], ensure_ascii=False), f"{time.time() - t0:.0f}s", flush=True)
    if not V3:
        seen = {r["text"] for r in sets["train"]}
        sets["test_v2_id"] = [r for r in sets["test_v2_id"] if r["text"] not in seen]
        sets["test_v2_noisy"] = noisy_copy(sets["test_v2_ood"], a.seed + 3)
    for name, rows in sets.items():
        with (out / f"{name}.jsonl").open("w", encoding="utf-8") as f:
            for r in rows:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
        stats.setdefault(name, {})["rows"] = len(rows)
    (out / "stats.json").write_text(json.dumps(stats, ensure_ascii=False, indent=1), encoding="utf-8")
    print("done", f"{time.time() - t0:.0f}s")


if __name__ == "__main__":
    main()
