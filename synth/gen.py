"""反事实因果语料合成 + 执行校验。

每条样本 = (表格画像, 自然语言指令[zh|en], 金标 Action Plan)。金标由程序直接构造（不经解析器），
随后在随机生成的真实 .xlsx 上用第一层执行器跑通才入库。

反事实对：对样本中的某一个意图只干预一个槽位（阈值/颜色/列/动作/方向…），得到 sibling；
两者都必须执行成功，且执行结果摘要（单元格值 + 条件格式 + 数字格式 + 产出文件）必须不同，
否则说明“干预没有因果效应”，整对丢弃。训练时成对出现，迫使模型依赖因果变量而非表面词。

    python synth/gen.py --n 3000 --out data/synth
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import random
import re
import shutil
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.formula import ArrayFormula

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from tabularedge.executor import Executor, Workspace  # noqa: E402
from tabularedge.protocol import profile_sheet, validate  # noqa: E402

PAREN_RX = r"[\(（][^\)）]*[\)）]"

# ------------------------------------------------------------------ 表格主题
# 每列：zh/en 表头、指令中的说法、类型、角色
def C(zh, en, t, role, say_zh=(), say_en=(), values=None):
    return {"zh": zh, "en": en, "type": t, "role": role, "say_zh": [zh, *say_zh],
            "say_en": [en.lower(), *say_en], "values": values}


DEPTS = (["销售部", "财务部", "研发部", "市场部", "行政部"], ["Sales", "Finance", "R&D", "Marketing", "Admin"])
REGIONS = (["华东", "华南", "华北", "西南"], ["East", "South", "North", "West"])
PRODUCTS = (["笔记本", "显示器", "键盘", "鼠标"], ["Laptop", "Monitor", "Keyboard", "Mouse"])
TYPES = (["差旅", "餐饮", "办公用品", "交通"], ["Travel", "Meals", "Supplies", "Transport"])
WH = (["一号仓", "二号仓", "三号仓"], ["WH-1", "WH-2", "WH-3"])
NAMES = (["张伟", "王芳", "李娜", "刘洋", "陈静", "赵磊"], ["Alice", "Bob", "Carol", "Dave", "Erin", "Frank"])

THEMES = {
    "ledger": [C("日期", "Date", "date", "date", ("时间",), ("day",)),
               C("部门", "Department", "text", "cat", ("部门名称",), ("dept", "departments"), DEPTS),
               C("金额", "Amount", "num", "amount", ("发生额",), ("amounts", "money")),
               C("备注", "Notes", "text", "note", ("说明",), ("note", "remarks", "comments"))],
    "sales": [C("订单号", "Order ID", "text", "id", ("单号",), ("order number",)),
              C("销售员", "Sales Rep", "text", "cat", ("业务员",), ("rep", "salesperson"), NAMES),
              C("区域", "Region", "text", "cat", ("大区",), ("area",), REGIONS),
              C("产品", "Product", "text", "cat", ("品类",), ("item",), PRODUCTS),
              C("数量", "Qty", "num", "qty", ("件数",), ("quantity", "units")),
              C("单价", "Unit Price", "num", "price", ("价格",), ("price",)),
              C("备注", "Remarks", "text", "note", (), ("notes", "comments"))],
    "expense": [C("员工", "Employee", "text", "cat", ("姓名", "报销人"), ("name", "staff"), NAMES),
                C("部门", "Dept", "text", "cat", (), ("department",), DEPTS),
                C("费用类型", "Category", "text", "cat", ("类别",), ("expense type", "type"), TYPES),
                C("报销金额", "Claim Amount", "num", "amount", ("金额",), ("amount", "claim")),
                C("说明", "Description", "text", "note", ("备注",), ("notes", "memo"))],
    "inventory": [C("物料编码", "SKU", "text", "id", ("编码",), ("item code",)),
                  C("物料名称", "Item Name", "text", "name", ("名称",), ("name",)),
                  C("仓库", "Warehouse", "text", "cat", ("库位",), ("location",), WH),
                  C("库存数量", "Stock Qty", "num", "amount", ("库存", "数量"), ("stock", "quantity")),
                  C("备注", "Comment", "text", "note", (), ("notes", "remarks"))],
    # ---- 以下两个主题只用于测试集（检验对未见表结构的泛化）
    "payroll": [C("工号", "Emp No", "text", "id", (), ("employee number",)),
                C("姓名", "Name", "text", "name", (), ("employee",)),
                C("部门", "Division", "text", "cat", (), ("department", "dept"), DEPTS),
                C("实发工资", "Net Pay", "num", "amount", ("工资",), ("salary", "pay")),
                C("备注", "Memo", "text", "note", (), ("notes", "remarks"))],
    "receivable": [C("客户", "Customer", "text", "cat", ("客户名称",), ("client",), NAMES),
                   C("区域", "Area", "text", "cat", (), ("region",), REGIONS),
                   C("应收金额", "Receivable", "num", "amount", ("应收", "金额"), ("amount", "amount due")),
                   C("账龄天数", "Days Overdue", "num", "num", ("账龄",), ("aging", "days")),
                   C("说明", "Note", "text", "note", ("备注",), ("notes", "comment"))],
}
TEST_THEMES = {"payroll", "receivable"}


class Table:
    def __init__(self, theme, hlang, rng):
        self.theme, self.hlang = theme, hlang
        self.cols = THEMES[theme]
        self.sheet = {"zh": {"ledger": "流水", "sales": "销售明细", "expense": "报销", "inventory": "库存",
                             "payroll": "工资表", "receivable": "应收"}, "en": {}}[hlang].get(theme, "Sheet1")
        stem = rng.choice({"zh": ["数据", "明细", "台账", f"2026年{rng.randint(1, 12)}月{theme}"],
                           "en": ["data", "report", f"{theme}_2026"]}[hlang])
        self.file = f"{stem}.xlsx"

    def h(self, c):
        return c[self.hlang]

    def by_role(self, *roles):
        return [c for c in self.cols if c["role"] in roles]

    def num_cols(self):
        return [c for c in self.cols if c["type"] == "num"]

    def headers(self):
        return [self.h(c) for c in self.cols]

    def make_xlsx(self, path, rng, n=24):
        wb = Workbook()
        ws = wb.active
        ws.title = self.sheet
        ws.append(self.headers())
        for i in range(n):
            row = []
            for c in self.cols:
                r = c["role"]
                if c["values"]:
                    v = rng.choice(c["values"][0 if self.hlang == "zh" else 1])
                elif r == "date":
                    v = f"2026-{rng.randint(1, 12):02d}-{rng.randint(1, 28):02d}"
                elif r == "id":
                    v = f"{c['en'][:2].upper()}-{1000 + i}"
                elif r == "name":
                    v = rng.choice(["螺丝", "垫片", "轴承", "Bolt", "Washer", "Bearing"])
                elif r == "note":
                    v = rng.choice(["", "正常", "客户A（首付）", "退款(重复扣款)", "  报销（打车） ", "已核对",
                                    "pending (check)", "ok", "加急 ", None])
                elif r == "amount":
                    v = rng.choice([rng.randint(-500, -1), rng.randint(0, 20000), rng.randint(100, 5000)])
                elif r == "qty":
                    v = rng.randint(1, 50)
                elif r == "price":
                    v = round(rng.uniform(5, 3000), 2)
                else:
                    v = rng.randint(0, 200)
                if rng.random() < 0.05 and r not in ("id",):
                    v = None
                row.append(v)
            ws.append(row)
        if rng.random() < 0.5:  # 制造重复行
            ws.append([c.value for c in ws[2]])
        wb.save(path)


def say(c, lang, rng, table):
    """指令里怎么称呼某列。表头语言 == 指令语言时多半直接用表头；否则用另一语言的说法。"""
    if lang == table.hlang and rng.random() < 0.7:
        return table.h(c)
    return rng.choice(c["say_zh"] if lang == "zh" else c["say_en"])


# ------------------------------------------------------------------ 意图
COLORS = {"FF0000": ("红", "red"), "FFFF00": ("黄", "yellow"), "00B050": ("绿", "green")}


class Intent:
    name = ""
    stage = 1  # 0 备份 1 清洗 2 加列 3 格式 4 汇总/拆分 5 保存

    def ok(self, t): return True
    def slots(self, t, rng): return {}
    def cf(self, s, t, rng): return None  # 返回 (新槽位, 干预描述) 或 None
    def step(self, s, t, lang): raise NotImplementedError
    def text(self, s, t, lang, rng): raise NotImplementedError
    def summ(self, s, t, lang): raise NotImplementedError


class StripParens(Intent):
    name = "strip_parens"
    def ok(self, t): return bool(t.by_role("note"))
    def slots(self, t, rng): return {"c": rng.choice(t.by_role("note"))}
    def step(self, s, t, lang): return {"op": "regex_replace", "args": {"col": t.h(s["c"]), "pattern": PAREN_RX, "repl": ""}}
    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        return rng.choice([f"把{c}里括号及括号里的内容去掉", f"删掉{c}中括号里的文字", f"{c}里带括号的部分清掉",
                           f"去掉{c}的括号内容", f"{c}里面的（…）都删了，括号也不要"] if lang == "zh" else
                          [f"remove the parentheses and the text inside them from {c}", f"strip the bracketed bits out of {c}",
                           f"get rid of anything in brackets in the {c} column", f"delete parenthetical text in {c}"])
    def summ(self, s, t, lang): return f"删除{t.h(s['c'])}中的括号内容" if lang == "zh" else f"strip parentheses in {t.h(s['c'])}"


class Trim(Intent):
    name = "trim"
    def ok(self, t): return any(c["type"] == "text" for c in t.cols)
    def slots(self, t, rng):
        tc = [c for c in t.cols if c["type"] == "text"]
        return {"cols": [rng.choice(tc)] if rng.random() < 0.6 else tc}
    def step(self, s, t, lang): return {"op": "trim", "args": {"cols": [t.h(c) for c in s["cols"]]}}
    def text(self, s, t, lang, rng):
        if len(s["cols"]) == 1:
            c = say(s["cols"][0], lang, rng, t)
            return rng.choice([f"把{c}的多余空格去掉", f"{c}前后的空格清理一下"] if lang == "zh" else
                              [f"trim the spaces in {c}", f"clean up extra whitespace in {c}"])
        return rng.choice(["所有文字列的多余空格都去掉", "把文本里的空格清理干净"] if lang == "zh" else
                          ["trim whitespace in all text columns", "clean up extra spaces everywhere"])
    def summ(self, s, t, lang): return ("去除空格：" if lang == "zh" else "trim: ") + "、".join(t.h(c) for c in s["cols"])


class DeleteEmpty(Intent):
    name = "delete_empty"
    def slots(self, t, rng): return {"c": rng.choice(t.cols)}
    def cf(self, s, t, rng):
        others = [c for c in t.cols if c is not s["c"]]
        return {"c": rng.choice(others)}, "列"
    def step(self, s, t, lang): return {"op": "delete_rows", "args": {"where": {"col": t.h(s["c"]), "cmp": "is_empty"}}}
    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        return rng.choice([f"删掉{c}为空的行", f"{c}没填的行都删了", f"把{c}空着的记录去掉"] if lang == "zh" else
                          [f"delete rows where {c} is empty", f"drop rows with a blank {c}", f"remove any row missing {c}"])
    def summ(self, s, t, lang): return f"删除{t.h(s['c'])}为空的行" if lang == "zh" else f"delete rows with empty {t.h(s['c'])}"


class DeleteThreshold(Intent):
    name = "delete_threshold"
    def ok(self, t): return bool(t.num_cols())
    def slots(self, t, rng): return {"c": rng.choice(t.num_cols()), "cmp": rng.choice(["<", ">"]), "v": rng.choice([0, 10, 100, 1000, 5000])}
    def cf(self, s, t, rng):
        if rng.random() < 0.5:
            return {**s, "cmp": {"<": ">", ">": "<"}[s["cmp"]]}, "比较方向"
        return {**s, "v": rng.choice([v for v in [0, 10, 100, 1000, 5000] if v != s["v"]])}, "阈值"
    def step(self, s, t, lang): return {"op": "delete_rows", "args": {"where": {"col": t.h(s["c"]), "cmp": s["cmp"], "value": s["v"]}}}
    def text(self, s, t, lang, rng):
        c, v = say(s["c"], lang, rng, t), s["v"]
        if lang == "zh":
            w = rng.choice(["小于", "低于", "不到"]) if s["cmp"] == "<" else rng.choice(["大于", "超过", "高于"])
            return rng.choice([f"把{c}{w}{v}的行删了", f"删除{c}{w}{v}的记录", f"{c}{w}{v}的都不要了"])
        w = rng.choice(["less than", "under", "below"]) if s["cmp"] == "<" else rng.choice(["greater than", "over", "above"])
        return rng.choice([f"delete rows where {c} is {w} {v}", f"drop everything with {c} {w} {v}", f"remove rows with {c} {w} {v}"])
    def summ(self, s, t, lang): return f"删除{t.h(s['c'])}{s['cmp']}{s['v']}的行" if lang == "zh" else f"delete rows with {t.h(s['c'])} {s['cmp']} {s['v']}"


class Dedupe(Intent):
    name = "dedupe"
    def slots(self, t, rng):
        ids = t.by_role("id")
        return {"key": ids[0] if ids and rng.random() < 0.5 else None}
    def step(self, s, t, lang): return {"op": "dedupe", "args": {"cols": [t.h(s["key"])]} if s["key"] else {}}
    def text(self, s, t, lang, rng):
        if s["key"]:
            c = say(s["key"], lang, rng, t)
            return rng.choice([f"按{c}去重", f"{c}重复的只留一条"] if lang == "zh" else [f"dedupe by {c}", f"remove duplicate {c} rows, keep the first"])
        return rng.choice(["删除重复行", "去重", "重复的记录去掉"] if lang == "zh" else ["remove duplicate rows", "dedupe the sheet", "get rid of duplicates"])
    def summ(self, s, t, lang): return "去重" if lang == "zh" else "remove duplicates"


class FillEmpty(Intent):
    name = "fill_empty"
    def slots(self, t, rng):
        c = rng.choice(t.cols)
        return {"c": c, "v": 0 if c["type"] == "num" else rng.choice(["无", "N/A", "待补"])}
    def step(self, s, t, lang): return {"op": "fill_empty", "args": {"col": t.h(s["c"]), "value": s["v"]}}
    def text(self, s, t, lang, rng):
        c, v = say(s["c"], lang, rng, t), s["v"]
        return rng.choice([f"{c}为空的填{v}", f"把{c}的空白补成{v}", f"{c}没有值的填上“{v}”"] if lang == "zh" else
                          [f"fill empty {c} with {v}", f"put {v} in blank {c} cells", f"replace missing {c} with {v}"])
    def summ(self, s, t, lang): return f"{t.h(s['c'])}空值填{s['v']}" if lang == "zh" else f"fill empty {t.h(s['c'])} with {s['v']}"


class Sort(Intent):
    name = "sort"
    stage = 3
    def ok(self, t): return bool(t.num_cols())
    def slots(self, t, rng): return {"c": rng.choice(t.num_cols() + t.by_role("date")), "desc": rng.random() < 0.6}
    def cf(self, s, t, rng): return {**s, "desc": not s["desc"]}, "排序方向"
    def step(self, s, t, lang): return {"op": "sort", "args": {"by": [{"col": t.h(s["c"]), "desc": s["desc"]}]}}
    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        if lang == "zh":
            d = rng.choice(["从大到小", "降序", "倒序"]) if s["desc"] else rng.choice(["从小到大", "升序"])
            return rng.choice([f"按{c}{d}排序", f"按照{c}{d}排一下", f"{c}{d}排列"])
        d = rng.choice(["descending", "largest first", "high to low"]) if s["desc"] else rng.choice(["ascending", "smallest first", "low to high"])
        return rng.choice([f"sort by {c} {d}", f"order the rows by {c}, {d}"])
    def summ(self, s, t, lang): return f"按{t.h(s['c'])}{'降序' if s['desc'] else '升序'}排序" if lang == "zh" else f"sort by {t.h(s['c'])} {'desc' if s['desc'] else 'asc'}"


class AddProduct(Intent):
    name = "add_product"
    stage = 2
    def ok(self, t): return bool(t.by_role("qty") and t.by_role("price"))
    def slots(self, t, rng): return {"name": rng.choice(["销售额", "金额", "小计"]) if t.hlang == "zh" else rng.choice(["Revenue", "Total", "Line Total"])}
    def step(self, s, t, lang):
        q, p = t.h(t.by_role("qty")[0]), t.h(t.by_role("price")[0])
        return {"op": "add_column", "args": {"name": s["name"], "formula": f"=[{q}]*[{p}]"}}
    def text(self, s, t, lang, rng):
        q, p = say(t.by_role("qty")[0], lang, rng, t), say(t.by_role("price")[0], lang, rng, t)
        return rng.choice([f"加一列{s['name']}，等于{q}乘{p}", f"新增“{s['name']}”列 = {q}×{p}"] if lang == "zh" else
                          [f"add a {s['name']} column that is {q} times {p}", f"new column {s['name']} = {q} * {p}"])
    def summ(self, s, t, lang): return f"新增列{s['name']}" if lang == "zh" else f"add column {s['name']}"


class CFNegative(Intent):
    name = "cf_negative"
    stage = 3
    def ok(self, t): return bool(t.by_role("amount"))
    def slots(self, t, rng): return {"c": rng.choice(t.by_role("amount")), "color": "FF0000", "row": rng.random() < 0.25}
    def cf(self, s, t, rng):
        if rng.random() < 0.5:
            return {**s, "color": rng.choice(["FFFF00", "00B050"])}, "颜色"
        return {**s, "row": not s["row"]}, "作用范围"
    def step(self, s, t, lang):
        style = {"font_color": s["color"]} if s["color"] != "FFFF00" else {"fill_color": s["color"]}
        return {"op": "conditional_format", "args": {"where": {"col": t.h(s["c"]), "cmp": "<", "value": 0},
                                                     "apply_to": "row" if s["row"] else "cell", "style": style}}
    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        zh_c, en_c = COLORS[s["color"]]
        if lang == "zh":
            act = f"标{zh_c}" if s["color"] != "FFFF00" else rng.choice(["标黄", "填黄色底"])
            scope = "整行" if s["row"] else ""
            return rng.choice([f"{c}是负数的{scope}{act}", f"把{c}为负的{scope}{act}", f"{c}小于0的{scope}{act}", f"负的{c}{scope}给我{act}"])
        scope = "the whole row " if s["row"] else ""
        act = f"in {en_c}" if s["color"] != "FFFF00" else "with a yellow fill"
        return rng.choice([f"highlight {scope}negative {c} {act}".replace("negative the whole row", "the whole row for negative"),
                           f"make negative {c} {en_c}" + (" across the row" if s["row"] else ""),
                           f"mark {scope}rows where {c} is below zero {act}" if s["row"] else f"color {c} below zero {en_c}"])
    def summ(self, s, t, lang): return f"{t.h(s['c'])}负数标{COLORS[s['color']][0]}" if lang == "zh" else f"negative {t.h(s['c'])} -> {COLORS[s['color']][1]}"


class CFThreshold(Intent):
    name = "cf_threshold"
    stage = 3
    def ok(self, t): return bool(t.num_cols())
    def slots(self, t, rng): return {"c": rng.choice(t.num_cols()), "cmp": rng.choice([">", "<"]), "v": rng.choice([10, 100, 1000, 5000, 10000]), "color": rng.choice(list(COLORS))}
    def cf(self, s, t, rng):
        k = rng.choice(["cmp", "v", "color"])
        if k == "cmp": return {**s, "cmp": {"<": ">", ">": "<"}[s["cmp"]]}, "比较方向"
        if k == "v": return {**s, "v": rng.choice([v for v in [10, 100, 1000, 5000, 10000] if v != s["v"]])}, "阈值"
        return {**s, "color": rng.choice([c for c in COLORS if c != s["color"]])}, "颜色"
    def step(self, s, t, lang):
        style = {"fill_color": s["color"]} if s["color"] == "FFFF00" else {"font_color": s["color"]}
        return {"op": "conditional_format", "args": {"where": {"col": t.h(s["c"]), "cmp": s["cmp"], "value": s["v"]}, "apply_to": "cell", "style": style}}
    def text(self, s, t, lang, rng):
        c, v = say(s["c"], lang, rng, t), s["v"]
        zh_c, en_c = COLORS[s["color"]]
        if lang == "zh":
            w = rng.choice(["大于", "超过", "高于"]) if s["cmp"] == ">" else rng.choice(["小于", "低于", "不到"])
            vv = rng.choice([str(v), f"{v // 10000}万"]) if v >= 10000 and v % 10000 == 0 else str(v)
            return rng.choice([f"{c}{w}{vv}的标{zh_c}", f"把{c}{w}{vv}的{'填黄色' if s['color'] == 'FFFF00' else '标成' + zh_c + '色'}"])
        w = rng.choice(["over", "above", "greater than"]) if s["cmp"] == ">" else rng.choice(["under", "below", "less than"])
        return rng.choice([f"highlight {c} {w} {v} in {en_c}", f"make {c} values {w} {v} {en_c}"])
    def summ(self, s, t, lang): return f"{t.h(s['c'])}{s['cmp']}{s['v']}标{COLORS[s['color']][0]}" if lang == "zh" else f"{t.h(s['c'])}{s['cmp']}{s['v']} -> {COLORS[s['color']][1]}"


class NumFormat(Intent):
    name = "number_format"
    stage = 3
    FM = {"#,##0.00": ("千分位保留两位小数", "thousands separators with 2 decimals"), "0.00%": ("百分比格式", "percentage format"),
          "#,##0": ("千分位取整", "thousands separators, no decimals")}
    def ok(self, t): return bool(t.num_cols())
    def slots(self, t, rng): return {"c": rng.choice(t.num_cols()), "f": rng.choice(list(self.FM))}
    def cf(self, s, t, rng): return {**s, "f": rng.choice([f for f in self.FM if f != s["f"]])}, "格式"
    def step(self, s, t, lang): return {"op": "number_format", "args": {"col": t.h(s["c"]), "format": s["f"]}}
    def text(self, s, t, lang, rng):
        c = say(s["c"], lang, rng, t)
        z, e = self.FM[s["f"]]
        return rng.choice([f"{c}设成{z}", f"把{c}改成{z}"] if lang == "zh" else [f"format {c} with {e}", f"show {c} in {e}"])
    def summ(self, s, t, lang): return f"{t.h(s['c'])}格式 {s['f']}" if lang == "zh" else f"format {t.h(s['c'])} {s['f']}"


class GroupAgg(Intent):
    name = "group_agg"
    stage = 4
    FN = {"sum": ("汇总", "合计", "求和"), "avg": ("平均", "均值"), "count": ("计数", "统计条数")}
    FN_EN = {"sum": ("total", "sum up", "summarize"), "avg": ("average",), "count": ("count",)}
    def ok(self, t): return bool(t.by_role("cat") and t.num_cols())
    def slots(self, t, rng):
        return {"by": rng.choice(t.by_role("cat")), "v": rng.choice(t.by_role("amount") or t.num_cols()),
                "fn": rng.choice(["sum"] * 4 + ["avg", "count"])}
    def cf(self, s, t, rng):
        if rng.random() < 0.5 and len(t.by_role("cat")) > 1:
            return {**s, "by": rng.choice([c for c in t.by_role("cat") if c is not s["by"]])}, "分组列"
        return {**s, "fn": rng.choice([f for f in self.FN if f != s["fn"]])}, "聚合函数"
    def step(self, s, t, lang):
        by = t.h(s["by"])
        return {"op": "group_aggregate", "args": {"by": [by], "aggs": [{"col": t.h(s["v"]), "fn": s["fn"]}],
                                                   "output_sheet": (f"按{by}汇总" if lang == "zh" else f"{by} summary")[:31], "mode": "formula"}}
    def text(self, s, t, lang, rng):
        by, v = say(s["by"], lang, rng, t), say(s["v"], lang, rng, t)
        if lang == "zh":
            w = rng.choice(self.FN[s["fn"]])
            if s["fn"] == "count":
                return rng.choice([f"按{by}统计一下有多少条，放新表", f"按{by}计数，结果放到新sheet"])
            return rng.choice([f"按{by}{w}{v}", f"按{by}{w}{v}，结果放新表", f"按{by}把{v}{w}到一张新表", f"分{by}{w}一下{v}"])
        w = rng.choice(self.FN_EN[s["fn"]])
        if s["fn"] == "count":
            return rng.choice([f"count the rows per {by} in a new sheet", f"how many records per {by}, new sheet"])
        return rng.choice([f"{w} {v} by {by}", f"{w} {v} by {by} into a new sheet", f"give me {w} {v} per {by}"])
    def summ(self, s, t, lang): return f"按{t.h(s['by'])}{self.FN[s['fn']][0]}{t.h(s['v'])}到新表" if lang == "zh" else f"{s['fn']} {t.h(s['v'])} by {t.h(s['by'])} -> new sheet"


class Split(Intent):
    name = "split"
    stage = 4
    def ok(self, t): return bool(t.by_role("cat"))
    def slots(self, t, rng): return {"by": rng.choice(t.by_role("cat"))}
    def cf(self, s, t, rng):
        cs = [c for c in t.by_role("cat") if c is not s["by"]]
        return ({"by": rng.choice(cs)}, "拆分列") if cs else None
    def step(self, s, t, lang):
        by = t.h(s["by"])
        return {"op": "split_to_files", "args": {"by": by, "path_template": f"拆分_{by}/{{value}}.xlsx"}}
    def text(self, s, t, lang, rng):
        by = say(s["by"], lang, rng, t)
        return rng.choice([f"按{by}拆成多个文件", f"每个{by}单独存一个文件"] if lang == "zh" else [f"split into one file per {by}", f"save a separate file for each {by}"])
    def summ(self, s, t, lang): return f"按{t.h(s['by'])}拆分文件" if lang == "zh" else f"split files by {t.h(s['by'])}"


class SaveAs(Intent):
    name = "save_as"
    stage = 5
    def slots(self, t, rng):
        csv = rng.random() < 0.3
        return {"p": rng.choice(["结果", "清洗后", "output", "final", "导出/汇总"]) + (".csv" if csv else ".xlsx")}
    def step(self, s, t, lang): return {"op": "save", "args": {"path": s["p"], "format": "csv" if s["p"].endswith(".csv") else "xlsx"}}
    def text(self, s, t, lang, rng):
        return rng.choice([f"另存为{s['p']}", f"保存到{s['p']}", f"导出为{s['p']}"] if lang == "zh" else [f"save as {s['p']}", f"export to {s['p']}"])
    def summ(self, s, t, lang): return f"另存为{s['p']}" if lang == "zh" else f"save as {s['p']}"


class Backup(Intent):
    name = "backup"
    stage = 0
    def step(self, s, t, lang): return {"op": "fs_backup", "args": {"path": t.file}}
    def text(self, s, t, lang, rng): return rng.choice(["先备份一下原表", "改之前先备份", "先把原文件备份"] if lang == "zh" else ["back up the file first", "make a backup before changing anything"])
    def summ(self, s, t, lang): return "备份原表" if lang == "zh" else "backup original"


INTENTS = [StripParens(), Trim(), DeleteEmpty(), DeleteThreshold(), Dedupe(), FillEmpty(), Sort(), AddProduct(),
           CFNegative(), CFThreshold(), NumFormat(), GroupAgg(), Split(), SaveAs(), Backup()]
WEIGHTS = {"strip_parens": 3, "cf_negative": 3, "group_agg": 4, "delete_threshold": 2, "cf_threshold": 2, "sort": 2,
           "delete_empty": 2, "dedupe": 2, "trim": 1, "fill_empty": 1, "add_product": 2, "number_format": 1, "split": 1,
           "save_as": 1, "backup": 1}

JOIN = {"zh": ["，", "，然后", "，再", "；", "，并且", "，接着"], "en": [", ", ", then ", " and then ", "; ", ", and "]}


# ------------------------------------------------------------------ 组装
def build_plan(t, items, lang):
    """items: [(intent, slots)] → 金标 plan（规范顺序 + 依赖 + 默认另存）。"""
    items = sorted(items, key=lambda x: x[0].stage)  # 稳定排序：同阶段保持用户说的顺序
    steps = []
    for i, (it, s) in enumerate(items, 1):
        st = {"id": f"s{i}", **it.step(s, t, lang)}
        if st["op"] not in ("fs_backup", "save"):
            st["sheet"] = t.sheet
        if it.stage == 4:
            dep = [x["id"] for x, (jt, _) in zip(steps, items) if jt.stage in (1, 2)]
            if dep:
                st["depends_on"] = dep
        steps.append(st)
    if not any(s["op"] == "save" for s in steps) and any(s["op"] not in ("fs_backup", "split_to_files") for s in steps):
        stem = t.file.rsplit(".", 1)[0]
        steps.append({"id": f"s{len(steps) + 1}", "op": "save",
                      "args": {"path": f"{stem}_处理后.xlsx" if lang == "zh" else f"{stem}_processed.xlsx"}})
    summ = ("；" if lang == "zh" else "; ").join(it.summ(s, t, lang) for it, s in items)
    return {"version": "0.1", "lang": lang, "summary": summ[:200],
            "inputs": [{"alias": "src", "path": t.file, "sheet": t.sheet}], "steps": steps}


def utter(t, items, lang, rng):
    parts = [it.text(s, t, lang, rng) for it, s in items]
    out = parts[0]
    for p in parts[1:]:
        out += rng.choice(JOIN[lang]) + p
    return out


def profile_text(t):
    """与 protocol.TableProfile.to_prompt 同格式（真实执行时由 profile_sheet 生成）。"""
    return None  # 占位：实际画像在执行校验时从真实 xlsx 生成，保证训练/推理一致


AMBIG = [  # 有歧义的说法 → 主解释 + 备选 + 追问
    ("strip_parens", ["把带括号的{c}删了", "带括号的{c}不要了"], ["delete the {c} with parentheses", "get rid of bracketed {c}"]),
]


def make_ambiguous(t, lang, rng):
    notes = t.by_role("note")
    if not notes:
        return None
    c = rng.choice(notes)
    zh, en = AMBIG[0][1], AMBIG[0][2]
    text = rng.choice(zh if lang == "zh" else en).format(c=say(c, lang, rng, t))
    items = [(StripParens(), {"c": c})]
    extra = []
    if rng.random() < 0.6:  # 和别的明确意图混在一起
        it = rng.choice([x for x in INTENTS if x.name not in ("strip_parens", "backup", "save_as") and x.ok(t)])
        s = it.slots(t, rng)
        extra = [(it, s)]
        text += rng.choice(JOIN[lang]) + it.text(s, t, lang, rng)
    plan = build_plan(t, items + extra, lang)
    sid = next(s["id"] for s in plan["steps"] if s["op"] == "regex_replace")
    h = t.h(c)
    q = (f"“删掉带括号的{h}”是指删除{h}里括号及括号中的文字，还是删除含括号的整行？" if lang == "zh" else
         f"Do you want to remove only the bracketed text inside {h}, or delete the whole rows whose {h} contains brackets?")
    plan["ambiguities"] = [{"step_id": sid, "question": q, "alternatives": [
        {"id": sid, "op": "delete_rows", "sheet": t.sheet, "args": {"where": {"col": h, "cmp": "regex", "value": PAREN_RX}}}]}]
    return text, plan, items + extra


# ------------------------------------------------------------------ 执行校验
def digest(ws_root: Path) -> str:
    h = hashlib.sha1()
    for p in sorted(ws_root.rglob("*")):
        if p.is_file() and "_backup" not in p.parts:
            h.update(str(p.relative_to(ws_root)).encode())
            if p.suffix == ".xlsx":
                wb = load_workbook(p)
                for ws in wb.worksheets:
                    h.update(ws.title.encode())
                    for row in ws.iter_rows():
                        for c in row:
                            v = (c.value.ref, c.value.text) if isinstance(c.value, ArrayFormula) else c.value  # 数组公式对象的 repr 含内存地址
                            h.update(repr((v, c.number_format, c.font.color.rgb if c.font and c.font.color else None)).encode())
                    for cf in ws.conditional_formatting:
                        for r in cf.rules:
                            h.update(repr((str(cf.sqref), r.formula, r.dxf.font.color.rgb if r.dxf and r.dxf.font and r.dxf.font.color else None,
                                           r.dxf.fill.fgColor.rgb if r.dxf and r.dxf.fill else None)).encode())
            else:
                h.update(p.read_bytes())
    return h.hexdigest()


def execute(t, plan, seed):
    """在全新临时工作区里生成表 → 执行 → 返回 (画像文本, 结果摘要)。失败抛异常。"""
    d = Path(tempfile.mkdtemp(prefix="synth_"))
    try:
        t.make_xlsx(d / t.file, random.Random(seed))
        prof = profile_sheet(load_workbook(d / t.file)[t.sheet])
        errs = validate(plan, prof)
        if errs:
            raise ValueError(errs)
        p2 = copy.deepcopy(plan)
        p2.pop("ambiguities", None)
        Executor(Workspace(d)).run(p2)
        return prof.to_prompt(), digest(d)
    finally:
        shutil.rmtree(d, ignore_errors=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--n", type=int, default=3000, help="训练样本数（不含反事实 sibling）")
    ap.add_argument("--n-test", type=int, default=200)
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--out", default=str(ROOT / "data/synth"))
    a = ap.parse_args()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    stats = {"ok": 0, "exec_fail": 0, "cf_pairs": 0, "cf_no_effect": 0, "ambiguous": 0}

    def gen(split, n, themes, rng):
        rows = []
        while len(rows) < n:
            theme = rng.choice(themes)
            t = Table(theme, "zh" if rng.random() < 0.7 else "en", rng)
            lang = "zh" if rng.random() < 0.6 else "en"
            seed = rng.randint(0, 10**9)
            if rng.random() < 0.08:
                r = make_ambiguous(t, lang, rng)
                if not r:
                    continue
                text, plan, _ = r
                try:
                    prof, _ = execute(t, plan, seed)
                except Exception:
                    stats["exec_fail"] += 1
                    continue
                rows.append({"split": split, "theme": theme, "lang": lang, "profile": prof, "file": t.file, "seed": seed, "hlang": t.hlang, "text": text, "plan": plan, "pair": None})
                stats["ambiguous"] += 1
                continue
            pool = [it for it in INTENTS if it.ok(t)]
            k = rng.choice([1, 1, 2, 2, 2, 3, 3])
            chosen = []
            for _ in range(k):
                cand = [it for it in pool if it not in chosen]
                chosen.append(rng.choices(cand, weights=[WEIGHTS[c.name] for c in cand])[0])
            items = [(it, it.slots(t, rng)) for it in chosen]
            plan = build_plan(t, items, lang)
            text = utter(t, items, lang, rng)
            try:
                prof, dg = execute(t, plan, seed)
            except Exception as e:
                stats["exec_fail"] += 1
                if stats["exec_fail"] < 5:
                    print("exec fail:", text, e)
                continue
            pid = None
            sib = None
            cfable = [i for i, (it, s) in enumerate(items) if it.cf(s, t, random.Random(0)) is not None]
            if cfable and rng.random() < 0.5:
                i = rng.choice(cfable)
                it, s = items[i]
                ns, var = it.cf(s, t, rng)
                items2 = list(items)
                items2[i] = (it, ns)
                plan2 = build_plan(t, items2, lang)
                # 同一组随机措辞：用相同种子重放除被干预意图之外的措辞
                text2 = utter(t, items2, lang, random.Random(seed))
                text = utter(t, items, lang, random.Random(seed))
                try:
                    _, dg2 = execute(t, plan2, seed)
                    if dg2 == dg:
                        stats["cf_no_effect"] += 1
                    else:
                        pid = hashlib.sha1(f"{seed}{text}".encode()).hexdigest()[:10]
                        sib = {"split": split, "theme": theme, "lang": lang, "profile": prof, "file": t.file, "seed": seed, "hlang": t.hlang,
                               "text": text2, "plan": plan2, "pair": pid, "intervened": f"{it.name}.{var}"}
                        stats["cf_pairs"] += 1
                except Exception:
                    stats["exec_fail"] += 1
            rows.append({"split": split, "theme": theme, "lang": lang, "profile": prof, "file": t.file, "seed": seed, "hlang": t.hlang, "text": text,
                         "plan": plan, "pair": pid, **({"intervened": sib["intervened"]} if sib else {})})
            if sib:
                rows.append(sib)
            stats["ok"] += 1
        return rows

    rng = random.Random(a.seed)
    train_themes = [k for k in THEMES if k not in TEST_THEMES]
    train = gen("train", a.n, train_themes, rng)
    test_id = gen("test_id", a.n_test // 2, train_themes, random.Random(a.seed + 1))
    test_ood = gen("test_ood", a.n_test // 2, sorted(TEST_THEMES), random.Random(a.seed + 2))
    # 训练集与测试集按指令文本去重，防止泄漏
    seen = {r["text"] for r in train}
    test_id = [r for r in test_id if r["text"] not in seen]
    for name, rows in [("train", train), ("test_id", test_id), ("test_ood", test_ood)]:
        with (out / f"{name}.jsonl").open("w", encoding="utf-8") as f:
            for r in rows:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
    stats.update({k: len(v) for k, v in [("train_rows", train), ("test_id_rows", test_id), ("test_ood_rows", test_ood)]})
    (out / "stats.json").write_text(json.dumps(stats, ensure_ascii=False, indent=1))
    print(json.dumps(stats, ensure_ascii=False))


if __name__ == "__main__":
    main()
