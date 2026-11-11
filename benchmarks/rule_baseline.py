"""规则基线解析器（中/英 → Action Plan）。

定位（重要）：这 *不是* 最终的第二层。它的用途有三个：
1. 协议与执行器的端到端冒烟测试，不依赖模型；
2. 合成 SFT 数据的“种子标注器”：模板化生成 (指令, plan) 对，再由大模型改写成口语变体；
3. 端侧兜底：模型输出连续校验失败时，对高频简单指令仍能给出可执行 plan。
真正的口语理解由 7B 模型完成（见 llm_parser.py）。
"""
from __future__ import annotations

import datetime as dt
import re

from tabularedge.protocol import TableProfile

SYNONYMS = {
    "金额": ["amount", "amounts", "money", "value", "金額", "钱"],
    "部门": ["department", "departments", "dept", "depts", "部門"],
    "备注": ["note", "notes", "remark", "remarks", "comment", "comments", "memo", "備註"],
    "日期": ["date", "dates"],
    "姓名": ["name", "names"],
    "数量": ["qty", "quantity", "quantities"],
    "单价": ["price", "unit price"],
}

RED, YELLOW, GREEN = "FF0000", "FFFF00", "00B050"
COLORS = {"红": RED, "red": RED, "黄": YELLOW, "yellow": YELLOW, "绿": GREEN, "green": GREEN}
PAREN_RX = r"[\(（][^\)）]*[\)）]"

SPLIT_RX = re.compile(r"[，,；;。]|然后|再|并且|\bthen\b|\band then\b|\band\b(?=\s+(?:make|highlight|sum|summari[sz]e|total|sort|remove|delete|drop|save|split|back))", re.I)


class Ctx:
    def __init__(self, profile: TableProfile | None):
        self.p = profile
        self.headers = profile.headers if profile else []
        self.alias: dict[str, str] = {}
        for h in self.headers:
            self.alias[h.lower()] = h
            for syn in SYNONYMS.get(h, []):
                self.alias.setdefault(syn.lower(), h)

    def cols_in(self, text: str) -> list[str]:
        """按出现位置返回文本中提到的列（表头名或同义词），最长匹配优先。"""
        t, hits = text.lower(), []
        for k in sorted(self.alias, key=len, reverse=True):
            for m in re.finditer(re.escape(k) if not k.isascii() else r"\b" + re.escape(k) + r"\b", t):
                if not any(s <= m.start() < e for s, e, _ in hits):
                    hits.append((m.start(), m.end(), self.alias[k]))
        out = []
        for _, _, h in sorted(hits):
            if h not in out:
                out.append(h)
        return out

    def numeric_cols(self, exclude=()):
        return [h for h in self.headers if self.p and self.p.types.get(h) == "num" and h not in exclude]

    def text_cols(self):
        return [h for h in self.headers if self.p and self.p.types.get(h) == "text"]


def _color(t: str):
    for k, v in COLORS.items():
        if k in t.lower():
            return v
    return None


def _clause(t: str, ctx: Ctx) -> tuple[dict | None, dict | None]:
    """单个子句 → (step 不含 id, ambiguity 或 None)。"""
    tl, cols = t.lower(), ctx.cols_in(t)
    col = cols[0] if cols else None

    # 括号内容删除
    if re.search(r"括号|parenthes|bracket", tl):
        target = col or next((h for h in ctx.headers if h == "备注"), None)
        if target:
            step = {"op": "regex_replace", "args": {"col": target, "pattern": PAREN_RX, "repl": ""}}
            alt = {"op": "delete_rows", "args": {"where": {"col": target, "cmp": "regex", "value": PAREN_RX}}}
            amb = {"question": f"“{t.strip()}”是指删除【{target}】里括号及括号中的文字，还是删除含括号的整行？",
                   "alternatives": [alt]}
            return step, amb

    # 负数/阈值 标色
    color = _color(t)
    if color and re.search(r"标|highlight|mark|make|color|colour|显示|填充|fill", tl):
        target = col or (ctx.numeric_cols()[:1] or [None])[0]
        if target:
            cond = None
            if re.search(r"负|negative|小于\s*0|<\s*0|below zero", tl):
                cond = {"col": target, "cmp": "<", "value": 0}
            elif m := re.search(r"(大于|超过|高于|>|over|above|greater than)\s*(-?\d+(?:\.\d+)?)", tl):
                cond = {"col": target, "cmp": ">", "value": float(m.group(2))}
            elif m := re.search(r"(小于|低于|<|under|below|less than)\s*(-?\d+(?:\.\d+)?)", tl):
                cond = {"col": target, "cmp": "<", "value": float(m.group(2))}
            elif re.search(r"空|empty|blank", tl):
                cond = {"col": target, "cmp": "is_empty"}
            if cond:
                fill = re.search(r"填充|底色|背景|fill|background", tl) or color == YELLOW
                style = {"fill_color": color} if fill else {"font_color": color}
                row = bool(re.search(r"整行|entire row|whole row|row", tl))
                return {"op": "conditional_format",
                        "args": {"where": cond, "apply_to": "row" if row else "cell", "style": style}}, None

    # 分组汇总
    if re.search(r"汇总|统计|小计|合计|summar|total|sum up|subtotal|group", tl) and \
            (m := re.search(r"(?:按照|按|根据)\s*(.+?)(?:分类|分组)?(?:汇总|统计|小计|合计|$)|\b(?:by|per|for each)\s+(\w+(?:\s+(?:and|&)\s+\w+)?)", t, re.I)):
        by = ctx.cols_in(m.group(1) or m.group(2))[:2]
        if by:
            vals = [c for c in cols if c not in by] or ctx.numeric_cols(exclude=by)
            if not vals:
                return None, {"question": f"按{by[0]}汇总时要汇总哪一列？"}
            fn = "count" if re.search(r"个数|人数|笔数|count|how many", tl) else "avg" if re.search(r"平均|average|mean", tl) else "sum"
            return {"op": "group_aggregate",
                    "args": {"by": by, "aggs": [{"col": v, "fn": fn} for v in vals],
                             "output_sheet": f"按{by[0]}汇总"[:31], "mode": "formula"}}, None

    # 排序
    if re.search(r"排序|排列|从大到小|从小到大|\bsort|\border\b|rank", tl) and col:
        desc = bool(re.search(r"降序|从大到小|倒序|desc|largest|highest|biggest", tl))
        return {"op": "sort", "args": {"by": [{"col": col, "desc": desc}]}}, None

    # 去重
    if re.search(r"去重|重复|duplicate|dedup", tl):
        return {"op": "dedupe", "args": ({"cols": cols} if cols else {})}, None

    # 去空格
    if re.search(r"空格|trim|whitespace|spaces", tl):
        return {"op": "trim", "args": {"cols": cols or ctx.text_cols()}}, None

    # 空值：填充 or 删行
    if re.search(r"空|empty|blank|missing", tl) and col:
        if m := re.search(r"(?:填|fill(?: with)?|补)\s*(-?\d+|“[^”]+”|\"[^\"]+\"|\S+)", t, re.I):
            v = m.group(1).strip('“”"')
            v = float(v) if re.fullmatch(r"-?\d+(\.\d+)?", v) else v
            return {"op": "fill_empty", "args": {"col": col, "value": v}}, None
        if re.search(r"删|去掉|remove|delete|drop", tl):
            return {"op": "delete_rows", "args": {"where": {"col": col, "cmp": "is_empty"}}}, None

    # 拆分成多个文件
    if re.search(r"拆分|拆成|split", tl) and col:
        return {"op": "split_to_files", "args": {"by": col, "path_template": f"拆分_{col}/{{value}}.xlsx"}}, None

    # 按月份建文件夹
    if re.search(r"文件夹|目录|folder|director", tl):
        year = dt.date.today().year
        if m := re.search(r"(20\d\d)", t):
            year = int(m.group(1))
        expand, parts = {}, [str(year)]
        if re.search(r"月|month", tl):
            expand["mm"] = [f"{i:02d}月" for i in range(1, 13)]
            parts.append("{mm}")
        if col or re.search(r"部门|department", tl):
            expand["dept"] = ["<从表中取值>"]  # 由执行前的 UI/执行器从数据中填充
            parts.append("{dept}")
        if expand:
            return {"op": "fs_mkdir", "args": {"path_template": "/".join(parts), "expand": expand}}, \
                ({"question": "部门文件夹需要读取表中部门列的去重值，确认使用哪一列？"} if "dept" in expand else None)

    # 备份
    if re.search(r"备份|back ?up", tl):
        return {"op": "fs_backup", "args": {"path": "<input>"}}, None

    # 另存
    if m := re.search(r"(?:另存为|保存为|保存到|导出为|导出到|save (?:it )?(?:as|to)|export (?:as|to))\s*([^\s，,；;]+)", t, re.I):
        path = m.group(1)
        return {"op": "save", "args": {"path": path, "format": "csv" if path.lower().endswith(".csv") else "xlsx"}}, None

    return None, {"question": f"没能理解“{t.strip()}”，能换个说法吗？"}


def parse(text: str, profile: TableProfile | None = None, input_path: str = "input.xlsx") -> dict:
    ctx = Ctx(profile)
    lang = "zh" if re.search(r"[一-鿿]", text) else "en"
    steps, ambs = [], []
    for raw in filter(None, (c.strip() for c in SPLIT_RX.split(text) if c)):
        step, amb = _clause(raw, ctx)
        if step:
            step["id"] = f"s{len(steps) + 1}"
            if step["op"] == "fs_backup":
                step["args"]["path"] = input_path
            if profile and step["op"] not in ("fs_mkdir", "fs_backup", "save"):
                step["sheet"] = profile.sheet
            steps.append(step)
        if amb:
            amb["step_id"] = step["id"] if step else "s0"
            for alt in amb.get("alternatives", []):
                alt["id"] = step["id"]
                if profile:
                    alt["sheet"] = profile.sheet
            ambs.append(amb)
    # 汇总应基于清洗后的数据：把 group_aggregate 挪到所有清洗步骤之后（格式步骤不影响值，可在其后）
    cleaning = {"regex_replace", "trim", "delete_rows", "dedupe", "fill_empty"}
    if any(s["op"] == "group_aggregate" for s in steps):
        clean = [s for s in steps if s["op"] in cleaning]
        for s in steps:
            if s["op"] == "group_aggregate":
                s["depends_on"] = [c["id"] for c in clean]
        order = [s for s in steps if s["op"] in cleaning] + [s for s in steps if s["op"] not in cleaning]
        remap = {s["id"]: f"s{i}" for i, s in enumerate(order, 1)}
        for s in order:
            s["id"] = remap[s["id"]]
            if "depends_on" in s:
                s["depends_on"] = sorted(remap[d] for d in s["depends_on"])
        for a in ambs:
            a["step_id"] = remap.get(a["step_id"], a["step_id"])
            for alt in a.get("alternatives", []):
                alt["id"] = remap.get(alt["id"], alt["id"])
        steps = order
    table_ops = [s for s in steps if s["op"] not in ("fs_mkdir", "fs_backup", "save")]
    if table_ops and not any(s["op"] == "save" for s in steps):
        stem = input_path.rsplit(".", 1)[0]
        steps.append({"id": f"s{len(steps) + 1}", "op": "save",
                      "args": {"path": f"{stem}_处理后.xlsx" if lang == "zh" else f"{stem}_processed.xlsx"}})
    plan = {"version": "0.1", "lang": lang, "summary": text[:200],
            "inputs": [{"alias": "src", "path": input_path, **({"sheet": profile.sheet} if profile else {})}],
            "steps": steps}  # 为空时校验会失败，由调用方走 ambiguities 追问
    if ambs:
        plan["ambiguities"] = ambs
    return plan
