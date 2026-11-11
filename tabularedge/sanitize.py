"""模型输出的宽容清洗（在 schema 校验之前运行）。

思路同 TypeChat / Instructor / Pydantic 的“先强制转换、再校验”：以 schema 为准逐层走一遍模型给出的 JSON，
能确定无歧义修好的就地修好（字段别名、"true"→true、颜色名→十六进制、cmp/fn 别名、步骤编号……），
schema 不认识的非核心字段静默丢弃（执行器本来也不会用它们），每一处改动都记在 fixes 里给界面展示。
修不了的（缺必填字段、引用了不存在的列等）原样保留，交给 validate() 报出具体错误、再让模型自我修正。

例：{"style": {"font_weight": "bold", "color": "red"}}  →  {"style": {"bold": true, "font_color": "FF0000"}}
"""
from __future__ import annotations

import copy
import re

from .protocol import _SCHEMA, OPS

_DROP = object()

# ---------------------------------------------------------------- 别名表
_COLORS = {
    "red": "FF0000", "红": "FF0000", "红色": "FF0000", "大红": "FF0000", "深红": "C00000", "darkred": "C00000",
    "lightred": "FFC7CE", "浅红": "FFC7CE", "淡红": "FFC7CE", "粉": "FFC7CE", "粉色": "FFC7CE", "pink": "FFC7CE",
    "green": "00B050", "绿": "00B050", "绿色": "00B050", "深绿": "006100", "darkgreen": "006100",
    "lightgreen": "C6EFCE", "浅绿": "C6EFCE", "淡绿": "C6EFCE", "浅绿色": "C6EFCE",
    "yellow": "FFFF00", "黄": "FFFF00", "黄色": "FFFF00", "lightyellow": "FFEB9C", "浅黄": "FFEB9C", "淡黄": "FFEB9C", "浅黄色": "FFEB9C",
    "orange": "FFC000", "橙": "FFC000", "橙色": "FFC000", "橘色": "FFC000",
    "blue": "0070C0", "蓝": "0070C0", "蓝色": "0070C0", "lightblue": "BDD7EE", "浅蓝": "BDD7EE", "浅蓝色": "BDD7EE",
    "darkblue": "002060", "深蓝": "002060", "purple": "7030A0", "紫": "7030A0", "紫色": "7030A0",
    "gray": "808080", "grey": "808080", "灰": "808080", "灰色": "808080", "lightgray": "D9D9D9", "lightgrey": "D9D9D9", "浅灰": "D9D9D9",
    "black": "000000", "黑": "000000", "黑色": "000000", "white": "FFFFFF", "白": "FFFFFF", "白色": "FFFFFF",
}
_STYLE_KEYS = {  # 规范化后的键名 → 目标键
    "font_color": "font_color", "color": "font_color", "text_color": "font_color", "fontcolor": "font_color",
    "foreground": "font_color", "fg": "font_color", "fg_color": "font_color", "font_colour": "font_color", "colour": "font_color",
    "字体颜色": "font_color", "文字颜色": "font_color", "颜色": "font_color",
    "fill_color": "fill_color", "fill": "fill_color", "background": "fill_color", "background_color": "fill_color",
    "bg": "fill_color", "bg_color": "fill_color", "bgcolor": "fill_color", "fillcolor": "fill_color", "highlight": "fill_color",
    "cell_color": "fill_color", "填充": "fill_color", "填充色": "fill_color", "背景": "fill_color", "背景色": "fill_color",
    "bold": "bold", "font_weight": "bold", "fontweight": "bold", "weight": "bold", "is_bold": "bold", "font_bold": "bold",
    "strong": "bold", "加粗": "bold", "粗体": "bold",
}
_CMP = {"=": "==", "eq": "==", "equals": "==", "equal": "==", "is": "==", "等于": "==", "<>": "!=", "ne": "!=", "neq": "!=",
        "not_equal": "!=", "不等于": "!=", "lt": "<", "le": "<=", "lte": "<=", "gt": ">", "ge": ">=", "gte": ">=",
        "小于": "<", "大于": ">", "小于等于": "<=", "大于等于": ">=", "不小于": ">=", "不大于": "<=",
        "in": "contains", "include": "contains", "includes": "contains", "like": "contains", "包含": "contains", "含": "contains",
        "not_in": "not_contains", "not_include": "not_contains", "notcontains": "not_contains", "不包含": "not_contains",
        "matches": "regex", "match": "regex", "正则": "regex", "empty": "is_empty", "is_null": "is_empty", "isempty": "is_empty",
        "isnull": "is_empty", "为空": "is_empty", "not_null": "not_empty", "notempty": "not_empty", "is_not_empty": "not_empty",
        "非空": "not_empty", "不为空": "not_empty", "starts_with": "startswith", "begins_with": "startswith", "开头是": "startswith",
        "ends_with": "endswith", "结尾是": "endswith"}
_ENUM_ALIAS = {
    "fn": {"average": "avg", "mean": "avg", "平均": "avg", "求和": "sum", "total": "sum", "合计": "sum", "计数": "count",
           "cnt": "count", "counta": "count", "minimum": "min", "最小": "min", "maximum": "max", "最大": "max"},
    "keep": {"first_one": "first", "earliest": "first", "最早": "first", "第一条": "first", "latest": "last", "newest": "last",
             "最新": "last", "最后一条": "last"},
    "apply_to": {"cells": "cell", "单元格": "cell", "rows": "row", "whole_row": "row", "entire_row": "row", "整行": "row"},
    "mode": {"formulas": "formula", "live": "formula", "公式": "formula", "values": "value", "static": "value", "值": "value"},
    "format": {"excel": "xlsx", "xls": "xlsx"},
    "lang": {"zh-cn": "zh", "cn": "zh", "chinese": "zh", "中文": "zh", "english": "en", "en-us": "en"},
}
_OP_ALIAS = {"filter": "filter_to_sheet", "filter_rows": "filter_to_sheet", "lookup": "lookup_column", "vlookup": "lookup_column",
             "xlookup": "lookup_column", "match": "lookup_column", "join": "lookup_join", "merge": "lookup_join",
             "group_by": "group_aggregate", "groupby": "group_aggregate", "aggregate": "group_aggregate", "pivot": "group_aggregate",
             "summarize": "group_aggregate", "highlight": "conditional_format", "format": "conditional_format",
             "delete": "delete_rows", "remove_rows": "delete_rows", "drop_rows": "delete_rows", "dedup": "dedupe",
             "deduplicate": "dedupe", "remove_duplicates": "dedupe", "drop_duplicates": "dedupe", "sort_rows": "sort",
             "order_by": "sort", "add_col": "add_column", "new_column": "add_column", "formula": "add_column",
             "unique": "unique_to_sheet", "split": "split_to_files", "fillna": "fill_empty", "fill_na": "fill_empty",
             "replace": "regex_replace", "strip": "trim", "save_as": "save", "export": "save"}
_KEY_ALIAS = {  # 通用字段别名（只在目标字段属于该位置的 schema、且原字段不认识时才映射）
    "column": "col", "field": "col", "col_name": "col", "column_name": "col", "columns": "cols", "fields": "cols",
    "condition": "where", "filter": "where", "conditions": "where", "op": "cmp", "operator": "cmp", "comparison": "cmp",
    "compare": "cmp", "val": "value", "values": "value", "target": "value", "and": "all", "or": "any",
    "group_by": "by", "groupby": "by", "aggregations": "aggs", "aggregates": "aggs", "agg": "aggs", "func": "fn",
    "function": "fn", "agg_fn": "fn", "alias": "as", "output": "output_sheet", "sheet_name": "output_sheet", "to_sheet": "output_sheet",
    "new_sheet": "output_sheet", "lookup_sheet": "from_sheet", "source_sheet": "from_sheet", "ref_sheet": "from_sheet",
    "table": "from_sheet", "lookup_key": "from_key", "source_key": "from_key", "ref_key": "from_key", "match_col": "from_key",
    "return_col": "return", "return_column": "return", "result": "return", "default": "if_not_found",
    "not_found": "if_not_found", "missing": "if_not_found", "column_name_new": "name", "new_name": "name", "new_column": "name",
    "expr": "formula", "expression": "formula", "regex": "pattern", "replace": "repl", "replacement": "repl",
    "descending": "desc", "reverse": "desc", "file": "path", "filename": "path", "output_path": "path", "keep_which": "keep",
    "highlight": "style", "format_style": "style", "apply": "apply_to", "scope": "apply_to", "description": "note", "comment": "note",
    "actions": "steps", "plan": "steps", "operations": "steps", "questions": "ambiguities", "options": "alternatives",
}


_COL_KEYS = {"col", "cols", "key", "from_key", "return", "on", "right_on", "fetch", "by"}


def _norm_key(k) -> str:
    k = str(k).strip()
    k = re.sub(r"(?<=[a-z0-9])([A-Z])", r"_\1", k)  # camelCase → snake_case
    return re.sub(r"[\s\-]+", "_", k).lower()


def color(v):
    """颜色写法 → 6 位十六进制：#FF0000 / ff0000 / FFFF0000（ARGB）/ #F00 / rgb(255,0,0) / red / 浅绿。识别不了返回 None。"""
    if isinstance(v, (list, tuple)) and len(v) == 3 and all(isinstance(x, (int, float)) for x in v):
        return "".join(f"{max(0, min(255, int(x))):02X}" for x in v)
    if not isinstance(v, str):
        return None
    s = v.strip().lower().replace(" ", "").replace("_", "")
    if s in _COLORS:
        return _COLORS[s]
    s = s.lstrip("#")
    if s.startswith("0x"):
        s = s[2:]
    if re.fullmatch(r"[0-9a-f]{6}", s):
        return s.upper()
    if re.fullmatch(r"[0-9a-f]{8}", s):
        return s[2:].upper()
    if re.fullmatch(r"[0-9a-f]{3}", s):
        return "".join(c * 2 for c in s).upper()
    m = re.fullmatch(r"rgba?\((\d+),(\d+),(\d+)(?:,[\d.]+)?\)", s)
    if m:
        return "".join(f"{min(255, int(x)):02X}" for x in m.groups())
    return None


def to_bool(v):
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v >= 600 if v > 1 else bool(v)  # font-weight 700 → 加粗
    if isinstance(v, str):
        s = v.strip().lower()
        if s in ("true", "1", "yes", "y", "on", "bold", "bolder", "是", "对", "加粗", "粗体") or (s.isdigit() and int(s) >= 600):
            return True
        if s in ("false", "0", "no", "n", "off", "normal", "lighter", "none", "否", "不", "不加粗", "") or s.isdigit():
            return False
    return None


class Sanitizer:
    def __init__(self, schema=_SCHEMA):
        self.root = schema
        self.fixes: list[str] = []

    # -------------------------------------------------------------- schema 工具
    def _resolve(self, sch):
        while isinstance(sch, dict) and "$ref" in sch:
            node = self.root
            for part in sch["$ref"].lstrip("#/").split("/"):
                node = node[part]
            sch = {**node, **{k: v for k, v in sch.items() if k != "$ref"}}
        return sch

    def _args_schema(self, op):
        for br in self.root["$defs"]["step"].get("allOf", []):
            if br.get("if", {}).get("properties", {}).get("op", {}).get("const") == op:
                return br["then"]["properties"]["args"]
        return None

    def _fix(self, path, msg):
        self.fixes.append(f"{path}: {msg}")

    # -------------------------------------------------------------- 入口
    def plan(self, obj: dict) -> dict:
        if not isinstance(obj, dict):
            return obj
        obj = copy.deepcopy(obj)
        if "steps" not in obj:
            for k in list(obj):
                if _norm_key(k) in ("actions", "plan", "operations") and isinstance(obj[k], list):
                    obj["steps"] = obj.pop(k)
                    self._fix("<root>", f"{k} → steps")
        steps = obj.get("steps")
        if isinstance(steps, dict):  # {"s1": {...}, "s2": {...}}
            obj["steps"] = steps = [{"id": k, **v} if isinstance(v, dict) else v for k, v in steps.items()]
            self._fix("steps", "object → list")
        if isinstance(steps, list):
            keep = []
            for st in steps:  # “表头在第二行”被模型写成一个步骤：表头行已在画像里识别并带进 inputs.header_row，这一步直接去掉
                if isinstance(st, dict) and _norm_key(st.get("op", "")) in ("header", "set_header", "header_row", "skip_rows", "set_header_row"):
                    self._fix("steps", f"removed pseudo-step {st.get('op')!r} (header row is detected automatically)")
                    continue
                keep.append(st)
            obj["steps"] = steps = keep
            self._renumber(steps)
            for amb in obj.get("ambiguities") or []:
                if isinstance(amb, dict) and str(amb.get("step_id")) in getattr(self, "_id_map", {}):
                    amb["step_id"] = self._id_map[str(amb["step_id"])]
            obj["steps"] = [self.step(s, f"steps[{i}]") for i, s in enumerate(steps) if isinstance(s, dict)]
        for a in obj.get("ambiguities") or []:
            if isinstance(a, dict) and isinstance(a.get("alternatives"), list):
                a["alternatives"] = [self.step(s, "ambiguities.alternatives") for s in a["alternatives"] if isinstance(s, dict)]
        return self.value(obj, self.root, "<root>", skip=("steps",))

    def _renumber(self, steps):
        """步骤编号必须是 s1, s2…：缺失 / step1 / 1 / 重复 → 按顺序重编，depends_on 与追问里的 step_id 跟着改。"""
        for s in steps:
            if isinstance(s, dict) and isinstance(s.get("depends_on"), str):
                s["depends_on"] = [s["depends_on"]]
        ids = [s.get("id") if isinstance(s, dict) else None for s in steps]
        ok = all(isinstance(i, str) and re.fullmatch(r"s\d+", i) for i in ids) and len(set(ids)) == len(ids)
        if ok:
            return
        mapping = {}
        for n, s in enumerate(steps, 1):
            if isinstance(s, dict):
                old = s.get("id")
                if old is not None:
                    mapping.setdefault(str(old), f"s{n}")
                s["id"] = f"s{n}"
        for s in steps:
            if isinstance(s, dict) and isinstance(s.get("depends_on"), list):
                s["depends_on"] = [mapping.get(str(d), str(d)) for d in s["depends_on"]]
        self._fix("steps", "step ids renumbered s1..sN")
        self._id_map = mapping

    def step(self, s: dict, path: str) -> dict:
        s = dict(s)
        op = s.get("op")
        if isinstance(op, str) and op not in OPS:
            k = _norm_key(op)
            new = k if k in OPS else _OP_ALIAS.get(k)
            if new:
                self._fix(path, f"op {op!r} → {new!r}")
                s["op"] = op = new
        if not isinstance(s.get("args"), dict):  # 参数被平铺在步骤上：{"id","op","col":…}
            known = {"id", "op", "sheet", "args", "depends_on", "note"}
            flat = {k: s.pop(k) for k in list(s) if k not in known}
            s.pop("args", None)
            s["args"] = flat
            if flat:
                self._fix(path, "args moved under 'args'")
        if isinstance(s.get("depends_on"), str):
            s["depends_on"] = [s["depends_on"]]
        sch = self._args_schema(op)
        if sch is not None:
            s["args"] = self.value(s["args"], sch, f"{path}({op}).args")
        return self.value(s, self.root["$defs"]["step"], path, skip=("args",))

    # -------------------------------------------------------------- 递归转换
    def value(self, v, sch, path, skip=()):
        sch = self._resolve(sch)
        if not isinstance(sch, dict):
            return v
        if "oneOf" in sch and isinstance(v, dict):
            return self._one_of(v, sch, path)
        if "enum" in sch:
            return self._enum(v, sch["enum"], path)
        if "const" in sch:
            return v
        t = sch.get("type")
        types = t if isinstance(t, list) else [t] if t else []
        if "object" in types or ("properties" in sch and not types):
            return self._object(v, sch, path, skip) if isinstance(v, dict) else v
        if "array" in types:
            if not isinstance(v, list):
                if v is None:
                    return v
                self._fix(path, "wrapped into a list")
                v = [v]
            it = sch.get("items")
            return [self.value(x, it, f"{path}[{i}]") for i, x in enumerate(v)] if it else v
        if "boolean" in types and not isinstance(v, bool):
            b = to_bool(v)
            if b is not None:
                self._fix(path, f"{v!r} → {b}")
                return b
            return v
        if ("integer" in types or "number" in types) and isinstance(v, str) and "string" not in types:
            try:
                x = float(v.strip().replace(",", ""))
                x = int(x) if "integer" in types or x.is_integer() else x
                self._fix(path, f"{v!r} → {x}")
                return x
            except ValueError:
                return v
        if "string" in types:
            last = re.sub(r"\[\d+\]$", "", path.rsplit(".", 1)[-1])
            if isinstance(v, str) and last in _COL_KEYS:
                m = re.fullmatch(r"\s*第?\s*([A-Za-z]{1,3})\s*列\s*", v)
                if m:
                    self._fix(path, f"{v!r} → {m.group(1).upper()!r}")
                    return m.group(1).upper()
            if isinstance(v, (int, float)) and not isinstance(v, bool) and not ({"number", "integer"} & set(types)):
                self._fix(path, f"{v!r} → string")
                v = str(v)
            if isinstance(v, str) and "pattern" in sch and not re.search(sch["pattern"], v):
                if sch["pattern"] == "^[0-9A-Fa-f]{6}$":
                    c = color(v)
                    if c:
                        self._fix(path, f"color {v!r} → {c}")
                        return c
        return v

    def _enum(self, v, enum, path):
        if v in enum:
            return v
        key = path.rsplit(".", 1)[-1].split("[")[0]
        s = _norm_key(v) if isinstance(v, str) else v
        cand = None
        if isinstance(v, str):
            low = {str(e).lower(): e for e in enum}
            cand = low.get(v.strip().lower()) or low.get(s)
            if cand is None:
                alias = _CMP if key == "cmp" else _ENUM_ALIAS.get(key, {})
                cand = alias.get(v.strip()) or alias.get(s)
                if cand not in enum:
                    cand = None
        if cand is not None:
            self._fix(path, f"{v!r} → {cand!r}")
            return cand
        return _DROP  # 由上层决定：可选字段就丢弃，必填字段保留原值让校验报错

    def _object(self, v: dict, sch, path, skip=()):
        props = sch.get("properties", {})
        is_style = set(props) == {"font_color", "fill_color", "bold"}
        out = {}
        if is_style:
            v = self._flatten_style(v, path)
        for k, x in v.items():
            if k in skip:
                out[k] = x
                continue
            tgt = k if k in props else None
            if tgt is None:
                nk = _norm_key(k)
                tgt = nk if nk in props else (_STYLE_KEYS.get(nk) or _STYLE_KEYS.get(str(k).strip())) if is_style else _KEY_ALIAS.get(nk)
                if tgt not in props:
                    tgt = None
                if tgt is None and "from_key" in props and nk in ("col", "column", "lookup_col", "match_col", "ref_col", "lookup_column"):
                    tgt = "from_key"  # 跨表查找里的“col”指参考表的匹配列（本表的匹配列是 key）
                if tgt is not None and (tgt in v or tgt in out):
                    tgt = None  # 规范字段已存在：别名的那份当多余字段
            if tgt is None:
                if sch.get("additionalProperties") is False:
                    self._fix(path, f"dropped unknown field {k!r}")
                    continue
                out[k] = x
                continue
            if tgt != k:
                self._fix(path, f"{k!r} → {tgt!r}")
            if is_style and tgt == "bold":
                b = to_bool(x)
                if b is None:
                    self._fix(path, f"dropped {k}={x!r}")
                    continue
                x = b
            y = self.value(x, props[tgt], f"{path}.{tgt}")
            if y is _DROP:
                if tgt in sch.get("required", []):
                    out[tgt] = x
                else:
                    self._fix(path, f"dropped invalid {tgt}={x!r}")
                continue
            if is_style and tgt.endswith("_color") and isinstance(y, str) and not re.fullmatch(r"[0-9A-Fa-f]{6}", y):
                self._fix(path, f"dropped unknown color {tgt}={x!r}")
                continue
            out[tgt] = y
        return out

    def _flatten_style(self, v, path):
        """{"font": {"bold": true, "color": "red"}, "fill": {"color": "yellow"}} → 平铺；字体/填充语义按所在子对象决定。"""
        flat = {}
        for k, x in v.items():
            nk = _norm_key(k)
            if isinstance(x, dict):
                where = "fill_color" if nk in ("fill", "background", "bg", "pattern_fill", "填充") else "font_color"
                for k2, x2 in x.items():
                    n2 = _norm_key(k2)
                    if n2 in ("color", "fg_color", "start_color", "rgb", "颜色"):
                        flat.setdefault(where, x2)
                    else:
                        flat.setdefault(k2, x2)
                self._fix(path, f"flattened nested style {k!r}")
            else:
                flat[k] = x
        return flat

    def _one_of(self, v: dict, sch, path):
        """条件（condition）：叶子 {col,cmp,value} / {all:[…]} / {any:[…]}。按出现的键挑分支。"""
        branches = [self._resolve(b) for b in sch["oneOf"]]
        keys = {_KEY_ALIAS.get(_norm_key(k), _norm_key(k)) for k in v}
        best = max(branches, key=lambda b: (set(b.get("required", [])) <= keys, len(keys & set(b.get("properties", {})))))
        return self._object(v, best, path)


def sanitize(obj: dict) -> tuple[dict, list[str]]:
    s = Sanitizer()
    out = s.plan(obj)
    return out, s.fixes
