"""Action Plan 协议：结构校验 + 语义校验 + 表格画像(模型唯一能看到的表格信息)。"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import jsonschema
from openpyxl.utils import column_index_from_string

SCHEMA_PATH = Path(__file__).resolve().parent.parent / "schema" / "action_plan.schema.json"
_SCHEMA = json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))
_VALIDATOR = jsonschema.Draft202012Validator(_SCHEMA)

OPS = set(_SCHEMA["$defs"]["step"]["allOf"][-1]["properties"]["op"]["enum"])
# 会写磁盘的操作：执行前必须经用户确认
WRITE_OPS = {"save", "fs_mkdir", "fs_backup", "fs_new_workbook", "fs_merge_workbooks", "split_to_files"}
# 破坏性操作：无论模型是否追问，执行前一律经 UI 确认（v0.2）
RISKY_OPS = {"delete_rows", "dedupe"}


@dataclass
class TableProfile:
    """喂给模型的表格画像：只含表头/类型/少量样例，绝不含整表数据（隐私 + 上下文长度）。"""
    sheet: str
    headers: list[str]
    types: dict[str, str]
    n_rows: int
    samples: list[list] = field(default_factory=list)
    header_row: int = 1  # 表头所在行（第 1 行是标题、第 2 行才是表头的表很常见）；不进提示词，只带进 plan.inputs

    def to_prompt(self) -> str:
        cols = ", ".join(f"{h}:{self.types.get(h, '?')}" for h in self.headers)
        lines = [f"[sheet={self.sheet} rows={self.n_rows}] columns: {cols}"]
        for r in self.samples:
            lines.append("  | " + " | ".join("" if v is None else str(v)[:24] for v in r))
        return "\n".join(lines)


def detect_header_row(rows: list, scan: int = 6) -> int:
    """猜表头行：跳过“只有一个格子有字”的大标题行，取第一行“非空格子数接近最宽一行、且多为文字”的行。"""
    counts = [sum(1 for v in r if v not in (None, "")) for r in rows[:12]]
    widest = max(counts, default=0)
    if widest <= 1:
        return 1
    for i, r in enumerate(rows[:scan]):
        filled = [v for v in r if v not in (None, "")]
        texts = sum(1 for v in filled if isinstance(v, str))
        if len(filled) >= max(2, 0.6 * widest) and texts >= 0.6 * len(filled):
            return i + 1
    return 1


def used_width(rows) -> int:
    """最后一个“有值”的列号（行数据里任何一行有值都算）。格式残留造成的几百个空列不算。"""
    w = 0
    for r in rows:
        for j in range(len(r) - 1, w - 1, -1):
            if r[j] not in (None, ""):
                w = j + 1
                break
    return w


def profile_sheet(ws, header_row: int = 1, n_samples: int = 3) -> TableProfile:
    headers = [str(c.value) if c.value is not None else f"col{c.column}" for c in ws[header_row]]
    types: dict[str, str] = {}
    samples = []
    for i, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True)):
        if i < n_samples:
            samples.append(list(row))
        for h, v in zip(headers, row):
            if v is None:
                continue
            t = "num" if isinstance(v, (int, float)) else "date" if hasattr(v, "year") else "text"
            prev = types.get(h)
            types[h] = t if prev in (None, t) else "mixed"
    return TableProfile(ws.title, headers, types, max(ws.max_row - header_row, 0), samples)


class PlanError(ValueError):
    def __init__(self, errors: list[str]):
        super().__init__("; ".join(errors))
        self.errors = errors


def _cols_in(obj):
    """递归收集 plan 中所有列引用（用于语义校验）。"""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k in ("col", "on") and isinstance(v, str):
                yield v
            elif k in ("cols", "by") and isinstance(v, list):
                for x in v:
                    if isinstance(x, str):
                        yield x
                    elif isinstance(x, dict) and "col" in x:
                        yield x["col"]
            elif k == "by" and isinstance(v, str):
                yield v
            else:
                yield from _cols_in(v)
    elif isinstance(obj, list):
        for x in obj:
            yield from _cols_in(x)


def validate(plan: dict, profile: TableProfile | None = None) -> list[str]:
    """返回错误列表（空=通过）。错误文本会回喂给模型做自修复，所以要具体。"""
    errs = [f"schema: {'/'.join(map(str, e.absolute_path)) or '<root>'}: {e.message}"
            for e in sorted(_VALIDATOR.iter_errors(plan), key=lambda e: list(e.absolute_path))]
    if errs:
        return errs
    ids = [s["id"] for s in plan["steps"]]
    if len(ids) != len(set(ids)):
        errs.append("semantic: duplicate step ids")
    seen: set[str] = set()
    created_cols: set[str] = set()
    for s in plan["steps"]:
        for d in s.get("depends_on", []):
            if d not in seen:
                errs.append(f"semantic: {s['id']} depends_on {d} which is not an earlier step")
        seen.add(s["id"])
        a = s["args"]
        if s["op"] == "regex_replace" or (isinstance(a.get("where"), dict) and a["where"].get("cmp") == "regex"):
            pat = a.get("pattern") or a["where"].get("value", "")
            try:
                re.compile(pat)
            except re.error as e:
                errs.append(f"semantic: {s['id']} bad regex {pat!r}: {e}")
        if profile is not None and s["op"] not in ("fs_mkdir", "fs_merge_workbooks", "fs_new_workbook", "fs_backup", "save"):
            known = set(profile.headers) | created_cols
            refs = {"lookup_join": {"col": a.get("on")}, "lookup_column": {"col": a.get("key")}}.get(s["op"], a)
            for c in _cols_in(refs):
                if c not in known and not re.fullmatch(r"[A-Z]{1,3}", c):
                    errs.append(f"semantic: {s['id']} unknown column {c!r}; available: {sorted(known)}")
        if s["op"] == "add_column" and profile is not None and isinstance(a.get("formula"), str):
            known = set(profile.headers) | created_cols
            bare = re.sub(r'"(?:[^"]|"")*"', '""', a["formula"])  # 字符串字面量里的方括号（正则字符类）不算列引用
            for ref in re.findall(r"\[([^\]]+)\]", bare):
                if ref not in known:
                    errs.append(f"semantic: {s['id']} formula references unknown column [{ref}]; available: {sorted(known)}")
        if s["op"] in ("add_column", "lookup_column"):
            created_cols.add(a["name"])
        if s["op"] == "lookup_join":
            created_cols.update(a["fetch"])
    return errs


def check(plan: dict, profile: TableProfile | None = None) -> dict:
    errs = validate(plan, profile)
    if errs:
        raise PlanError(errs)
    return plan


def col_index(headers: list[str], ref: str) -> int:
    """列引用 -> 1-based 列号。表头名优先，其次列字母。"""
    if ref in headers:
        return headers.index(ref) + 1
    m = re.fullmatch(r"\s*第?\s*([A-Za-z]{1,3})\s*列?\s*", str(ref))  # "F" / "F列" / "第F列"
    if m and (m.group(1).isupper() or str(ref).strip() != m.group(1)):
        return column_index_from_string(m.group(1).upper())
    raise KeyError(f"unknown column {ref!r}")
