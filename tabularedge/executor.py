"""第一层执行器：把已校验的 Action Plan 确定性地落到 xlsx / 文件系统。

设计原则
- 模型不产出可执行代码，只产出 plan；所有副作用都在这里，且可审计（返回逐步日志）。
- 所有路径限制在 workspace 根目录内（沙箱），越界直接拒绝。
- 写入用户原文件前自动备份；默认另存，不覆盖源文件。
- 能写成 Excel 活公式/活条件格式的就写公式（用户在 Excel 里改数据结果会跟着变），
  写不成的（如正则）才落静态值。
"""
from __future__ import annotations

import csv
import datetime as dt
import glob as globmod
import itertools
import re
import shutil
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

from .compat import incompatible, to_2019
from .formula_eval import Evaluator, Unsupported, XLError
from .protocol import RISKY_OPS, WRITE_OPS, check, col_index, profile_sheet


class SandboxError(PermissionError):
    pass


class Workspace:
    def __init__(self, root: str | Path):
        self.root = Path(root).resolve()
        self.root.mkdir(parents=True, exist_ok=True)

    def path(self, p: str) -> Path:
        q = (self.root / p).resolve()
        if q != self.root and self.root not in q.parents:
            raise SandboxError(f"path escapes workspace: {p}")
        return q


# ---------------------------------------------------------------- conditions
def _eval_cond(cond: dict, getv) -> bool:
    if "all" in cond:
        return all(_eval_cond(c, getv) for c in cond["all"])
    if "any" in cond:
        return any(_eval_cond(c, getv) for c in cond["any"])
    v, cmp, x = getv(cond["col"]), cond["cmp"], cond.get("value")
    if cmp == "is_empty":
        return v is None or str(v).strip() == ""
    if cmp == "not_empty":
        return not (v is None or str(v).strip() == "")
    if v is None:
        return False
    if cmp in ("<", "<=", ">", ">="):
        try:
            v, x = float(v), float(x)
        except (TypeError, ValueError):
            return False
        return {"<": v < x, "<=": v <= x, ">": v > x, ">=": v >= x}[cmp]
    s = str(v)
    return {
        "==": v == x or s == str(x), "!=": not (v == x or s == str(x)),
        "contains": str(x) in s, "not_contains": str(x) not in s,
        "startswith": s.startswith(str(x)), "endswith": s.endswith(str(x)),
        "regex": re.search(str(x), s) is not None,
    }[cmp]


def _xl_str(x) -> str:
    return '"' + str(x).replace('"', '""') + '"'


def _cond_to_formula(cond: dict, letter_of, row: int) -> str | None:
    """条件 -> Excel 公式（相对行引用）。不可翻译返回 None。"""
    if "all" in cond or "any" in cond:
        parts = [_cond_to_formula(c, letter_of, row) for c in cond.get("all") or cond.get("any")]
        if any(p is None for p in parts):
            return None
        return ("AND(" if "all" in cond else "OR(") + ",".join(parts) + ")"
    ref = f"${letter_of(cond['col'])}{row}"
    cmp, x = cond["cmp"], cond.get("value")
    num = isinstance(x, (int, float)) and not isinstance(x, bool)
    lit = str(x) if num else _xl_str(x)
    if cmp in ("<", "<=", ">", ">="):
        return f"AND(ISNUMBER({ref}),{ref}{cmp}{lit})"
    return {
        "==": f"{ref}={lit}", "!=": f"{ref}<>{lit}",
        "contains": f"ISNUMBER(SEARCH({_xl_str(x)},{ref}))",
        "not_contains": f"NOT(ISNUMBER(SEARCH({_xl_str(x)},{ref})))",
        "startswith": f"LEFT({ref},{len(str(x))})={_xl_str(x)}",
        "endswith": f"RIGHT({ref},{len(str(x))})={_xl_str(x)}",
        "is_empty": f"LEN(TRIM({ref}))=0", "not_empty": f"LEN(TRIM({ref}))>0",
    }.get(cmp)  # regex -> None（REGEXTEST 仅新版 365 有，保守走静态）


# Excel 2010 之后新增的函数在文件里必须带“未来函数”前缀，否则 Excel 打开显示 #NAME?。
# plan 里始终写干净的函数名，只在落盘时加前缀（参见 ECMA-376 / MS-XLSX 的 _xlfn 约定）。
_XLFN = {n: "_xlfn." + n for n in (
    "XLOOKUP", "XMATCH", "UNIQUE", "SORTBY", "SEQUENCE", "RANDARRAY", "TEXTJOIN", "CONCAT", "IFS", "SWITCH", "MAXIFS",
    "MINIFS", "LET", "LAMBDA", "TEXTBEFORE", "TEXTAFTER", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL", "TOROW", "CHOOSECOLS",
    "CHOOSEROWS", "TAKE", "DROP", "REGEXTEST", "REGEXEXTRACT", "REGEXREPLACE", "IFNA", "DAYS", "ISOWEEKNUM", "STDEV.S",
    "STDEV.P", "VAR.S", "VAR.P", "PERCENTILE.INC", "QUARTILE.INC", "RANK.EQ", "FORECAST.LINEAR", "AGGREGATE",
    "PERCENTRANK.INC", "PERCENTRANK.EXC", "PERCENTILE.EXC", "QUARTILE.EXC", "RANK.AVG", "MODE.SNGL", "CEILING.MATH",
    "FLOOR.MATH", "NETWORKDAYS.INTL", "WORKDAY.INTL", "NUMBERVALUE", "XOR", "SEQUENCE", "ISFORMULA", "FORMULATEXT")}
_XLFN.update({"FILTER": "_xlfn._xlws.FILTER", "SORT": "_xlfn._xlws.SORT"})
_XLFN_RX = re.compile(r"(?<![A-Za-z0-9_.])(" + "|".join(sorted(map(re.escape, _XLFN), key=len, reverse=True)) + r")(?=\()")


def xl(formula: str) -> str:
    """给新函数补 _xlfn 前缀（幂等：已带前缀的不会重复加）。字符串字面量里的同名文字不动。"""
    parts = re.split(r'("(?:[^"]|"")*")', formula)
    return "".join(p if i % 2 else _XLFN_RX.sub(lambda m: _XLFN[m.group(1)], p) for i, p in enumerate(parts))


def _cond_to_array(cond: dict, rng_of) -> str | None:
    """条件 -> 数组表达式（给 FILTER 用），返回 1/0 数组。rng_of(col) -> 绝对区域。不可翻译返回 None。"""
    if "all" in cond or "any" in cond:
        parts = [_cond_to_array(c, rng_of) for c in cond.get("all") or cond.get("any")]
        if any(p is None for p in parts):
            return None
        return "(" + "*".join(parts) + ")" if "all" in cond else "((" + "+".join(parts) + ")>0)"
    R, cmp, x = rng_of(cond["col"]), cond["cmp"], cond.get("value")
    num = isinstance(x, (int, float)) and not isinstance(x, bool)
    lit = str(x) if num else _xl_str(x)
    if cmp in ("<", "<=", ">", ">="):
        return f"(ISNUMBER({R})*({R}{cmp}{lit}))"
    return {"==": f"({R}={lit})", "!=": f"({R}<>{lit})",
            "contains": f"ISNUMBER(SEARCH({_xl_str(x)},{R}))", "not_contains": f"NOT(ISNUMBER(SEARCH({_xl_str(x)},{R})))",
            "is_empty": f"(LEN(TRIM({R}))=0)", "not_empty": f"(LEN(TRIM({R}))>0)"}.get(cmp)


def _is_zh(x) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in str(x))


def _qs(name: str) -> str:
    return "'" + name.replace("'", "''") + "'"


# 公式引用随行变化调整（openpyxl 删行 / 排序不会改公式，必须自己做，否则公式指向错行）
_SHEET_P = r"(?:'(?:[^']|'')+'|[A-Za-z_\u4e00-\u9fff][\w\u4e00-\u9fff.]*)!"
_REF_RX = re.compile(rf"(?<![\w.$])(?P<sh>{_SHEET_P})?(?:(?P<a>\$?[A-Z]{{1,3}}\$?\d+)(?::(?P<b>\$?[A-Z]{{1,3}}\$?\d+))?"
                     r"|(?P<ra>\$?\d+):(?P<rb>\$?\d+))(?![\w(!])")
_CELL_RX = re.compile(r"(\$?[A-Z]{0,3})(\$?)(\d+)")


def _sheet_of(prefix, host):
    if not prefix:
        return host
    n = prefix[:-1]
    return n[1:-1].replace("''", "'") if n.startswith("'") else n


def adjust_formula(f: str, host: str, target: str, row_map, *, relative_only=False) -> str:
    """把公式里指向工作表 target 的行号按 row_map(old)->new|None 改写。字符串字面量不动。
    relative_only=True：只改不带 $ 的行号（排序搬动行时的 Excel 语义）。"""
    def one(cell, end=None):
        m = _CELL_RX.fullmatch(cell)
        col, dollar, row = m.group(1), m.group(2), int(m.group(3))
        if relative_only and dollar:
            return cell
        new = row_map(row, end)
        return None if new is None else f"{col}{dollar}{new}"

    def sub(m):
        if _sheet_of(m.group("sh"), host) != target:
            return m.group(0)
        a, b = m.group("a"), m.group("b")
        if a is None:
            a, b = m.group("ra"), m.group("rb")
        if b is None:
            na = one(a)
            return (m.group("sh") or "") + (na if na else "#REF!")
        na, nb = one(a, "start"), one(b, "end")
        if na is None or nb is None:
            return m.group(0)
        return (m.group("sh") or "") + na + ":" + nb
    parts = re.split(r'("(?:[^"]|"")*")', f)
    return "".join(p if i % 2 else _REF_RX.sub(sub, p) for i, p in enumerate(parts))


# ---------------------------------------------------------------- executor
# [列名] 在“区域参数”位置表示整列（如 MAXIFS([更新时间],[订单号],[订单号]) 的前两个参数），其余位置表示本行单元格
_RANGE_ARGS = {
    "SUMIFS": "0,odd", "AVERAGEIFS": "0,odd", "MAXIFS": "0,odd", "MINIFS": "0,odd", "COUNTIFS": "even",
    "SUMIF": "0,2", "AVERAGEIF": "0,2", "COUNTIF": "0", "PERCENTRANK": "0", "PERCENTRANK.INC": "0", "PERCENTRANK.EXC": "0",
    "RANK": "1", "RANK.EQ": "1", "RANK.AVG": "1", "LARGE": "0", "SMALL": "0", "PERCENTILE": "0", "PERCENTILE.INC": "0",
    "PERCENTILE.EXC": "0", "QUARTILE": "0", "QUARTILE.INC": "0", "VLOOKUP": "1", "HLOOKUP": "1", "MATCH": "1", "XMATCH": "1",
    "XLOOKUP": "1,2", "INDEX": "0", "SUMPRODUCT": "all", "UNIQUE": "0", "FILTER": "0,1", "SORT": "0", "COUNTA": "single",
    "SUM": "single", "AVERAGE": "single", "MAX": "single", "MIN": "single", "COUNT": "single", "MEDIAN": "single",
    "STDEV": "single", "STDEV.S": "single", "STDEV.P": "single", "COUNTBLANK": "single",
}


def _is_range_arg(fn, idx, nargs):
    spec = _RANGE_ARGS.get(fn)
    if not spec:
        return False
    for part in spec.split(","):
        if (part == "all" or (part == "single" and nargs == 1) or (part == "odd" and idx % 2 == 1)
                or (part == "even" and idx % 2 == 0) or (part.isdigit() and int(part) == idx)):
            return True
    return False


def expand_template(formula: str, heads: list) -> str:
    """add_column 公式模板：[列名] → 本行单元格 X{r}，或在区域参数位置 → 整列 $X:$X。
    字符串字面量里的方括号（如正则 "[0-9]"）与不是列名的 [..] 原样保留。"""
    out, i, n = [], 0, len(formula)
    stack = []  # [函数名, 当前参数序号, 参数总数（预扫描）]
    while i < n:
        ch = formula[i]
        if ch == '"':
            j = i + 1
            while j < n:
                if formula[j] == '"':
                    if j + 1 < n and formula[j + 1] == '"':
                        j += 2
                        continue
                    break
                j += 1
            out.append(formula[i:j + 1])
            i = j + 1
            continue
        m = re.match(r"(?:_xlfn\.(?:_xlws\.)?)?([A-Za-z][A-Za-z0-9.]*)\(", formula[i:])
        if m and (i == 0 or not (formula[i - 1].isalnum() or formula[i - 1] in "_.")):
            stack.append([m.group(1).upper(), 0, _count_args(formula, i + m.end())])
            out.append(m.group(0))
            i += m.end()
            continue
        if ch == "(":
            stack.append([None, 0, 1])
        elif ch == ")":
            if stack:
                stack.pop()
        elif ch == "," and stack:
            stack[-1][1] += 1
        elif ch == "[":
            j = formula.find("]", i)
            name = formula[i + 1:j] if j > 0 else None
            if name is not None and name in heads:
                L = get_column_letter(heads.index(name) + 1)
                top = stack[-1] if stack else None
                out.append(f"${L}:${L}" if top and top[0] and _is_range_arg(top[0], top[1], top[2]) else f"{L}{{r}}")
                i = j + 1
                continue
        out.append(ch)
        i += 1
    return "".join(out)


def _count_args(f, start):
    depth, k, i, in_str = 0, 1, start, False
    if f[start:start + 1] == ")":
        return 0
    while i < len(f):
        ch = f[i]
        if in_str:
            if ch == '"':
                in_str = False
        elif ch == '"':
            in_str = True
        elif ch == "(":
            depth += 1
        elif ch == ")":
            if depth == 0:
                return k
            depth -= 1
        elif ch == "," and depth == 0:
            k += 1
        i += 1
    return k


class Executor:
    def __init__(self, workspace: Workspace, confirm=None, compat: str | None = None):
        """compat="2019"：目标是 Excel 2016/2019（无 XLOOKUP / FILTER / UNIQUE 等动态数组函数），公式改写成等价老函数。"""
        self.compat = compat
        self.ws_root = workspace
        self.confirm = confirm or (lambda step: True)  # UI 层注入：写盘前确认
        self.books: dict[str, object] = {}
        self.paths: dict[str, Path] = {}
        self.header_row: dict[str, int] = {}
        self.wb = None
        self.sheet = None
        self.log: list[str] = []
        self._backed_up: set[Path] = set()
        self.ev = None

    def _val(self, ws, r, c):
        """单元格的“值”：公式单元格用内置求值器算出结果（不支持的函数则退回公式文本）。"""
        v = ws.cell(r, c).value
        if isinstance(v, str) and v.startswith("=") and self.ev is not None:
            try:
                v = self.ev.cell(ws, r, c)
            except (Unsupported, RecursionError, ValueError, IndexError, KeyError, TypeError):
                return v
            if isinstance(v, XLError):
                return v.code
            if isinstance(v, float) and v.is_integer():
                return int(v)
        return v

    def _fx(self, f: str) -> str:
        """落盘前的公式：按目标 Excel 版本改写 + 补 _xlfn 前缀。"""
        return xl(to_2019(f) if self.compat == "2019" else f)

    def _delete_rows(self, ws, rows):
        """删除若干行，并像 Excel 一样调整全簿公式里指向这些行之后的引用。"""
        import bisect
        dels = sorted(set(rows))
        if not dels:
            return
        for r in reversed(dels):
            ws.delete_rows(r)
        delset = set(dels)

        def row_map(r, end=None):
            if r in delset:
                if end == "start":  # 区域起点落在被删行：取下一条保留的行
                    while r in delset:
                        r += 1
                elif end == "end":
                    while r in delset and r > 0:
                        r -= 1
                else:
                    return None
            return r - bisect.bisect_left(dels, r)
        for sh in self.wb.worksheets:
            for row in sh.iter_rows():
                for c in row:
                    v = c.value
                    if isinstance(v, str) and v.startswith("="):
                        c.value = adjust_formula(v, sh.title, ws.title, row_map)
                    elif isinstance(v, ArrayFormula):
                        c.value = ArrayFormula(v.ref, adjust_formula(v.text, sh.title, ws.title, row_map))

    # -- helpers
    def _sheet(self, step):
        if step.get("sheet"):
            self.sheet = self.wb[step["sheet"]]
        return self.sheet

    def _hr(self) -> int:
        return self.header_row.get(self.sheet.title, 1)

    def _headers(self, ws=None):
        ws = ws or self.sheet
        hr = self.header_row.get(ws.title, 1)
        return [c.value for c in ws[hr]]

    def _rows(self, ws=None):
        ws = ws or self.sheet
        return range(self.header_row.get(ws.title, 1) + 1, ws.max_row + 1)

    def _ci(self, ref, ws=None):
        return col_index(self._headers(ws), ref)

    def _backup(self, p: Path):
        if p.exists() and p not in self._backed_up:
            bdir = p.parent / "_backup"
            bdir.mkdir(exist_ok=True)
            stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
            dst = bdir / f"{p.stem}.{stamp}{p.suffix}"
            shutil.copy2(p, dst)
            self._backed_up.add(p)
            self.log.append(f"auto-backup {p.name} -> {dst.relative_to(self.ws_root.root)}")

    def _trim_sheet(self, ws):
        """去掉格式残留造成的“空尾巴”：最后一个有值的列 / 行之后的空列空行（常见：标题合并到第 239 列、
        整张表刷过边框）。否则“新增一列”会加在第 240 列，逐行操作也会扫到上千个空行。合并区域同步收缩。"""
        width = last = 0
        for i, r in enumerate(ws.iter_rows(values_only=True), 1):
            hit = False
            for j in range(len(r) - 1, -1, -1):
                if r[j] not in (None, ""):
                    hit = True
                    width = max(width, j + 1)
                    break
            if hit:
                last = i
        mc, mr = ws.max_column, ws.max_row
        if not width or (mc <= width and mr <= last):
            return
        for rng in list(ws.merged_cells.ranges):
            if rng.max_col <= width and rng.max_row <= last:
                continue
            ws.unmerge_cells(str(rng))
            c2, r2 = min(rng.max_col, width), min(rng.max_row, last)
            if rng.min_col <= c2 and rng.min_row <= r2 and (c2 > rng.min_col or r2 > rng.min_row):
                ws.merge_cells(start_row=rng.min_row, start_column=rng.min_col, end_row=r2, end_column=c2)
        if mc > width:
            ws.delete_cols(width + 1, mc - width)
        if mr > last:
            ws.delete_rows(last + 1, mr - last)
        self.log.append(f"trimmed empty tail of {ws.title}: {mc}x{mr} -> {width}x{last}")

    def after_load(self):
        """子类钩子：输入已载入、步骤未执行（客户端在这里注入跨文件参考表）。"""

    # -- entry
    def run(self, plan: dict) -> list[str]:
        prof = None
        for inp in plan.get("inputs", []):
            p = self.ws_root.path(inp["path"])
            wb = load_workbook(p)
            for sh in wb.worksheets:
                self._trim_sheet(sh)
            self.books[inp["alias"]], self.paths[inp["alias"]] = wb, p
            ws = wb[inp["sheet"]] if inp.get("sheet") else wb.active
            self.header_row[ws.title] = inp.get("header_row", 1)
            if self.wb is None:
                self.wb, self.sheet = wb, ws
                self.ev = Evaluator(wb)
                prof = profile_sheet(ws, inp.get("header_row", 1))
        self.after_load()
        check(plan, prof)
        for step in plan["steps"]:
            # 写盘与破坏性操作（整行删除、去重、覆盖原文件）一律经 UI 确认；这是代码里的硬规则，不依赖模型是否追问
            if (step["op"] in WRITE_OPS or step["op"] in RISKY_OPS) and not self.confirm(step):
                self.log.append(f"{step['id']} {step['op']}: skipped (user declined)")
                continue
            if self.ev is not None:
                self.ev.invalidate()
            cb = getattr(self, "progress", None)
            if cb:
                cb(f"step {plan['steps'].index(step) + 1}/{len(plan['steps'])} {step['op']}")
            msg = getattr(self, "op_" + step["op"])(step, step["args"])
            self.log.append(f"{step['id']} {step['op']}: {msg}")
        return self.log

    # -- table ops
    def op_regex_replace(self, step, a):
        ws = self._sheet(step)
        c, rx, n = self._ci(a["col"]), re.compile(a["pattern"]), 0
        for r in self._rows():
            cell = ws.cell(r, c)
            if isinstance(cell.value, str) and not cell.value.startswith("="):  # 公式单元格不动
                new = rx.sub(a.get("repl", ""), cell.value)
                if a.get("strip", True):
                    new = new.strip()
                if new != cell.value:
                    cell.value, n = (new or None), n + 1
        return f"{n} cells changed"

    def op_trim(self, step, a):
        ws, n = self._sheet(step), 0
        for ref in a["cols"]:
            c = self._ci(ref)
            for r in self._rows():
                v = ws.cell(r, c).value
                if isinstance(v, str) and not v.startswith("=") and v != re.sub(r"\s+", " ", v).strip():
                    ws.cell(r, c).value, n = re.sub(r"\s+", " ", v).strip(), n + 1
        return f"{n} cells trimmed"

    def _getter(self, ws, r):
        return lambda ref: self._val(ws, r, self._ci(ref, ws))

    def op_delete_rows(self, step, a):
        ws = self._sheet(step)
        hit = [r for r in self._rows() if _eval_cond(a["where"], self._getter(ws, r))]
        self._delete_rows(ws, hit)
        return f"{len(hit)} rows deleted"

    def op_dedupe(self, step, a):
        ws = self._sheet(step)
        cols = [self._ci(x) for x in a.get("cols") or self._headers()]
        rows = list(self._rows())
        if a.get("keep") == "last":
            rows = rows[::-1]
        seen, dup = set(), []
        for r in rows:
            k = tuple(self._val(ws, r, c) for c in cols)
            (dup.append(r) if k in seen else seen.add(k))
        self._delete_rows(ws, dup)
        return f"{len(dup)} duplicate rows removed"

    def op_fill_empty(self, step, a):
        ws, c, n = self._sheet(step), self._ci(a["col"]), 0
        for r in self._rows():
            if ws.cell(r, c).value in (None, ""):
                ws.cell(r, c).value, n = a["value"], n + 1
        return f"{n} cells filled"

    def op_sort(self, step, a):
        ws = self._sheet(step)
        width = ws.max_column
        # 每行带上“值”（公式列按求值结果排序），再整行搬动；行模板公式最后按新行号重写
        data = [([ws.cell(r, c).value for c in range(1, width + 1)], [self._val(ws, r, c) for c in range(1, width + 1)], r)
                for r in self._rows()]
        for key in reversed(a["by"]):  # 稳定排序，从次键到主键
            i = self._ci(key["col"]) - 1
            # 数字与文本混排时不能直接比较；空值无论升降序都排最后（与 Excel 行为一致）
            filled = [r for r in data if r[1][i] not in (None, "")]
            empty = [r for r in data if r[1][i] in (None, "")]
            filled.sort(key=lambda row: (0, row[1][i], "") if isinstance(row[1][i], (int, float)) and not isinstance(row[1][i], bool)
                        else (0, row[1][i].timestamp(), "") if isinstance(row[1][i], dt.datetime) else (1, 0, str(row[1][i])),
                        reverse=key.get("desc", False))
            data = filled + empty
        for r, (row, _, old) in zip(self._rows(), data):
            for c, v in enumerate(row, 1):
                if isinstance(v, str) and v.startswith("=") and old != r:  # 整行搬动：相对行号随之平移（Excel 排序语义）
                    v = adjust_formula(v, ws.title, ws.title, lambda x, end=None, d=r - old: x + d, relative_only=True)
                ws.cell(r, c).value = v
        return f"{len(data)} rows sorted"

    def _safe_insert_col(self, ws, c) -> bool:
        """在第 c 列前插入空列。openpyxl 插列不会改写公式，所以只有本表公式都不引用第 c 列及以后时才插。"""
        from openpyxl.utils import column_index_from_string as cidx
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                f = v.text if isinstance(v, ArrayFormula) else v
                if isinstance(f, str) and f.startswith("="):
                    body = re.sub(r'"(?:[^"]|"")*"', "", f)
                    body = re.sub(r"(?:'[^']+'|[A-Za-z0-9_\u4e00-\u9fff]+)!\$?[A-Z]{1,3}\$?\d*(?::\$?[A-Z]{1,3}\$?\d*)?", "", body)
                    if any(cidx(m) >= c for m in re.findall(r"\$?\b([A-Z]{1,3})\$?\d*\b(?![(!])", body)):
                        return False
        ws.insert_cols(c)
        return True

    def op_add_column(self, step, a):
        ws = self._sheet(step)
        heads = self._headers()
        c = len(heads) + 1
        if a.get("at"):  # 指定放在某一列（如“将 H 列作为统计类型列”）：原来在该列及之后的列右移
            from openpyxl.utils import column_index_from_string
            want = column_index_from_string(a["at"])
            if want < c:
                if self._safe_insert_col(ws, want):
                    c = want
                    heads = self._headers()
                else:
                    self.log.append(f"{step['id']}: kept {a['name']} at the end (formulas depend on columns after {a['at']})")
            elif want > c:
                c = want
        ws.cell(self._hr(), c).value = a["name"]
        tmpl = expand_template(a["formula"], heads)
        for r in self._rows():
            ws.cell(r, c).value = xl(tmpl.replace("{r}", str(r)))
        if re.match(r"=\s*(?:_xlfn\.)?(DATEVALUE|DATE|EDATE|EOMONTH|TODAY|WORKDAY)\(", tmpl, re.I):
            for r in self._rows():  # 日期型结果默认显示成日期，而不是 46000 这样的序列号
                ws.cell(r, c).number_format = "yyyy-mm-dd"
        if self.compat == "2019":
            fixed = to_2019(tmpl)
            bad = incompatible(fixed)
            if not bad:
                for r in self._rows():
                    ws.cell(r, c).value = xl(fixed.replace("{r}", str(r)))
                return f"column {a['name']} = {fixed}" + (" (rewritten for Excel 2019)" if fixed != tmpl else "")
            # 改写不了：写入计算结果（静态值），并说明原因
            self.ev.invalidate()
            vals = {r: self._val(ws, r, c) for r in self._rows()}
            for r, v in vals.items():
                ws.cell(r, c).value = v
            return f"column {a['name']}: values (Excel 2019 lacks {', '.join(sorted(bad))}; formula was {tmpl})"
        return f"column {a['name']} = {tmpl}"

    def op_conditional_format(self, step, a):
        ws = self._sheet(step)
        heads, first, last = self._headers(), self._hr() + 1, ws.max_row
        if last < first:
            return "no data rows; rule skipped"
        st = a["style"]
        font = Font(color=st.get("font_color"), bold=st.get("bold")) if ("font_color" in st or "bold" in st) else None
        fill = PatternFill("solid", fgColor=st["fill_color"], bgColor=st["fill_color"]) if "fill_color" in st else None
        letter_of = lambda ref: get_column_letter(col_index(heads, ref))
        if a.get("apply_to", "cell") == "row":
            rng = f"A{first}:{get_column_letter(ws.max_column)}{last}"
        else:
            cols = sorted({letter_of(x) for x in _leaf_cols(a["where"])})
            rng = " ".join(f"{L}{first}:{L}{last}" for L in cols)
        f = _cond_to_formula(a["where"], letter_of, first)
        if f is not None:
            ws.conditional_formatting.add(rng, FormulaRule(formula=[f], font=font, fill=fill))
            return f"live rule on {rng}: ={f}"
        n = 0  # 静态兜底
        targets = [col_index(heads, x) for x in _leaf_cols(a["where"])] if a.get("apply_to", "cell") == "cell" \
            else list(range(1, ws.max_column + 1))
        for r in self._rows():
            if _eval_cond(a["where"], self._getter(ws, r)):
                for c in targets:
                    if font: ws.cell(r, c).font = font
                    if fill: ws.cell(r, c).fill = fill
                n += 1
        return f"static style on {n} rows (condition not expressible as Excel formula)"

    def op_number_format(self, step, a):
        ws, c = self._sheet(step), self._ci(a["col"])
        for r in self._rows():
            ws.cell(r, c).number_format = a["format"]
        return f"format {a['format']!r} on column {a['col']}"

    def op_group_aggregate(self, step, a):
        src = self._sheet(step)
        heads = self._headers()
        by = [col_index(heads, x) for x in a["by"]]
        keys: dict[tuple, list[int]] = {}
        for r in self._rows():
            k = tuple(self._val(src, r, c) for c in by)
            if all(v is not None for v in k):
                keys.setdefault(k, []).append(r)
        name = a["output_sheet"]
        if name in self.wb.sheetnames:
            del self.wb[name]
        out = self.wb.create_sheet(name)
        self.header_row[name] = 1
        out.append(list(a["by"]) + [g.get("as") or f"{g['fn']}_{g['col']}" for g in a["aggs"]])
        fx = {"sum": "SUMIFS", "count": "COUNTIFS", "avg": "AVERAGEIFS", "min": "MINIFS", "max": "MAXIFS"}
        S = _qs(src.title)
        for i, (k, rows) in enumerate(sorted(keys.items(), key=lambda kv: tuple(map(str, kv[0]))), start=2):
            line = list(k)
            for g in a["aggs"]:
                vc = col_index(heads, g["col"])
                if a.get("mode", "formula") == "formula":
                    crit = ",".join(f"{S}!${get_column_letter(c)}:${get_column_letter(c)},${get_column_letter(j + 1)}{i}"
                                    for j, c in enumerate(by))
                    L = get_column_letter(vc)
                    line.append(xl(f"=COUNTIFS({crit})" if g["fn"] == "count" else f"={fx[g['fn']]}({S}!${L}:${L},{crit})"))
                else:
                    vals = [self._val(src, r, vc) for r in rows]
                    nums = [float(v) for v in vals if isinstance(v, (int, float))]
                    line.append({"sum": sum(nums), "count": len(vals), "avg": sum(nums) / len(nums) if nums else None,
                                 "min": min(nums, default=None), "max": max(nums, default=None)}[g["fn"]])
            out.append(line)
        return f"{len(keys)} groups -> sheet {name!r} ({a.get('mode', 'formula')})"

    def op_lookup_join(self, step, a):
        ws = self._sheet(step)
        right = self.books[a["right"]].active if a["right"] in self.books else self.wb[a["right"]]
        rh = self._headers(right)
        on_r = col_index(rh, a.get("right_on") or a["on"])
        index = {}
        for r in self._rows(right):
            index.setdefault(right.cell(r, on_r).value, r)
        on_l = self._ci(a["on"])
        base = ws.max_column
        miss = 0
        for j, f in enumerate(a["fetch"], 1):
            ws.cell(self._hr(), base + j).value = f
        fc = [col_index(rh, f) for f in a["fetch"]]
        for r in self._rows():
            hit = index.get(ws.cell(r, on_l).value)
            miss += hit is None
            for j, c in enumerate(fc, 1):
                ws.cell(r, base + j).value = right.cell(hit, c).value if hit else a.get("if_not_found", "")
        return f"joined {len(fc)} cols, {miss} keys not found"

    def _new_sheet(self, name):
        if name in self.wb.sheetnames:
            del self.wb[name]
        out = self.wb.create_sheet(name[:31])
        self.header_row[out.title] = 1
        return out

    def op_lookup_column(self, step, a):
        """跨表查找：本表按 key 去 from_sheet 的 from_key 列匹配，取 return 列。formula 模式写 XLOOKUP 活公式。"""
        ws = self._sheet(step)
        dim = self.wb[a["from_sheet"]]
        dh = self._headers(dim)
        kc, fk, rc = self._ci(a["key"]), col_index(dh, a["from_key"]), col_index(dh, a["return"])
        nf = a.get("if_not_found", "")
        col = len(self._headers()) + 1
        ws.cell(self._hr(), col).value = a["name"]
        index = {}
        for r in self._rows(dim):
            index.setdefault(dim.cell(r, fk).value, dim.cell(r, rc).value)
        D, K, F, Rr = _qs(dim.title), get_column_letter(kc), get_column_letter(fk), get_column_letter(rc)
        miss = 0
        for r in self._rows():
            key = self._val(ws, r, kc)
            miss += key not in index
            if a.get("mode", "formula") == "formula":
                lit = str(nf) if isinstance(nf, (int, float)) and not isinstance(nf, bool) else _xl_str(nf)  # 数字缺省值写数值字面量
                ws.cell(r, col).value = self._fx(f"=XLOOKUP({K}{r},{D}!${F}:${F},{D}!${Rr}:${Rr},{lit})")
            else:
                ws.cell(r, col).value = index.get(key, nf)
        return f"column {a['name']} <- {a['from_sheet']}.{a['return']} ({miss} keys not found, {a.get('mode', 'formula')})"

    def op_filter_to_sheet(self, step, a):
        """筛选到新表。formula 模式写 FILTER 动态数组（按当前结果大小落为数组公式，Excel/LibreOffice 都能显示）。"""
        src = self._sheet(step)
        heads = self._headers()
        cols = [col_index(heads, c) for c in a.get("cols") or heads]
        hit = [r for r in self._rows() if _eval_cond(a["where"], self._getter(src, r))]
        out = self._new_sheet(a["output_sheet"])
        out.append([heads[c - 1] for c in cols])
        first, last = self._hr() + 1, src.max_row
        S = _qs(src.title)
        rng_of = lambda ref: f"{S}!${get_column_letter(col_index(heads, ref))}${first}:${get_column_letter(col_index(heads, ref))}${last}"
        cond = _cond_to_array(a["where"], rng_of) if a.get("mode", "formula") == "formula" and self.compat != "2019" else None
        contiguous = cols == list(range(cols[0], cols[-1] + 1))
        if cond is not None and hit and contiguous:
            data = f"{S}!${get_column_letter(cols[0])}${first}:${get_column_letter(cols[-1])}${last}"
            ref = f"A2:{get_column_letter(len(cols))}{1 + len(hit)}"
            out["A2"] = ArrayFormula(ref, xl(f'=FILTER({data},{cond},"")'))
            return f"{len(hit)} rows -> sheet {out.title!r} (live FILTER over {ref})"
        for r in hit:
            out.append([self._val(src, r, c) for c in cols])
        return f"{len(hit)} rows -> sheet {out.title!r} (values)"

    def op_unique_to_sheet(self, step, a):
        """某列去重后的取值清单到新表（可附计数）。formula 模式：UNIQUE 动态数组 + COUNTIF 活公式。"""
        src = self._sheet(step)
        c = self._ci(a["col"])
        vals = []
        for r in self._rows():
            v = self._val(src, r, c)
            if v not in (None, "") and v not in vals:
                vals.append(v)
        out = self._new_sheet(a["output_sheet"])
        out.append([a["col"]] + (["计数" if a.get("count_header") is None and _is_zh(a["col"]) else a.get("count_header") or "Count"]
                                 if a.get("with_count") else []))
        first, last = self._hr() + 1, src.max_row
        S, L = _qs(src.title), get_column_letter(c)
        R = f"{S}!${L}${first}:${L}${last}"
        live = a.get("mode", "formula") == "formula" and vals
        if live and self.compat == "2019":
            # Excel 2019 没有 UNIQUE：用经典数组公式逐行取“上面还没出现过的第一个值”，源表增删数据后清单自动更新（预留 20 行）
            n = len(vals) + 20
            for i in range(2, 2 + n):
                out.cell(i, 1).value = ArrayFormula(
                    f"A{i}", f'=IFERROR(INDEX({R},MATCH(0,COUNTIF($A$1:A{i - 1},{R})+({R}=""),0)),"")')
                if a.get("with_count"):
                    out.cell(i, 2).value = f'=IF(A{i}="","",COUNTIF({S}!${L}:${L},A{i}))'
            return f"{len(vals)} unique values -> sheet {out.title!r} (live INDEX/MATCH array formulas for Excel 2019)"
        if live:
            out["A2"] = ArrayFormula(f"A2:A{1 + len(vals)}", xl(f'=UNIQUE(FILTER({R},LEN(TRIM({R}))>0))'))
        for i, v in enumerate(vals, start=2):
            if not live:
                out.cell(i, 1).value = v
            if a.get("with_count"):
                out.cell(i, 2).value = f"=COUNTIF({S}!${L}:${L},A{i})" if live else sum(
                    1 for r in self._rows() if self._val(src, r, c) == v)
        return f"{len(vals)} unique values -> sheet {out.title!r} ({'live UNIQUE' if live else 'values'})"

    # -- write ops (sandboxed)
    def op_save(self, step, a):
        p = self.ws_root.path(a["path"])
        p.parent.mkdir(parents=True, exist_ok=True)
        if p in self.paths.values():
            self._backup(p)
        if a.get("format") == "csv" or p.suffix.lower() == ".csv":
            with p.open("w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: Excel 打开中文不乱码
                csv.writer(f).writerows(self.sheet.iter_rows(values_only=True))
        else:
            self.wb.save(p)
        return f"saved {p.relative_to(self.ws_root.root)}"

    def op_fs_mkdir(self, step, a):
        exp = a.get("expand") or {}
        keys = list(exp)
        made = []
        for combo in itertools.product(*(exp[k] for k in keys)) if keys else [()]:
            rel = a["path_template"].format(**dict(zip(keys, combo)))
            p = self.ws_root.path(rel)
            p.mkdir(parents=True, exist_ok=True)
            made.append(rel)
        return f"{len(made)} dirs ensured"

    def op_fs_backup(self, step, a):
        p = self.ws_root.path(a["path"])
        bdir = self.ws_root.path(str(Path(a["path"]).parent / a.get("backup_dir", "_backup")))
        bdir.mkdir(parents=True, exist_ok=True)
        dst = bdir / f"{p.stem}.{dt.datetime.now():%Y%m%d-%H%M%S}{p.suffix}"
        shutil.copy2(p, dst)
        return f"backup -> {dst.relative_to(self.ws_root.root)}"

    def op_fs_new_workbook(self, step, a):
        p = self.ws_root.path(a["path"])
        if p.exists():
            raise FileExistsError(f"{a['path']} exists; refusing to overwrite")
        p.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        wb.active.title = a.get("sheet", "Sheet1")
        wb.active.append(a["headers"])
        for c in wb.active[1]:
            c.font = Font(bold=True)
        wb.active.freeze_panes = "A2"
        wb.save(p)
        return f"created {a['path']}"

    def op_fs_merge_workbooks(self, step, a):
        pattern = str(self.ws_root.path(a["glob"]))
        files = sorted(Path(f) for f in globmod.glob(pattern, recursive=True))
        files = [f for f in files if "_backup" not in f.parts and not f.name.startswith("~$")]  # 跳过备份与 Excel 锁文件
        out_p = self.ws_root.path(a["output"])
        wb = Workbook()
        out = wb.active
        header = None
        n = 0
        for f in files:
            if f.resolve() == out_p:
                continue
            ws = load_workbook(f, read_only=True).active
            rows = ws.iter_rows(values_only=True)
            h = list(next(rows, []))
            if header is None:
                header = h
                out.append((["来源文件"] if a.get("add_source_col", True) else []) + header)
            elif h != header:
                raise ValueError(f"header mismatch in {f.name}: {h} != {header}")
            for row in rows:
                out.append(([f.name] if a.get("add_source_col", True) else []) + list(row))
                n += 1
        out_p.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_p)
        return f"merged {len(files)} files, {n} rows -> {a['output']}"

    def op_split_to_files(self, step, a):
        ws = self._sheet(step)
        c = self._ci(a["by"])
        groups: dict = {}
        for r in self._rows():
            # 拆出的文件里没有其它工作表，公式会断引用 → 写求值后的值
            groups.setdefault(self._val(ws, r, c), []).append([self._val(ws, r, k) for k in range(1, ws.max_column + 1)])
        heads = self._headers()
        for v, rows in groups.items():
            safe = re.sub(r'[\\/:*?"<>|]', "_", str(v))
            p = self.ws_root.path(a["path_template"].format(value=safe))
            p.parent.mkdir(parents=True, exist_ok=True)
            wb = Workbook()
            wb.active.append(heads)
            for row in rows:
                wb.active.append(row)
            wb.save(p)
        return f"{len(groups)} files written"


def _leaf_cols(cond):
    if "all" in cond or "any" in cond:
        for c in cond.get("all") or cond.get("any"):
            yield from _leaf_cols(c)
    else:
        yield cond["col"]
