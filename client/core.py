"""客户端核心逻辑（无界面依赖，桌面程序与自动化评测共用）。"""
from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook

from tabularedge.executor import Executor
from tabularedge.formula_eval import Evaluator, XLError
from tabularedge.protocol import TableProfile, col_index, detect_header_row, used_width
from tabularedge.xlsx_stream import XlsxStream, norm_key

# 超过这个大小的文件不再用 openpyxl 打开做画像，改用流式读取（只读表头 + 前几十行）。
# 小文件仍走原路径：openpyxl 存的文件里公式没有缓存值，需要内置求值器算出样例值（链式第 2 步要用）。
SMALL_BYTES = 2 * 2**20
# 参考表超过这个规模（文件大小或单元格数）就不复制进输出文件，查找改为流式建索引、写入结果值。
BIG_REF_BYTES = 20 * 2**20
BIG_REF_CELLS = 1_500_000
# 主表超过这个单元格数，逐格编辑（openpyxl 全量载入）需要十几 GB 内存，执行前拦下并给出建议。
BIG_MAIN_CELLS = 3_000_000
TYPE_ROWS = 200  # 流式画像时用前多少行推断列类型


def is_big(path) -> bool:
    try:
        return Path(path).stat().st_size > SMALL_BYTES
    except OSError:
        return False


def sheet_names(path) -> list[str]:
    """只读 workbook.xml，毫秒级（150MB 的文件也一样）。"""
    with XlsxStream(path) as xs:
        return xs.sheet_names


def stream_profile(xs: XlsxStream, sheet: str, header_row: int | None = None, n_samples: int = 3, text: str = "") -> TableProfile:
    """流式画像：只解压工作表开头一小段（表头 + 前 TYPE_ROWS 行），行数取自 <dimension>。表头行自动识别。"""
    rows = xs.head(sheet, 12 + TYPE_ROWS)
    header_row = header_row or detect_header_row(rows)
    max_row, _ = xs.dims(sheet)
    head = rows[header_row - 1] if len(rows) >= header_row else []
    width = used_width(rows[header_row - 1:])
    headers = [str(head[i]) if i < len(head) and head[i] is not None else f"col{i + 1}" for i in range(width)]
    body = [(r + [None] * (width - len(r)))[:width] for r in rows[header_row:header_row + TYPE_ROWS]]
    types: dict[str, str] = {}
    for row in body:
        for h, v in zip(headers, row):
            if v is None:
                continue
            t = "num" if isinstance(v, (int, float)) and not isinstance(v, bool) else "date" if hasattr(v, "year") else "text"
            prev = types.get(h)
            types[h] = t if prev in (None, t) else "mixed"
    if max_row and max_row > len(rows):
        n_rows = max_row - header_row
    else:  # 小表：按实际有值的行算（<dimension> 可能包含格式残留的空行）
        n_rows = max((i + 1 for i, r in enumerate(rows) if any(v not in (None, "") for v in r)), default=0) - header_row
    return TableProfile(sheet, headers, types, max(n_rows, 0), body[:n_samples], header_row=header_row)


def quick_profile(path, sheet: str | None):
    """界面左侧“表结构”面板用：(画像, 全部表名)。大文件流式，小文件照旧（带公式求值）。"""
    if is_big(path):
        with XlsxStream(path) as xs:
            names = xs.sheet_names
            return stream_profile(xs, sheet or names[0]), names
    wb = load_workbook(path)
    ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
    return profile_eval(ws, Evaluator(wb)), wb.sheetnames


def main_size(path, sheet: str | None):
    """(行数, 列数)：读 <dimension>，毫秒级；读不到返回 (None, None)。"""
    try:
        with XlsxStream(path) as xs:
            return xs.dims(sheet if sheet in xs.sheet_names else None)
    except Exception:  # noqa: BLE001
        return None, None


def is_big_ref(path, sheet) -> bool:
    try:
        if Path(path).stat().st_size > BIG_REF_BYTES:
            return True
        r, c = main_size(path, sheet)
        return bool(r and c and r * c > BIG_REF_CELLS)
    except OSError:
        return False


def _grams(s: str) -> set:
    s = re.sub(r"\s+", "", s)
    return {s[i:i + 2] for i in range(len(s) - 1) if not s[i:i + 2].isascii() or s[i:i + 2].isalnum()}


def sheet_shape(ws, header_row: int | None = None, text: str = ""):
    """(表头行, 有效列数, 最后一个有值的行)：表头行 = 指令明说的 > 自动识别；格式残留的空列 / 空行不算。"""
    vals = [list(r) for r in ws.iter_rows(values_only=True)]
    hr = header_row or detect_header_row(vals)
    last = max((i + 1 for i, r in enumerate(vals) if any(v not in (None, "") for v in r)), default=0)
    return hr, used_width(vals), last


def profile_eval(ws, ev: Evaluator | None, header_row: int | None = None, n_samples: int = 3, text: str = "",
                 pick_samples: bool = False) -> TableProfile:
    """表格画像（同 protocol.profile_sheet，格式不变），但
    - 表头行自动识别（第 1 行是大标题的表）；只算有值的列和行（格式残留的几百个空列不进提示词）；
    - 公式单元格用求值结果，避免把公式文本当样例喂给模型；
    - pick_samples=True 时样例行“看指令挑”（实测会让模型照抄样例里的姓名，默认关闭）。"""
    header_row, width, last = sheet_shape(ws, header_row, text)
    head = [c.value for c in ws[header_row]][:width] if width else []
    head += [None] * (width - len(head))
    headers = [str(v) if v is not None else f"col{i + 1}" for i, v in enumerate(head)]
    tg = _grams(text) - {g for h in headers for g in _grams(h)} if text and pick_samples else set()  # 列名本身的字不算
    scored = []
    types, samples = {}, []
    for i, r in enumerate(range(header_row + 1, last + 1)):
        row = []
        for c in range(1, width + 1):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("=") and ev is not None:
                try:
                    v = ev.cell(ws, r, c)
                    v = v.code if isinstance(v, XLError) else (int(v) if isinstance(v, float) and v.is_integer() else v)
                except Exception:  # noqa: BLE001
                    pass
            row.append(v)
        if i < n_samples:
            samples.append(row)
        if tg and i < 5000:
            sc = sum(1 for v in row if isinstance(v, str) and len(v) <= 40 and _grams(v) & tg)
            if sc:
                scored.append((-sc, i, row))
        for h, v in zip(headers, row):
            if v is None:
                continue
            t = "num" if isinstance(v, (int, float)) else "date" if hasattr(v, "year") else "text"
            prev = types.get(h)
            types[h] = t if prev in (None, t) else "mixed"
    if scored and samples:
        picked = [row for _, i, row in sorted(scored, key=lambda x: (x[0], x[1])) if i != 0][:n_samples - 1]
        samples = (samples[:1] + picked + samples[1:])[:n_samples]
    return TableProfile(ws.title, headers, types, max(last - header_row, 0), samples, header_row=header_row)


class RefExecutor(Executor):
    """在载入主工作簿后，把“参考文件”的工作表复制进来（跨文件查找；活公式需要这些表在同一个输出文件里）。

    超大参考表（如 76 万行的基本单位导出表）不复制：整表复制进输出文件既占十几 GB 内存，也会让结果文件膨胀到上百 MB。
    对它的查找（lookup_column / lookup_join）改为流式建索引——只保留主表里出现过的键——并把匹配结果写成值。"""

    def __init__(self, ws, confirm, refs, compat=None, progress=None):
        super().__init__(ws, confirm, compat=compat)
        self.refs = refs  # [(sheet_title, src_path, src_sheet)]
        self.big_refs: dict[str, tuple[str, str]] = {}
        self.progress = progress or (lambda msg: None)

    def after_load(self):
        for title, path, sheet in self.refs:
            if title in self.wb.sheetnames:
                continue
            if is_big_ref(path, sheet):
                self.big_refs[title] = (path, sheet)
                r, c = main_size(path, sheet)
                self.log.append(f"ref-sheet {title} <- {Path(path).name} (large: {r} rows x {c} cols, streamed lookup only)")
                continue
            self.progress(f"copy-ref {title}")
            src = load_workbook(path, read_only=True, data_only=False)[sheet]  # 流式读源表，不建完整对象模型
            src.reset_dimensions()  # 有些导出工具写的 <dimension> 不准，不信它，读到底
            dst = self.wb.create_sheet(title)
            for row in src.iter_rows():
                for c in row:
                    if c.value is None or not hasattr(c, "column"):
                        continue
                    d = dst.cell(c.row, c.column, c.value)
                    if c.number_format != "General":
                        d.number_format = c.number_format
            src.parent.close()
            self.header_row[dst.title] = detect_header_row([list(r) for r in dst.iter_rows(max_row=12, values_only=True)])
            self.log.append(f"ref-sheet {title} <- {Path(path).name}"
                            + (f" (header row {self.header_row[dst.title]})" if self.header_row[dst.title] != 1 else ""))
        if self.ev is not None:
            self.ev.invalidate()

    def run(self, plan):
        self._guard_big_main(plan)
        self._guard_big_refs(plan)
        self._lookup_cache = {}
        self._planned_rets: dict[tuple, list] = {}
        for st in plan.get("steps", []):
            a = st.get("args", {})
            if st.get("op") == "lookup_column" and a.get("from_key"):
                self._planned_rets.setdefault((a.get("from_sheet"), a["from_key"]), []).append(a.get("return"))
            elif st.get("op") == "lookup_join" and a.get("on"):
                self._planned_rets.setdefault((a.get("right"), a.get("right_on") or a["on"]), []).extend(a.get("fetch", []))
        return super().run(plan)

    def _guard_big_main(self, plan):
        """主表（要逐格修改的那张）太大：openpyxl 全量载入 76 万行 × 25 列需要十几 GB 内存、十几分钟，直接说明并给出做法。"""
        for inp in plan.get("inputs", [])[:1]:
            p = self.ws_root.path(inp["path"])
            r, c = main_size(p, inp.get("sheet"))
            if r and c and r * c > BIG_MAIN_CELLS:
                raise ValueError(
                    f"《{p.name}》有 {r - 1:,} 行 × {c} 列（约 {r * c / 1e4:,.0f} 万个单元格），超出逐格编辑的能力"
                    f"（需要十几 GB 内存）。可以这样用：选中要补充信息的小表，在指令里写“从《{p.stem}》按××匹配××”，"
                    f"大表会作为参考表流式查找；或先在 Excel 里筛出需要的部分另存后再处理。\n"
                    f"{p.name} has {r - 1:,} rows x {c} cols - too large to edit cell by cell; use it as a lookup reference instead.")

    def _guard_big_refs(self, plan):
        """大参考表只支持“查找/匹配”两种操作；公式里直接引用它（如 VLOOKUP(…,大表!A:B,…)）给出明确提示。"""
        bigs = {t for t, p, sh in self.refs if is_big_ref(p, sh)}

        def mentions(x, t):
            if isinstance(x, dict):
                return any(mentions(v, t) for v in x.values())
            if isinstance(x, list):
                return any(mentions(v, t) for v in x)
            return isinstance(x, str) and (x == t or f"{t}!" in x or f"'{t}'!" in x)
        for st in plan.get("steps", []):
            if st.get("op") in ("lookup_column", "lookup_join"):
                continue
            for t in bigs:
                if st.get("sheet") == t or mentions(st.get("args", {}), t):
                    raise ValueError(f"参考表《{t}》太大（几十万行），不能整表放进公式；请说“从《{t}》按××匹配××列”，"
                                     f"程序会用流式查找写入结果。/ Table {t!r} is too large for formulas; ask for a lookup instead.")

    def _stream_lookup(self, ws, title, key_ref, from_key, rets):
        """同一张大表、同一个匹配键的多个查找步骤共用一次扫描：第一次就把计划里所有要取的列一起取回。"""
        path, sheet = self.big_refs[title]
        kc = self._ci(key_ref)
        keys = {k for r in self._rows() if (k := norm_key(self._val(ws, r, kc))) is not None}
        ck = (title, from_key)
        cached = self._lookup_cache.get(ck)
        if cached is None or not keys <= cached[0] or not set(rets) <= set(cached[1]):
            cols = list(dict.fromkeys(list(rets) + self._planned_rets.get(ck, [])))
            with XlsxStream(path) as xs:
                top = xs.head(sheet, 12)
                hr = detect_header_row(top)
                heads = [str(h) if h is not None else "" for h in (top[hr - 1] if len(top) >= hr else [])]
                fk = col_index(heads, from_key)
                self.progress(f"stream-lookup {title}")
                hits = xs.lookup(sheet, fk, [col_index(heads, x) for x in cols], keys, header_row=hr, progress=self.progress)
            cached = self._lookup_cache[ck] = (keys, cols, hits)
        _, cols, hits = cached
        pos = [cols.index(x) for x in rets]
        return kc, {k: [v[i] for i in pos] for k, v in hits.items()}

    def op_lookup_column(self, step, a):
        if a.get("from_sheet") not in self.big_refs:
            return super().op_lookup_column(step, a)
        ws = self._sheet(step)
        kc, hits = self._stream_lookup(ws, a["from_sheet"], a["key"], a["from_key"], [a["return"]])
        nf = a.get("if_not_found", "")
        col = len(self._headers()) + 1
        ws.cell(self._hr(), col).value = a["name"]
        miss = 0
        for r in self._rows():
            v = hits.get(norm_key(self._val(ws, r, kc)))
            miss += v is None
            c = ws.cell(r, col)
            c.value = nf if v is None else v[0]
            if hasattr(c.value, "year"):
                c.number_format = "yyyy-mm-dd"
        return (f"column {a['name']} <- {a['from_sheet']}.{a['return']} ({miss} keys not found, "
                f"values: reference table too large for live formulas)")

    def op_lookup_join(self, step, a):
        if a.get("right") not in self.big_refs:
            return super().op_lookup_join(step, a)
        ws = self._sheet(step)
        kc, hits = self._stream_lookup(ws, a["right"], a["on"], a.get("right_on") or a["on"], list(a["fetch"]))
        base = ws.max_column
        for j, f in enumerate(a["fetch"], 1):
            ws.cell(self._hr(), base + j).value = f
        miss = 0
        for r in self._rows():
            v = hits.get(norm_key(self._val(ws, r, kc)))
            miss += v is None
            for j in range(len(a["fetch"])):
                ws.cell(r, base + j + 1).value = a.get("if_not_found", "") if v is None else v[j]
        return f"joined {len(a['fetch'])} cols from large table {a['right']!r} (streamed), {miss} keys not found"


def _generic_sheet_name(name: str, text: str) -> bool:
    """导出工具起的无意义表名（Sheet1 / 0 / 201_1_202606 / 带一长串十六进制的），用户不会这么称呼它。"""
    n = str(name).strip()
    if n and n in text and not re.fullmatch(r"[\d_\-. ]+", n):  # 纯数字表名（“0”）会在“2024”这类数字里误命中
        return False
    if re.fullmatch(r"(?i)(sheet|工作表|表)?\s*\d*", n) or re.search(r"[0-9a-f]{12,}", n, re.I):
        return True
    return not re.search(r"[\u4e00-\u9fff]", n)


def find_refs(wd: Path, cur_file: str, text: str, sheetnames):
    """指令里提到工作目录中其它文件名（去掉序号前缀）→ 作为参考表载入。返回 [(表名, 路径, 源表名)]。
    只读各文件的表名（workbook.xml），不打开工作表数据。
    表名：单表文件且原表名无意义（“0”“Sheet1”“201_1_202606”）时用文件名——用户在指令里就是这么叫它的，模型也会照着写。"""
    refs = []
    taken = set(sheetnames)
    for f in sorted(wd.glob("*.xlsx")):
        rel = str(f.relative_to(wd))
        if rel == cur_file or f.name.startswith("~$"):
            continue
        stem = re.sub(r"^\d+[_\-. ]*", "", f.stem)
        named = f.stem if f.stem in text else stem
        if not any(n and n in text for n in {stem, f.stem}):
            continue
        try:
            names = sheet_names(f)
        except Exception:  # noqa: BLE001
            continue
        for sh in names:
            if len(names) == 1 and _generic_sheet_name(sh, text):
                title = named[:31]
            else:
                title = sh if sh not in taken else stem[:31]
            if title in taken:
                continue  # 已在当前工作簿里（例如上一步已带入）
            taken.add(title)
            refs.append((title, str(f), sh))
    return refs


COMPACT_REF_FILES = 3      # 指令点名的参考文件超过这个数 → 参考表画像只保留被提到的列
COMPACT_EXTRA_CHARS = 3000  # 或者附加画像总长超过这个字数（模型上下文只有 4096 token）


def _mentioned(h, text: str) -> bool:
    """列名是否在指令里被提到：完整出现，或（长列名）任意连续 4 个字出现（“信用代码”命中“统一社会信用代码”）。"""
    h = str(h).strip()
    if len(h) < 2:
        return False
    if h in text:
        return True
    return len(h) >= 5 and any(h[i:i + 4] in text for i in range(len(h) - 3))


def compact_profile(p: TableProfile, text: str, keep_also=()) -> TableProfile:
    """精简画像：表名 + 行数 + 指令提到的列 + 与主表同名的列（多半是匹配键），1 行样例。都没命中时保留前 4 列。"""
    keep = [h for h in p.headers if _mentioned(h, text) or h in keep_also]
    if not keep:
        keep = p.headers[:4]
    idx = [p.headers.index(h) for h in keep]
    samples = [[r[i] if i < len(r) else None for i in idx] for r in p.samples[:1]]
    return TableProfile(p.sheet, keep, {h: p.types[h] for h in keep if h in p.types}, p.n_rows, samples)


def prompt_context(wd: Path, cur_file: str, sheet: str, text: str, progress=None):
    """返回 (主表画像, 其它表画像文本, 参考表列表)。与训练时的提示格式一致：主表 + 同簿其它表 + 参考表。
    必须在后台线程调用（界面线程只负责显示）。大文件一律流式读取，150MB 的表也在 1 秒内完成。
    指令牵涉的参考文件多于 3 个（或附加画像过长）时，除主表外的画像只保留被提到的列，避免上下文膨胀、注意力漂移。"""
    progress = progress or (lambda msg: None)
    path = wd / cur_file
    progress("main")
    if is_big(path):
        with XlsxStream(path) as xs:
            names = xs.sheet_names
            sheet = sheet if sheet in names else names[0]
            prof = stream_profile(xs, sheet, text=text)
            others = [stream_profile(xs, x) for x in names if x != sheet]
    else:
        wb = load_workbook(path)
        ws = wb[sheet] if sheet in wb.sheetnames else wb.worksheets[0]
        names = wb.sheetnames
        ev = Evaluator(wb)
        prof = profile_eval(ws, ev, text=text)  # 指令里明说“表头在第二行”时以它为准
        others = [profile_eval(x, ev) for x in wb.worksheets if x.title != ws.title]
    refs = find_refs(wd, cur_file, text, names)
    for title, rpath, sh in refs:
        progress("ref " + title)
        with XlsxStream(rpath) as xs:  # 参考表一律流式（不需要公式求值；3~4MB 的表 openpyxl 全量载入要 10 秒以上）
            p = stream_profile(xs, sh)
        p.sheet = title
        others.append(p)
    full = "\n".join(p.to_prompt() for p in others)
    if len({r[1] for r in refs}) > COMPACT_REF_FILES or len(full) > COMPACT_EXTRA_CHARS:
        others = [compact_profile(p, text, prof.headers) for p in others]
    return prof, "\n".join(p.to_prompt() for p in others), refs


def mark_big_lookups(plan: dict, refs) -> dict:
    """对超大参考表的查找，执行器会写结果值而不是活公式；计划预览要如实显示。"""
    bigs = {t for t, p, sh in refs if is_big_ref(p, sh)}
    for st in plan.get("steps", []):
        if st.get("op") == "lookup_column" and st.get("args", {}).get("from_sheet") in bigs:
            st["args"]["mode"] = "value"
    return plan


_DEFAULT_OUT = re.compile(r"^(?P<base>.*?)(?:_(?:处理后|processed))+(?:\(\d+\))?$")


def tidy_save_paths(plan: dict, cur_file: str, wd: Path) -> dict:
    """链式处理时模型会把默认输出名越叠越长（X_处理后_处理后_处理后.xlsx）。
    默认风格的输出名统一成 <原名>_处理后.xlsx，已存在则 <原名>_处理后(2).xlsx …；
    用户明确要求覆盖（保存路径 == 当前输入）或自定义的文件名不动。"""
    for st in plan.get("steps", []):
        if st.get("op") != "save":
            continue
        path = st.get("args", {}).get("path", "")
        if not path or path == cur_file:
            continue
        p = Path(path)
        m = _DEFAULT_OUT.match(p.stem)
        if not m or str(p.parent) not in ("", "."):
            continue
        base = _DEFAULT_OUT.match(Path(cur_file).stem)
        base = base.group("base") if base else Path(cur_file).stem
        suffix = "_processed" if "_processed" in p.stem else "_处理后"
        cand, n = f"{base}{suffix}{p.suffix}", 1
        while cand == cur_file or (wd / cand).exists():
            n += 1
            cand = f"{base}{suffix}({n}){p.suffix}"
        st["args"]["path"] = cand
    return plan
