"""超大 .xlsx 的流式轻量读取（纯标准库：zipfile + 正则，不依赖 openpyxl / lxml）。

为什么不用 openpyxl 的 read_only 模式：它在打开文件时就把 xl/sharedStrings.xml 整个解析进内存。
实测一张真实政务名录导出表（150MB，76 万行 × 25 列，共享字符串 335 万条 / 解压后 165MB）
load_workbook(read_only=True) 本身就要 67 秒。这里的做法：
  - 表名 / 行列数：只读 workbook.xml 和工作表 XML 开头的 <dimension>（毫秒级）
  - 表头 + 前几十行样例：只解压工作表开头一小段；共享字符串“用到哪条才读到哪条”
  - 跨表匹配（lookup）：三遍流式扫描，内存里只保留主表需要的那些键，不把大表整体载入
"""
from __future__ import annotations

import datetime as dt
import html
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import PurePosixPath

CHUNK = 1 << 22  # 每次解压 4MB
_NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

_SI_RX = re.compile(rb"<si>(.*?)</si>|<si\s*/>", re.S)
_T_RX = re.compile(rb"<t(?:\s[^>]*)?>(.*?)</t>|<t\s*/>", re.S)
_RPH_RX = re.compile(rb"<rPh\b.*?</rPh>", re.S)
_ROW_RX = re.compile(rb"<row\b([^>]*?)(?:/>|>(.*?)</row>)", re.S)
_CELL_RX = re.compile(rb"<c\b([^>]*?)(?:/>|>(.*?)</c>)", re.S)
_ATTR_RX = re.compile(rb'\b(r|t|s)="([^"]*)"')
_V_RX = re.compile(rb"<v>(.*?)</v>", re.S)
_XESC_RX = re.compile(r"_x([0-9A-Fa-f]{4})_")
_DATE_IDS = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))


def _text(b: bytes) -> str:
    s = b.decode("utf-8", "replace")
    if "&" in s:
        s = html.unescape(s)
    if "_x" in s:
        s = _XESC_RX.sub(lambda m: chr(int(m.group(1), 16)), s)
    return s


def _si_text(body: bytes | None) -> str:
    if not body:
        return ""
    if b"<rPh" in body:
        body = _RPH_RX.sub(b"", body)
    return "".join(_text(m.group(1)) for m in _T_RX.finditer(body) if m.group(1) is not None)


def col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def norm_key(v):
    """匹配用的键：数字 123 / 123.0 / 文本 "123" 视为同一个（政企导出表里代码列常常一半是文本一半是数字）。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (int, float)):
        return repr(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return v.isoformat()
    s = str(v).strip()
    return s or None


def _serial_to_dt(x: float):
    try:
        return dt.datetime(1899, 12, 30) + dt.timedelta(milliseconds=round(float(x) * 86_400_000))
    except (OverflowError, ValueError):
        return x


def _iter_chunks(zf: zipfile.ZipFile, name: str, sep: bytes):
    """按完整的 XML 元素切块流式解压：每块都以 sep（如 </row>）结尾。"""
    buf = b""
    with zf.open(name) as f:
        while True:
            data = f.read(CHUNK)
            if not data:
                break
            buf += data
            cut = buf.rfind(sep)
            if cut < 0:
                continue
            cut += len(sep)
            yield buf[:cut]
            buf = buf[cut:]
    if buf:
        yield buf


class XlsxStream:
    def __init__(self, path):
        self.path = str(path)
        self.zf = zipfile.ZipFile(self.path)
        self._names = set(self.zf.namelist())
        self.sheets = self._read_workbook()          # [(名称, 部件路径)]
        self._date_styles = None
        self._sst_name = next((n for n in ("xl/sharedStrings.xml", "xl/sharedstrings.xml") if n in self._names), None)
        self._sst_cache: dict[int, str] = {}
        self._sst_upto = -1                          # 已顺序读到的最大下标

    def close(self):
        self.zf.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # ------------------------------------------------------------------ 结构
    def _read_workbook(self):
        wb = ET.fromstring(self.zf.read("xl/workbook.xml"))
        rels = {}
        rel_name = "xl/_rels/workbook.xml.rels"
        if rel_name in self._names:
            for r in ET.fromstring(self.zf.read(rel_name)):
                tgt = r.get("Target", "")
                tgt = tgt.lstrip("/") if tgt.startswith("/") else str(PurePosixPath("xl") / tgt)
                rels[r.get("Id")] = tgt
        out = []
        for i, sh in enumerate(wb.iter(_NS_MAIN + "sheet"), 1):
            part = rels.get(sh.get(_NS_REL + "id")) or f"xl/worksheets/sheet{i}.xml"
            out.append((sh.get("name"), part))
        return out

    @property
    def sheet_names(self):
        return [n for n, _ in self.sheets]

    def _part(self, sheet):
        if sheet is None:
            return self.sheets[0][1]
        for n, p in self.sheets:
            if n == sheet:
                return p
        raise KeyError(f"工作表不存在：{sheet}")

    def dims(self, sheet=None):
        """(最大行, 最大列)，取自工作表开头的 <dimension ref="A1:Y761508"/>；没有该标记时返回 (None, None)。"""
        with self.zf.open(self._part(sheet)) as f:
            head = f.read(8192)
        m = re.search(rb'<dimension ref="(?:[A-Z]+\d+:)?([A-Z]+)(\d+)"', head)
        if not m:
            return None, None
        return int(m.group(2)), col_num(m.group(1).decode())

    def date_styles(self):
        if self._date_styles is None:
            ds = set()
            if "xl/styles.xml" in self._names:
                st = ET.fromstring(self.zf.read("xl/styles.xml"))
                custom = {}
                nf = st.find(_NS_MAIN + "numFmts")
                for x in (nf if nf is not None else []):
                    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", x.get("formatCode", "")).lower()
                    custom[int(x.get("numFmtId"))] = bool(re.search(r"[dmyh]", code)) and "general" not in code
                xfs = st.find(_NS_MAIN + "cellXfs")
                for i, xf in enumerate(xfs if xfs is not None else []):
                    fid = int(xf.get("numFmtId", 0))
                    if fid in _DATE_IDS or custom.get(fid):
                        ds.add(i)
            self._date_styles = ds
        return self._date_styles

    # ------------------------------------------------------------------ 共享字符串
    def strings(self, indices) -> dict[int, str]:
        """取若干条共享字符串：从头顺序读到需要的最大下标为止（样例行用到的通常都在很前面）。"""
        need = {i for i in indices if i not in self._sst_cache}
        if need and self._sst_name and max(need) > self._sst_upto:
            target, k = max(need), -1
            for chunk in _iter_chunks(self.zf, self._sst_name, b"</si>"):
                for m in _SI_RX.finditer(chunk):
                    k += 1
                    if k > self._sst_upto:
                        self._sst_cache[k] = _si_text(m.group(1))
                    if k >= target:
                        break
                if k >= target:
                    break
            self._sst_upto = max(self._sst_upto, k)
        return {i: self._sst_cache.get(i, "") for i in indices}

    def find_strings(self, wanted: set[str]) -> dict[int, str]:
        """整遍流式扫描共享字符串，只记下“内容（去首尾空格）在 wanted 里”的那些下标。内存只与 wanted 大小有关。
        快速路径：绝大多数条目是 <si><t>纯文本</t></si>，直接按字节比较，不解码。"""
        out = {}
        if not self._sst_name or not wanted:
            return out
        wb = {html.escape(w, quote=False).encode(): w for w in wanted}
        k = -1
        for chunk in _iter_chunks(self.zf, self._sst_name, b"</si>"):
            for m in _SI_RX.finditer(chunk):
                k += 1
                body = m.group(1)
                if body is None:
                    continue
                if body[:2] == b"<t" and body[-4:] == b"</t>" and body.count(b"<") == 2:  # <t> 或 <t xml:space="preserve">
                    hit = wb.get(body[body.index(b">") + 1:-4].strip())
                    if hit is not None:
                        out[k] = hit
                    continue
                s = _si_text(body).strip()
                if s in wanted:
                    out[k] = s
        return out

    # ------------------------------------------------------------------ 单元格
    def _cell(self, attrs: bytes, body: bytes | None):
        """返回 (列号或 None, 原始值)；共享字符串返回 ("s", 下标) 待解析。"""
        a = dict(_ATTR_RX.findall(attrs))
        ref = a.get(b"r")
        col = col_num(re.match(rb"[A-Z]+", ref).group().decode()) if ref else None
        if body is None:
            return col, None
        t = a.get(b"t", b"n")
        if t == b"inlineStr":
            return col, _si_text(body)
        m = _V_RX.search(body)
        if not m:
            return col, None
        v = m.group(1)
        if t == b"s":
            return col, ("s", int(v))
        if t in (b"str", b"e"):
            return col, _text(v)
        if t == b"b":
            return col, v == b"1"
        if t == b"d":
            try:
                return col, dt.datetime.fromisoformat(v.decode())
            except ValueError:
                return col, v.decode()
        try:
            x = float(v)
        except ValueError:
            return col, _text(v)
        if int(a.get(b"s", b"0") or 0) in self.date_styles():
            return col, _serial_to_dt(x)
        return col, int(x) if x.is_integer() and abs(x) < 1e15 else x

    def iter_rows(self, sheet=None):
        """逐行产出 (行号, {列号: 原始值})；共享字符串仍是 ("s", 下标)。"""
        last = 0
        for chunk in _iter_chunks(self.zf, self._part(sheet), b"</row>"):
            for m in _ROW_RX.finditer(chunk):
                rm = re.search(rb'\br="(\d+)"', m.group(1))
                r = int(rm.group(1)) if rm else last + 1
                last = r
                vals, c_auto = {}, 0
                for cm in _CELL_RX.finditer(m.group(2) or b""):
                    c, v = self._cell(cm.group(1), cm.group(2))
                    c = c or c_auto + 1
                    c_auto = c
                    if v is not None:
                        vals[c] = v
                yield r, vals

    def head(self, sheet=None, n_rows: int = 50):
        """前 n_rows 行（第 1 行起，按行号对齐，空行补空列表），共享字符串已解析成文本。"""
        rows: dict[int, list] = {}
        for r, vals in self.iter_rows(sheet):
            if r > n_rows:
                break
            rows[r] = [vals.get(i) for i in range(1, max(vals, default=0) + 1)]
        idx = {v[1] for row in rows.values() for v in row if isinstance(v, tuple)}
        sst = self.strings(idx) if idx else {}
        out = []
        for r in range(1, min(n_rows, max(rows, default=0)) + 1):
            out.append([sst[v[1]] if isinstance(v, tuple) else v for v in rows.get(r, [])])
        return out

    def lookup(self, sheet, key_col: int, ret_cols: list[int], keys: set[str], header_row: int = 1, progress=None):
        """在大表里按 key_col 匹配 keys（norm_key 之后的字符串），返回 {键: [各返回列的值]}，同键取第一次出现。

        三遍流式：① 共享字符串里找出键对应的下标；② 扫工作表，只看键列和返回列；③ 解析命中行返回值里的共享字符串。
        内存只与“主表的键数量”有关，与大表行数无关。"""
        progress = progress or (lambda msg: None)
        progress("index-strings")
        key_idx = self.find_strings(keys)
        want = {key_col, *ret_cols}
        hits: dict[str, list] = {}
        progress("scan-rows")
        for r, cur in (self._scan_cols(sheet, want) if self.has_cell_refs(sheet) else self.iter_rows(sheet)):
            self._take(cur, r, header_row, key_col, ret_cols, key_idx, keys, hits)
        progress("resolve")
        need = {v[1] for vals in hits.values() for v in vals if isinstance(v, tuple)}
        sst = self._resolve_scattered(need)
        return {k: [sst.get(v[1], "") if isinstance(v, tuple) else v for v in vals] for k, vals in hits.items()}

    def _scan_cols(self, sheet, cols):
        """快速路径：正则只抓指定列的单元格（依赖 r="A1" 坐标），逐行产出 (行号, {列号: 原始值})。
        最常见的两种单元格（共享字符串 / 无样式数字，正文只有 <v>）由正则直接取值，其余交给通用的 _cell。"""
        lm = {col_letters(c).encode(): c for c in cols}
        cell_rx = re.compile(rb'<c r="(' + b"|".join(sorted(lm)) + rb')(\d+)"([^>]*?)(?:/>|><v>([^<]*)</v></c>|>(.*?)</c>)', re.S)
        dates = self.date_styles()
        s_rx = re.compile(rb'\bs="(\d+)"')
        cell = self._cell
        for chunk in _iter_chunks(self.zf, self._part(sheet), b"</row>"):
            cur_r, cur = None, {}
            for L, rs, attrs, v, body in cell_rx.findall(chunk):
                if rs != cur_r:
                    if cur_r is not None:
                        yield int(cur_r), cur
                    cur_r, cur = rs, {}
                if v:
                    if b't="s"' in attrs:
                        cur[lm[L]] = ("s", int(v))
                        continue
                    if b"t=" not in attrs:
                        sm = s_rx.search(attrs) if dates and b"s=" in attrs else None
                        if sm is None or int(sm.group(1)) not in dates:
                            try:
                                x = float(v)
                                cur[lm[L]] = int(x) if x.is_integer() and abs(x) < 1e15 else x
                                continue
                            except ValueError:
                                pass
                    cur[lm[L]] = cell(attrs, b"<v>" + v + b"</v>")[1]
                elif body:
                    cur[lm[L]] = cell(attrs, body)[1]
                else:
                    cur[lm[L]] = None
            if cur_r is not None:
                yield int(cur_r), cur

    @staticmethod
    def _take(cur, r, header_row, key_col, ret_cols, key_idx, keys, hits):
        if r <= header_row:
            return
        kv = cur.get(key_col)
        if isinstance(kv, tuple):
            k = key_idx.get(kv[1])
        else:
            k = norm_key(kv)
            k = k if k in keys else None
        if k is not None and k not in hits:
            hits[k] = [cur.get(c) for c in ret_cols]

    def _resolve_scattered(self, need: set[int]) -> dict[int, str]:
        """下标分散在整个共享字符串表里：一遍流式扫描，只留需要的那些。"""
        out = {i: self._sst_cache[i] for i in need if i in self._sst_cache}
        need = need - set(out)
        if not need or not self._sst_name:
            return out
        top, k = max(need), -1
        for chunk in _iter_chunks(self.zf, self._sst_name, b"</si>"):
            for m in _SI_RX.finditer(chunk):
                k += 1
                if k in need:
                    out[k] = _si_text(m.group(1))
            if k >= top:
                break
        return out

    def has_cell_refs(self, sheet=None) -> bool:
        """单元格是否带 r="A1" 坐标（Excel / WPS / openpyxl 都会写；lookup 的快速正则依赖它）。"""
        with self.zf.open(self._part(sheet)) as f:
            head = f.read(65536)
        m = _CELL_RX.search(head)
        return bool(m and re.match(rb'\s*r="', m.group(1)))
