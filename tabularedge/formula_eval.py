"""轻量 Excel 公式求值器（执行器内部用）。

为什么需要：openpyxl 只写公式不算值。执行器在“新增公式列之后再按它分组 / 筛选 / 去重 / 拆分文件”时，
必须知道公式的结果，否则会按公式文本分组（每行都不同）。这里实现模型常写的 ~80 个函数；
遇到不支持的函数抛 Unsupported，调用方回退到原行为（保留公式文本）。

写进 Excel 的仍然是活公式；这里算出的值只用于执行器自己的决策与界面预览。
"""
from __future__ import annotations

import datetime as dt
import math
import re

from openpyxl.utils import column_index_from_string


class Unsupported(Exception):
    pass


class XLError:
    def __init__(self, code):
        self.code = code

    def __repr__(self):
        return self.code

    def __eq__(self, o):
        return isinstance(o, XLError) and o.code == self.code

    def __hash__(self):
        return hash(self.code)


NA, VALUE, DIV0, REF, NUM = (XLError(c) for c in ("#N/A", "#VALUE!", "#DIV/0!", "#REF!", "#NUM!"))
EPOCH = dt.datetime(1899, 12, 30)


def to_serial(d):
    if isinstance(d, dt.datetime):
        return (d - EPOCH).total_seconds() / 86400
    if isinstance(d, dt.date):
        return float((d - EPOCH.date()).days)
    return d


def from_serial(x) -> dt.datetime:
    return EPOCH + dt.timedelta(days=float(x))


# ------------------------------------------------------------------ tokenizer
_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_一-鿿][\w一-鿿.]*)!"
_CELL = r"\$?[A-Z]{1,3}\$?\d+"
_TOK = re.compile(
    r"\s*(?:"
    r"(?P<str>\"(?:[^\"]|\"\")*\")"
    rf"|(?P<ref>(?:{_SHEET})?(?:{_CELL}(?::{_CELL})?|\$?[A-Z]{{1,3}}:\$?[A-Z]{{1,3}}|\$?\d+:\$?\d+))(?![\w(])"
    r"|(?P<num>\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+)"
    r"|(?P<func>[A-Za-z_][A-Za-z0-9_.]*)\s*\("
    r"|(?P<bool>TRUE|FALSE)(?![\w(])"
    r"|(?P<op><>|<=|>=|[-+*/^&=<>%(),:])"
    r")")


def tokenize(s):
    s = s.strip()
    if s.startswith("="):
        s = s[1:]
    pos, out = 0, []
    while pos < len(s):
        if s[pos:].strip() == "":
            break
        m = _TOK.match(s, pos)
        if not m:
            raise Unsupported(f"cannot tokenize at {s[pos:pos + 20]!r}")
        kind = m.lastgroup
        out.append((kind, m.group(kind)))
        pos = m.end()
    return out


# ------------------------------------------------------------------ parser -> AST (tuples)
class _P:
    PREC = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}

    def __init__(self, toks):
        self.t, self.i = toks, 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expr(self, minp=1):
        left = self.unary()
        while True:
            k, v = self.peek()
            if k != "op" or v not in self.PREC or self.PREC[v] < minp:
                return left
            self.take()
            right = self.expr(self.PREC[v] + (0 if v == "^" else 1))
            left = ("bin", v, left, right)

    def unary(self):
        k, v = self.peek()
        if k == "op" and v in "+-":
            self.take()
            return ("neg", self.unary()) if v == "-" else self.unary()
        node = self.atom()
        while self.peek() == ("op", "%"):
            self.take()
            node = ("bin", "/", node, ("lit", 100.0))
        return node

    def atom(self):
        k, v = self.take()
        if k == "num":
            return ("lit", float(v))
        if k == "str":
            return ("lit", v[1:-1].replace('""', '"'))
        if k == "bool":
            return ("lit", v == "TRUE")
        if k == "ref":
            return ("ref", v)
        if k == "func":
            name = v.upper()
            for p in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
                if name.startswith(p):
                    name = name[len(p):]
            args = []
            if self.peek() == ("op", ")"):
                self.take()
                return ("call", name, args)
            while True:
                if self.peek() in (("op", ","), ("op", ")")):
                    args.append(("lit", None))  # 省略的参数
                else:
                    args.append(self.expr())
                k2, v2 = self.take()
                if v2 == ")":
                    return ("call", name, args)
                if v2 != ",":
                    raise Unsupported(f"expected , or ) got {v2!r}")
        if (k, v) == ("op", "("):
            e = self.expr()
            if self.take() != ("op", ")"):
                raise Unsupported("unbalanced (")
            return e
        raise Unsupported(f"unexpected token {v!r}")


def parse(formula):
    p = _P(tokenize(formula))
    node = p.expr()
    if p.i != len(p.t):
        raise Unsupported(f"trailing tokens {p.t[p.i:]}")
    return node


# ------------------------------------------------------------------ helpers
class Rng:
    """二维区域（行优先）。"""

    def __init__(self, rows):
        self.rows = rows

    def flat(self):
        return [v for r in self.rows for v in r]

    @property
    def shape(self):
        return len(self.rows), len(self.rows[0]) if self.rows else 0


def _flat(args):
    for a in args:
        if isinstance(a, Rng):
            yield from a.flat()
        elif isinstance(a, list):
            yield from a
        else:
            yield a


def _num(v):
    if isinstance(v, XLError):
        raise _Err(v)
    if v is None or v == "":
        return 0.0
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, (dt.date, dt.datetime)):
        return to_serial(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        d = _parse_date(str(v))
        if d is not None:
            return to_serial(d)
        raise _Err(VALUE)


def _str(v):
    if isinstance(v, XLError):
        raise _Err(v)
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, (dt.date, dt.datetime)):
        return _str(to_serial(v))
    return str(v)


def _bool(v):
    if isinstance(v, XLError):
        raise _Err(v)
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        raise _Err(VALUE)
    return bool(_num(v))


class _Err(Exception):
    def __init__(self, e):
        self.e = e


_DATE_PATS = [
    (re.compile(r"^\s*(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*$"), "ymd"),
    (re.compile(r"^\s*(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?\s*$"), "mdy"),
    (re.compile(r"^\s*(\d{4})(\d{2})(\d{2})\s*$"), "ymd"),
]


def _parse_date(s):
    for rx, kind in _DATE_PATS:
        m = rx.match(s)
        if m:
            g = [int(x) if x else 0 for x in m.groups()] + [0, 0, 0]
            y, mo, d = (g[0], g[1], g[2]) if kind == "ymd" else (g[2], g[0], g[1])
            try:
                return dt.datetime(y, mo, d, g[3], g[4], g[5])
            except ValueError:
                return None
    return None


def _cmp_key(v):
    """Excel 排序/比较：数字 < 文本(不区分大小写) < 逻辑值。"""
    if v is None:
        return (0, 0.0)
    if isinstance(v, bool):
        return (2, float(v))
    if isinstance(v, (int, float)):
        return (0, float(v))
    if isinstance(v, (dt.date, dt.datetime)):
        return (0, to_serial(v))
    return (1, str(v).lower())


def _eq(a, b):
    if a is None:
        a = "" if isinstance(b, str) else 0.0
    if b is None:
        b = "" if isinstance(a, str) else 0.0
    return _cmp_key(a) == _cmp_key(b)


def _criteria(c):
    """COUNTIF/SUMIFS 条件 -> 谓词。"""
    if isinstance(c, (int, float)) and not isinstance(c, bool):
        return lambda v: isinstance(v, (int, float)) and not isinstance(v, bool) and float(v) == float(c)
    s = _str(c)
    m = re.match(r"^(<>|<=|>=|=|<|>)(.*)$", s, re.S)
    op, rest = (m.group(1), m.group(2)) if m else ("=", s)
    try:
        x = float(rest)
        isnum = True
    except ValueError:
        x, isnum = rest, False
    if isnum and op in ("<", "<=", ">", ">=", "=", "<>"):
        def num_pred(v):
            if isinstance(v, (dt.date, dt.datetime)):
                v = to_serial(v)
            if not isinstance(v, (int, float)) or isinstance(v, bool):
                return op == "<>"
            return {"<": v < x, "<=": v <= x, ">": v > x, ">=": v >= x, "=": v == x, "<>": v != x}[op]
        return num_pred
    if op in ("<", "<=", ">", ">="):
        return lambda v: isinstance(v, str) and {"<": v.lower() < x.lower(), "<=": v.lower() <= x.lower(),
                                                 ">": v.lower() > x.lower(), ">=": v.lower() >= x.lower()}[op]
    if x == "":
        return (lambda v: v not in (None, "")) if op == "<>" else (lambda v: v in (None, ""))
    rx = re.compile("^" + re.escape(x).replace(r"\*", ".*").replace(r"\?", ".") .replace("~.*", r"\*") + "$", re.I | re.S)
    hit = lambda v: v is not None and rx.match(_str(v)) is not None
    return (lambda v: not hit(v)) if op == "<>" else hit


def _ifs_pairs(args):
    if len(args) % 2:
        raise _Err(VALUE)
    pairs = [(args[i], _criteria(args[i + 1])) for i in range(0, len(args), 2)]
    return pairs


def _mask(pairs, n):
    out = []
    for i in range(n):
        ok = True
        for rng, pred in pairs:
            fl = rng.flat() if isinstance(rng, Rng) else [rng]
            if i >= len(fl) or not pred(fl[i]):
                ok = False
                break
        out.append(ok)
    return out


def _text_fmt(v, fmt):
    f = fmt.lower()
    if isinstance(v, str):
        try:
            v = _num(v)
        except _Err:
            return v
    if any(t in f for t in ("yy", "dd", "m")) and not re.fullmatch(r"[0#,.%]+", f):
        d = from_serial(_num(v))
        out = fmt
        for pat, rep in (("yyyy", f"{d.year:04d}"), ("YYYY", f"{d.year:04d}"), ("yy", f"{d.year % 100:02d}"),
                         ("mm", f"{d.month:02d}"), ("MM", f"{d.month:02d}"), ("dd", f"{d.day:02d}"),
                         ("DD", f"{d.day:02d}"), ("m", str(d.month)), ("M", str(d.month)), ("d", str(d.day)), ("D", str(d.day))):
            out = out.replace(pat, rep)
        return out
    x = _num(v)
    if f.endswith("%"):
        dec = len(f.split(".")[1]) - 1 if "." in f else 0
        return f"{x * 100:.{dec}f}%"
    dec = len(f.split(".")[1]) if "." in f else 0
    return f"{x:,.{dec}f}" if "," in f else f"{x:.{dec}f}"


def _lookup_pos(key, arr, mode=0, search=1):
    """返回 0-based 位置或 None。mode: 0 精确 / -1 精确或下一个较小 / 1 精确或下一个较大 / 2 通配。"""
    idx = range(len(arr)) if search >= 0 else range(len(arr) - 1, -1, -1)
    if mode == 2:
        pred = _criteria(_str(key))
        return next((i for i in idx if pred(arr[i])), None)
    for i in idx:
        if _eq(arr[i], key) and arr[i] is not None:
            return i
    if mode == 0:
        return None
    kk = _cmp_key(key)
    best = None
    for i, v in enumerate(arr):
        if v is None or _cmp_key(v)[0] != kk[0]:
            continue
        if (mode == -1 and _cmp_key(v) < kk and (best is None or _cmp_key(v) > _cmp_key(arr[best]))) or \
           (mode == 1 and _cmp_key(v) > kk and (best is None or _cmp_key(v) < _cmp_key(arr[best]))):
            best = i
    return best


def _round(x, n, how):
    q = 10 ** n
    v = x * q
    if how == "half":
        r = math.floor(abs(v) + 0.5 + 1e-9)
    elif how == "up":
        r = math.ceil(abs(v) - 1e-9)
    else:
        r = math.floor(abs(v) + 1e-9)
    return math.copysign(r, v) / q


def _weeknum(serial, rt=1):
    d = from_serial(serial).date()
    if rt == 21:
        return d.isocalendar()[1]
    start = 1 if rt in (1, 17) else 2  # 1=周日开始；2/11=周一开始
    jan1 = dt.date(d.year, 1, 1)
    wd = (jan1.weekday() + 1) % 7 if start == 1 else jan1.weekday()  # 0=周首
    return (d - jan1).days // 7 + (1 if ((d - jan1).days % 7 + wd) < 7 else 2)


def _percentrank(arr, x, sig=3):
    vals = sorted(float(v) for v in arr if isinstance(v, (int, float)) and not isinstance(v, bool))
    if not vals or x < vals[0] or x > vals[-1]:
        raise _Err(NA)
    n = len(vals)
    if n == 1:
        return 1.0
    less = sum(1 for v in vals if v < x)
    if x in vals:
        r = less / (n - 1)
    else:
        lo = max(v for v in vals if v < x)
        hi = min(v for v in vals if v > x)
        r = (less - 1 + (x - lo) / (hi - lo)) / (n - 1)
    q = 10 ** sig
    return math.floor(r * q + 1e-9) / q


def _percentile(arr, k):
    vals = sorted(float(v) for v in arr if isinstance(v, (int, float)) and not isinstance(v, bool))
    if not vals:
        raise _Err(NUM)
    h = (len(vals) - 1) * k
    lo = int(math.floor(h))
    return vals[lo] + (h - lo) * (vals[min(lo + 1, len(vals) - 1)] - vals[lo])


def _nums(args):
    return [float(v) if not isinstance(v, (dt.date, dt.datetime)) else to_serial(v)
            for v in _flat(args) if isinstance(v, (int, float, dt.date, dt.datetime)) and not isinstance(v, bool)]


def _raise_err(args):
    for v in _flat(args):
        if isinstance(v, XLError):
            raise _Err(v)


# ------------------------------------------------------------------ evaluator
class Evaluator:
    def __init__(self, wb, max_depth=40):
        self.wb = wb
        self.cache: dict = {}
        self.depth = 0
        self.max_depth = max_depth

    def invalidate(self):
        self.cache.clear()

    # 单元格值（公式单元格递归求值）
    def cell(self, ws, r, c):
        key = (ws.title, r, c)
        if key in self.cache:
            return self.cache[key]
        v = ws.cell(r, c).value
        if hasattr(v, "text") and hasattr(v, "ref"):  # ArrayFormula：不在此求值
            raise Unsupported("array formula")
        if isinstance(v, str) and v.startswith("="):
            self.depth += 1
            if self.depth > self.max_depth:
                self.depth -= 1
                raise Unsupported("depth")
            try:
                v = self.eval(v, ws, r)
            finally:
                self.depth -= 1
        elif isinstance(v, (dt.date, dt.datetime)):
            v = to_serial(v)
        self.cache[key] = v
        return v

    def eval(self, formula, ws, row=None):
        """公式 -> 值（XLError 也作为值返回）。不支持的函数抛 Unsupported。"""
        node = parse(formula)
        try:
            v = self._ev(node, ws)
        except _Err as e:
            return e.e
        if isinstance(v, Rng):
            fl = v.flat()
            v = fl[0] if fl else None
        return v

    def _sheet(self, ref, ws):
        if "!" in ref:
            name, addr = ref.rsplit("!", 1)
            name = name[1:-1].replace("''", "'") if name.startswith("'") else name
            if name not in self.wb.sheetnames:
                raise _Err(REF)
            return self.wb[name], addr
        return ws, ref

    def _ref(self, ref, ws):
        sh, addr = self._sheet(ref, ws)
        addr = addr.replace("$", "")
        if ":" in addr:
            a, b = addr.split(":")
        else:
            a = b = addr
        def split(x):
            m = re.fullmatch(r"([A-Z]{0,3})(\d*)", x)
            return m.group(1), m.group(2)
        ca, ra = split(a)
        cb, rb = split(b)
        c1 = column_index_from_string(ca) if ca else 1
        c2 = column_index_from_string(cb) if cb else max(sh.max_column, 1)
        r1 = int(ra) if ra else 1
        r2 = int(rb) if rb else max(sh.max_row, 1)
        if not ra:  # 整列：去掉表头行之外不受影响；Excel 整列也包含表头，这里保持一致
            r2 = max(sh.max_row, 1)
        if r1 == r2 and c1 == c2:
            return self.cell(sh, r1, c1)
        return Rng([[self.cell(sh, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)])

    def _ev(self, n, ws):
        t = n[0]
        if t == "lit":
            return n[1]
        if t == "ref":
            return self._ref(n[1], ws)
        if t == "neg":
            v = self._ev(n[1], ws)
            if isinstance(v, Rng):
                return Rng([[-_num(x) for x in row] for row in v.rows])
            return -_num(v)
        if t == "bin":
            return self._bin(n[1], self._ev(n[2], ws), self._ev(n[3], ws))
        if t == "call":
            return self._call(n[1], n[2], ws)
        raise Unsupported(t)

    def _bin(self, op, a, b):
        if isinstance(a, Rng) or isinstance(b, Rng):  # 数组运算（SUMPRODUCT/FILTER 条件里常见）
            A = a.flat() if isinstance(a, Rng) else None
            B = b.flat() if isinstance(b, Rng) else None
            n = len(A) if A is not None else len(B)
            shape_src = a if isinstance(a, Rng) else b
            vals = []
            for i in range(n):
                try:
                    vals.append(self._bin(op, A[i] if A is not None else a, B[i] if B is not None else b))
                except _Err as e:
                    vals.append(e.e)
            rows, cols = shape_src.shape
            return Rng([vals[i * cols:(i + 1) * cols] for i in range(rows)])
        if isinstance(a, XLError):
            raise _Err(a)
        if isinstance(b, XLError):
            raise _Err(b)
        if op == "&":
            return _str(a) + _str(b)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            if a is None:
                a = "" if isinstance(b, str) else 0.0
            if b is None:
                b = "" if isinstance(a, str) else 0.0
            ka, kb = _cmp_key(a), _cmp_key(b)
            return {"=": ka == kb, "<>": ka != kb, "<": ka < kb, ">": ka > kb, "<=": ka <= kb, ">=": ka >= kb}[op]
        x, y = _num(a), _num(b)
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            if y == 0:
                raise _Err(DIV0)
            return x / y
        if op == "^":
            return x ** y
        raise Unsupported(op)

    # 惰性求值的函数
    def _call(self, name, args, ws):
        ev = lambda i: self._ev(args[i], ws) if i < len(args) else None
        if name == "IF":
            c = ev(0)
            if isinstance(c, Rng):
                raise Unsupported("array IF")
            if _bool(c):
                return ev(1) if len(args) > 1 else True
            return ev(2) if len(args) > 2 else False
        if name in ("IFERROR", "IFNA"):
            try:
                v = ev(0)
            except _Err as e:
                v = e.e
            if isinstance(v, XLError) and (name == "IFERROR" or v == NA):
                return ev(1)
            return v
        if name == "IFS":
            for i in range(0, len(args) - 1, 2):
                if _bool(ev(i)):
                    return ev(i + 1)
            raise _Err(NA)
        if name == "SWITCH":
            x = ev(0)
            i = 1
            while i + 1 < len(args):
                if _eq(ev(i), x):
                    return ev(i + 1)
                i += 2
            if i < len(args):
                return ev(i)
            raise _Err(NA)
        if name == "LET":
            raise Unsupported("LET")
        vals = [self._ev(a, ws) for a in args]
        f = _FUNCS.get(name)
        if f is None:
            raise Unsupported(f"function {name}")
        return f(*vals)


def _sumifs(sr, *pairs):
    fl = sr.flat() if isinstance(sr, Rng) else [sr]
    m = _mask(_ifs_pairs(list(pairs)), len(fl))
    return sum(float(v) for v, ok in zip(fl, m) if ok and isinstance(v, (int, float)) and not isinstance(v, bool))


def _sel(sr, pairs):
    fl = sr.flat() if isinstance(sr, Rng) else [sr]
    m = _mask(_ifs_pairs(list(pairs)), len(fl))
    return [float(v) for v, ok in zip(fl, m) if ok and isinstance(v, (int, float)) and not isinstance(v, bool)]


def _xlookup(key, look, ret, nf=None, mode=0, search=1):
    arr = look.flat() if isinstance(look, Rng) else [look]
    i = _lookup_pos(key, arr, int(_num(mode)) if mode is not None else 0, int(_num(search)) if search is not None else 1)
    if i is None:
        if nf is None:
            raise _Err(NA)
        return nf
    rv = ret.flat() if isinstance(ret, Rng) else [ret]
    if isinstance(ret, Rng) and ret.shape[0] == len(arr) and ret.shape[1] > 1:
        return Rng([ret.rows[i]])
    return rv[i] if i < len(rv) else None


def _vlookup(key, table, col, approx=True):
    col = int(_num(col))
    first = [row[0] for row in table.rows]
    approx = True if approx is None else _bool(approx)
    i = _lookup_pos(key, first, -1 if approx else 0)
    if i is None:
        raise _Err(NA)
    return table.rows[i][col - 1]


def _match(key, arr, mode=1):
    fl = arr.flat() if isinstance(arr, Rng) else [arr]
    mode = 1 if mode is None else int(_num(mode))
    i = _lookup_pos(key, fl, {0: 0, 1: -1, -1: 1}[mode])
    if i is None:
        raise _Err(NA)
    return float(i + 1)


def _index(arr, r, c=None):
    if not isinstance(arr, Rng):
        return arr
    r = int(_num(r)) if r is not None else 0
    c = int(_num(c)) if c is not None else 0
    rows, cols = arr.shape
    if rows == 1 and c == 0:
        r, c = 1, r
    try:
        if c == 0:
            return arr.rows[r - 1][0] if cols == 1 else Rng([arr.rows[r - 1]])
        return arr.rows[r - 1][c - 1]
    except IndexError:
        raise _Err(REF)


def _rank(x, arr, order=0):
    vals = _nums([arr])
    x = _num(x)
    if x not in vals:
        raise _Err(NA)
    order = 0 if order is None else _num(order)
    return float(1 + sum(1 for v in vals if (v > x if not order else v < x)))


def _datedif(a, b, unit):
    d1, d2, u = from_serial(_num(a)), from_serial(_num(b)), _str(unit).upper()
    if d2 < d1:
        raise _Err(NUM)
    if u == "D":
        return float((d2 - d1).days)
    months = (d2.year - d1.year) * 12 + d2.month - d1.month - (d2.day < d1.day)
    if u == "M":
        return float(months)
    if u == "Y":
        return float(months // 12)
    raise Unsupported("DATEDIF unit")


def _datevalue(s):
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        return float(s)
    d = _parse_date(_str(s))
    if d is None:
        raise _Err(VALUE)
    return float(to_serial(dt.datetime(d.year, d.month, d.day)))


def _text_before_after(after):
    def f(s, delim, inst=None, *rest):
        s, d = _str(s), _str(delim)
        n = 1 if inst is None else int(_num(inst))
        parts = s.split(d)
        if len(parts) - 1 < abs(n):
            if len(rest) >= 3 and rest[2] is not None:
                return rest[2]
            raise _Err(NA)
        if n > 0:
            return d.join(parts[n:]) if after else d.join(parts[:n])
        return d.join(parts[n:]) if after else d.join(parts[:len(parts) + n])
    return f


def _sumproduct(*arrs):
    fls = [a.flat() if isinstance(a, Rng) else [a] for a in arrs]
    n = len(fls[0])
    tot = 0.0
    for i in range(n):
        p = 1.0
        for fl in fls:
            v = fl[i]
            if isinstance(v, XLError):
                raise _Err(v)
            p *= float(v) if isinstance(v, (int, float)) else 0.0
        tot += p
    return tot


def _countif(rng, c):
    pred = _criteria(c)
    return float(sum(1 for v in (rng.flat() if isinstance(rng, Rng) else [rng]) if pred(v)))


def _today():
    return float(to_serial(dt.datetime.combine(dt.date.today(), dt.time())))


def _find(case):
    def f(sub, s, start=None):
        s, sub = _str(s), _str(sub)
        st = 1 if start is None else int(_num(start))
        i = (s.find(sub, st - 1) if case else s.lower().find(sub.lower(), st - 1))
        if i < 0:
            raise _Err(VALUE)
        return float(i + 1)
    return f


def _regex(kind):
    def f(s, pat, *rest):
        s, pat = _str(s), _str(pat)
        if kind == "test":
            return re.search(pat, s) is not None
        if kind == "extract":
            m = re.search(pat, s)
            if not m:
                raise _Err(NA)
            return m.group(0)
        return re.sub(pat, re.sub(r"\$(\d)", r"\\\1", _str(rest[0]) if rest else ""), s)
    return f


def _sub(s, old, new, inst=None):
    s, old, new = _str(s), _str(old), _str(new)
    if inst is None:
        return s.replace(old, new)
    n, i = int(_num(inst)), -1
    for _ in range(n):
        i = s.find(old, i + 1)
        if i < 0:
            return s
    return s[:i] + new + s[i + len(old):]


def _textjoin(d, ignore, *xs):
    d, ig = _str(d), _bool(ignore)
    return d.join(_str(v) for v in _flat(xs) if not (ig and v in (None, "")))


def _value(s):
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        return float(s)
    t = _str(s).strip().replace(",", "").replace("¥", "").replace("￥", "").replace("$", "")
    if t.endswith("%"):
        return float(t[:-1]) / 100
    try:
        return float(t)
    except ValueError:
        return _datevalue(s)


def _eomonth(s, m):
    d = from_serial(_num(s))
    y, mo = divmod(d.month - 1 + int(_num(m)) + 1, 12)
    first_next = dt.datetime(d.year + y, mo + 1, 1)
    return float(to_serial(first_next - dt.timedelta(days=1)))


def _edate(s, m):
    d = from_serial(_num(s))
    y, mo = divmod(d.month - 1 + int(_num(m)), 12)
    import calendar
    day = min(d.day, calendar.monthrange(d.year + y, mo + 1)[1])
    return float(to_serial(dt.datetime(d.year + y, mo + 1, day)))


def _guard(f):
    def g(*a):
        _raise_err([x for x in a if not isinstance(x, Rng)])
        return f(*a)
    return g


_FUNCS = {
    "SUM": lambda *a: (_raise_err(a), sum(_nums(a)))[1],
    "AVERAGE": lambda *a: (lambda v: sum(v) / len(v) if v else (_ for _ in ()).throw(_Err(DIV0)))(_nums(a)),
    "MIN": lambda *a: min(_nums(a), default=0.0), "MAX": lambda *a: max(_nums(a), default=0.0),
    "COUNT": lambda *a: float(len(_nums(a))), "COUNTA": lambda *a: float(sum(1 for v in _flat(a) if v not in (None, ""))),
    "COUNTBLANK": lambda *a: float(sum(1 for v in _flat(a) if v in (None, ""))),
    "SUMIF": lambda r, c, s=None: _sumifs(s if s is not None else r, r, c),
    "SUMIFS": _sumifs, "COUNTIF": _countif,
    "COUNTIFS": lambda *p: float(sum(_mask(_ifs_pairs(list(p)), len(p[0].flat() if isinstance(p[0], Rng) else [p[0]])))),
    "AVERAGEIF": lambda r, c, s=None: (lambda v: sum(v) / len(v) if v else (_ for _ in ()).throw(_Err(DIV0)))(_sel(s if s is not None else r, [r, c])),
    "AVERAGEIFS": lambda s, *p: (lambda v: sum(v) / len(v) if v else (_ for _ in ()).throw(_Err(DIV0)))(_sel(s, p)),
    "MAXIFS": lambda s, *p: max(_sel(s, p), default=0.0), "MINIFS": lambda s, *p: min(_sel(s, p), default=0.0),
    "SUMPRODUCT": _sumproduct,
    "ROUND": _guard(lambda x, n=0: _round(_num(x), int(_num(n)), "half")),
    "ROUNDUP": _guard(lambda x, n=0: _round(_num(x), int(_num(n)), "up")),
    "ROUNDDOWN": _guard(lambda x, n=0: _round(_num(x), int(_num(n)), "down")),
    "INT": _guard(lambda x: float(math.floor(_num(x)))), "ABS": _guard(lambda x: abs(_num(x))),
    "MOD": _guard(lambda x, y: _num(x) % _num(y) if _num(y) else (_ for _ in ()).throw(_Err(DIV0))),
    "SQRT": _guard(lambda x: math.sqrt(_num(x))), "POWER": _guard(lambda x, y: _num(x) ** _num(y)),
    "AND": lambda *a: all(_bool(v) for v in _flat(a) if v is not None),
    "OR": lambda *a: any(_bool(v) for v in _flat(a) if v is not None),
    "NOT": _guard(lambda x: not _bool(x)), "TRUE": lambda: True, "FALSE": lambda: False,
    "ISNUMBER": lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
    "ISTEXT": lambda v: isinstance(v, str), "ISBLANK": lambda v: v is None,
    "ISERROR": lambda v: isinstance(v, XLError), "ISNA": lambda v: v == NA,
    "ISERR": lambda v: isinstance(v, XLError) and v != NA,
    "VLOOKUP": _guard(_vlookup), "XLOOKUP": _xlookup, "MATCH": _match,
    "XMATCH": lambda k, a, m=0, s=1: float(1 + (lambda i: i if i is not None else (_ for _ in ()).throw(_Err(NA)))(
        _lookup_pos(k, a.flat(), int(_num(m)) if m is not None else 0, int(_num(s)) if s is not None else 1))),
    "INDEX": _index,
    "LEFT": _guard(lambda s, n=1: _str(s)[:int(_num(n))]),
    "RIGHT": _guard(lambda s, n=1: _str(s)[-int(_num(n)):] if int(_num(n)) else ""),
    "MID": _guard(lambda s, i, n: _str(s)[int(_num(i)) - 1:int(_num(i)) - 1 + int(_num(n))]),
    "LEN": _guard(lambda s: float(len(_str(s)))), "TRIM": _guard(lambda s: re.sub(r" +", " ", _str(s)).strip()),
    "CLEAN": _guard(lambda s: re.sub(r"[\x00-\x1f]", "", _str(s))),
    "UPPER": _guard(lambda s: _str(s).upper()), "LOWER": _guard(lambda s: _str(s).lower()),
    "PROPER": _guard(lambda s: _str(s).title()),
    "FIND": _guard(_find(True)), "SEARCH": _guard(_find(False)),
    "SUBSTITUTE": _guard(_sub), "REPLACE": _guard(lambda s, i, n, new: _str(s)[:int(_num(i)) - 1] + _str(new) + _str(s)[int(_num(i)) - 1 + int(_num(n)):]),
    "CONCAT": lambda *a: "".join(_str(v) for v in _flat(a)), "CONCATENATE": lambda *a: "".join(_str(v) for v in _flat(a)),
    "TEXTJOIN": _textjoin, "TEXT": _guard(lambda v, f: _text_fmt(v, _str(f))),
    "VALUE": _guard(_value), "NUMBERVALUE": _guard(_value), "N": lambda v: _num(v) if isinstance(v, (int, float)) else 0.0,
    "T": lambda v: v if isinstance(v, str) else "",
    "TEXTBEFORE": _guard(_text_before_after(False)), "TEXTAFTER": _guard(_text_before_after(True)),
    "REGEXTEST": _guard(_regex("test")), "REGEXEXTRACT": _guard(_regex("extract")), "REGEXREPLACE": _guard(_regex("replace")),
    "DATE": _guard(lambda y, m, d: _edate(to_serial(dt.datetime(int(_num(y)), 1, 1)), _num(m) - 1) + _num(d) - 1),
    "DATEVALUE": _guard(_datevalue), "TODAY": _today,
    "YEAR": _guard(lambda s: float(from_serial(_num(s)).year)), "MONTH": _guard(lambda s: float(from_serial(_num(s)).month)),
    "DAY": _guard(lambda s: float(from_serial(_num(s)).day)),
    "WEEKDAY": _guard(lambda s, t=1: float({1: (from_serial(_num(s)).weekday() + 1) % 7 + 1, 2: from_serial(_num(s)).weekday() + 1,
                                            3: from_serial(_num(s)).weekday()}[int(_num(t)) if t is not None else 1])),
    "WEEKNUM": _guard(lambda s, t=1: float(_weeknum(_num(s), int(_num(t)) if t is not None else 1))),
    "ISOWEEKNUM": _guard(lambda s: float(from_serial(_num(s)).isocalendar()[1])),
    "EOMONTH": _guard(_eomonth), "EDATE": _guard(_edate), "DATEDIF": _guard(_datedif),
    "DAYS": _guard(lambda e, s: float(math.floor(_num(e)) - math.floor(_num(s)))),
    "RANK": _rank, "RANK.EQ": _rank,
    "LARGE": lambda a, k: sorted(_nums([a]), reverse=True)[int(_num(k)) - 1],
    "SMALL": lambda a, k: sorted(_nums([a]))[int(_num(k)) - 1],
    "PERCENTRANK": lambda a, x, s=None: _percentrank(_nums([a]), _num(x), 3 if s is None else int(_num(s))),
    "PERCENTRANK.INC": lambda a, x, s=None: _percentrank(_nums([a]), _num(x), 3 if s is None else int(_num(s))),
    "PERCENTILE": lambda a, k: _percentile(_nums([a]), _num(k)),
    "PERCENTILE.INC": lambda a, k: _percentile(_nums([a]), _num(k)),
    "MEDIAN": lambda *a: _percentile(_nums(a), 0.5),
    "CHOOSE": lambda i, *xs: xs[int(_num(i)) - 1],
    "NA": lambda: (_ for _ in ()).throw(_Err(NA)),
}
