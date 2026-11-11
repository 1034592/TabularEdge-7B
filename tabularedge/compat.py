"""Excel 版本兼容：把只有 Microsoft 365 / Excel 2021+ 才有的函数改写成 Excel 2016/2019 也能算的等价活公式。

  XLOOKUP 精确匹配      → IFERROR(INDEX(返回列, MATCH(键, 查找列, 0)), 缺省值)
  XLOOKUP match_mode=-1 → 找“不超过键的最大值”：INDEX(返回列, MATCH(MAXIFS(查找列, 查找列, "<="&键), 查找列, 0))（不要求排序）
  XLOOKUP match_mode=1  → 同理用 MINIFS
  XMATCH                → MATCH
改写不了的（FILTER / UNIQUE / TEXTBEFORE / REGEX* / LET ...）由执行器写入计算结果（静态值）并在日志里说明。
"""
from __future__ import annotations

import re

# Excel 2019 没有的函数（出现即判为不兼容）
NEW_FUNCS = {"XLOOKUP", "XMATCH", "FILTER", "UNIQUE", "SORT", "SORTBY", "SEQUENCE", "RANDARRAY", "LET", "LAMBDA", "TEXTBEFORE",
             "TEXTAFTER", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL", "TOROW", "CHOOSECOLS", "CHOOSEROWS", "TAKE", "DROP", "EXPAND",
             "WRAPROWS", "WRAPCOLS", "REGEXTEST", "REGEXEXTRACT", "REGEXREPLACE", "GROUPBY", "PIVOTBY", "IMAGE", "ARRAYTOTEXT",
             "VALUETOTEXT", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "ISOMITTED", "PERCENTOF", "TRIMRANGE"}
_NAME_RX = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?([A-Z][A-Z0-9.]*)(?=\()")


def _split_strings(f):
    return re.split(r'("(?:[^"]|"")*")', f)


def used_funcs(f: str) -> set[str]:
    out = set()
    for i, p in enumerate(_split_strings(f)):
        if i % 2 == 0:
            out.update(m.group(1).upper() for m in _NAME_RX.finditer(p))
    return out


def incompatible(f: str) -> set[str]:
    return used_funcs(f) & NEW_FUNCS


def _find_call(f: str, name: str, start=0):
    """在公式里找 name( ... ) 的调用（跳过字符串），返回 (起点, 终点, 参数列表) 或 None。"""
    rx = re.compile(r"(?<![A-Za-z0-9_.])(?:_xlfn\.(?:_xlws\.)?)?" + re.escape(name) + r"\(", re.I)
    pos = start
    while True:
        m = rx.search(f, pos)
        if not m:
            return None
        # 跳过位于字符串字面量里的匹配
        if f[:m.start()].count('"') % 2:
            pos = m.end()
            continue
        i, depth, args, cur, in_str = m.end(), 1, [], "", False
        while i < len(f):
            ch = f[i]
            if in_str:
                cur += ch
                if ch == '"':
                    if i + 1 < len(f) and f[i + 1] == '"':
                        cur += '"'
                        i += 1
                    else:
                        in_str = False
            elif ch == '"':
                in_str, cur = True, cur + ch
            elif ch == "(":
                depth, cur = depth + 1, cur + ch
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    args.append(cur.strip())
                    return m.start(), i + 1, args
                cur += ch
            elif ch == "," and depth == 1:
                args.append(cur.strip())
                cur = ""
            else:
                cur += ch
            i += 1
        return None


def _xlookup_2019(args):
    k, L, R = args[0], args[1], args[2]
    nf = args[3] if len(args) > 3 and args[3] != "" else None
    mode = args[4].strip() if len(args) > 4 and args[4] != "" else "0"
    search = args[5].strip() if len(args) > 5 and args[5] != "" else "1"
    if search not in ("1", ""):
        return None
    nf_expr = nf if nf is not None else "NA()"
    if mode == "0" or mode == "2":
        base = f"INDEX({R},MATCH({k},{L},0))"
        return f"IFERROR({base},{nf})" if nf is not None else base
    if mode in ("-1", "1"):
        op, fn = ("<=", "MAXIFS") if mode == "-1" else (">=", "MINIFS")
        crit = f'"{op}"&({k})'
        return f"IF(COUNTIFS({L},{crit})=0,{nf_expr},INDEX({R},MATCH({fn}({L},{L},{crit}),{L},0)))"
    return None


def _xmatch_2019(args):
    mode = args[2].strip() if len(args) > 2 and args[2] != "" else "0"
    if mode != "0" or (len(args) > 3 and args[3].strip() not in ("", "1")):
        return None
    return f"MATCH({args[0]},{args[1]},0)"


REWRITE = {"XLOOKUP": _xlookup_2019, "XMATCH": _xmatch_2019}


def to_2019(f: str) -> str:
    """尽量把公式改写成 Excel 2019 可用的等价形式；改写不了的部分原样保留（调用方再用 incompatible() 判断）。"""
    for name, fn in REWRITE.items():
        pos = 0
        for _ in range(50):
            hit = _find_call(f, name, pos)
            if not hit:
                break
            s, e, args = hit
            args = [to_2019(a) for a in args]
            new = fn(args)
            if new is None:
                pos = e
                continue
            f = f[:s] + new + f[e:]
            pos = s + len(new)
    return f


def detect_excel() -> str:
    """Windows：读取 Office 版本，返回 "2019"（无动态数组）或 "365"。非 Windows / 读不到时返回 "365"。"""
    try:
        import winreg  # type: ignore
    except ImportError:
        return "365"
    try:
        k = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration")
        ids, _ = winreg.QueryValueEx(k, "ProductReleaseIds")
    except OSError:
        return "2019"  # MSI 版 Office（2016/2019）读不到 ClickToRun → 按保守的 2019 处理
    ids = str(ids)
    if any(x in ids for x in ("O365", "365", "2021", "2024")):
        return "365"
    return "2019"
