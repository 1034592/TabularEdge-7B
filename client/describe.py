"""Action Plan 步骤 → 人话（中 / 英），用于界面上的计划预览与确认弹窗。"""
from __future__ import annotations

from tabularedge.protocol import RISKY_OPS

CMP_ZH = {"==": "等于", "!=": "不等于", "<": "小于", "<=": "小于等于", ">": "大于", ">=": "大于等于", "contains": "包含",
          "not_contains": "不包含", "regex": "匹配正则", "is_empty": "为空", "not_empty": "不为空", "startswith": "开头是",
          "endswith": "结尾是"}
CMP_EN = {"==": "=", "!=": "≠", "<": "<", "<=": "≤", ">": ">", ">=": "≥", "contains": "contains", "not_contains": "does not contain",
          "regex": "matches regex", "is_empty": "is empty", "not_empty": "is not empty", "startswith": "starts with",
          "endswith": "ends with"}
FN_ZH = {"sum": "求和", "count": "计数", "avg": "平均", "min": "最小", "max": "最大"}
COLOR_ZH = {"FF0000": "红色", "C00000": "深红", "FFFF00": "黄色", "FFEB9C": "浅黄", "FFC7CE": "浅红", "C6EFCE": "浅绿",
            "00B050": "绿色", "92D050": "浅绿", "0070C0": "蓝色", "FFC000": "橙色", "808080": "灰色"}
COLOR_EN = {"FF0000": "red", "C00000": "dark red", "FFFF00": "yellow", "FFEB9C": "light yellow", "FFC7CE": "light red",
            "C6EFCE": "light green", "00B050": "green", "92D050": "light green", "0070C0": "blue", "FFC000": "orange", "808080": "grey"}


def cond(c, zh=True) -> str:
    if not isinstance(c, dict):
        return str(c)
    if "all" in c:
        return ("且" if zh else " and ").join(cond(x, zh) for x in c["all"])
    if "any" in c:
        return ("或" if zh else " or ").join(cond(x, zh) for x in c["any"])
    m = (CMP_ZH if zh else CMP_EN).get(c.get("cmp"), c.get("cmp"))
    v = "" if c.get("cmp") in ("is_empty", "not_empty") else f" {c.get('value')!r}" if isinstance(c.get("value"), str) else f" {c.get('value')}"
    return f"「{c.get('col')}」{m}{v}" if zh else f"[{c.get('col')}] {m}{v}"


def style(s, zh=True) -> str:
    parts = []
    names = COLOR_ZH if zh else COLOR_EN
    if s.get("font_color"):
        n = names.get(s["font_color"].upper(), "#" + s["font_color"])
        parts.append(f"字体{n}" if zh else f"{n} font")
    if s.get("fill_color"):
        n = names.get(s["fill_color"].upper(), "#" + s["fill_color"])
        parts.append(f"填充{n}" if zh else f"{n} fill")
    if s.get("bold"):
        parts.append("加粗" if zh else "bold")
    return "、".join(parts) if zh else ", ".join(parts)


def step(st: dict, zh: bool = True) -> str:
    op, a = st.get("op"), st.get("args", {})
    sheet = f"［{st['sheet']}］" if st.get("sheet") and zh else (f"[{st['sheet']}] " if st.get("sheet") else "")
    if zh:
        d = {
            "regex_replace": lambda: f"列「{a.get('col')}」按正则 {a.get('pattern')!r} 替换为 {a.get('repl', '')!r}",
            "trim": lambda: f"清除 {'、'.join(a.get('cols', []))} 的多余空格",
            "delete_rows": lambda: f"⚠ 删除整行：{cond(a.get('where'), True)}",
            "dedupe": lambda: f"⚠ 去重（按 {'、'.join(a['cols']) if a.get('cols') else '整行'}，保留{'最后' if a.get('keep') == 'last' else '第一'}条，其余行删除）",
            "fill_empty": lambda: f"「{a.get('col')}」空白填为 {a.get('value')!r}",
            "sort": lambda: "排序：" + "，".join(f"{b['col']}{'降序' if b.get('desc') else '升序'}" for b in a.get("by", [])),
            "add_column": lambda: f"新增列「{a.get('name')}」= {a.get('formula')}（活公式）",
            "conditional_format": lambda: f"条件格式（改数自动变色）：{cond(a.get('where'), True)} → {style(a.get('style', {}), True)}"
                                          f"{'（整行）' if a.get('apply_to') == 'row' else ''}",
            "number_format": lambda: f"「{a.get('col')}」数字格式设为 {a.get('format')}",
            "group_aggregate": lambda: f"按 {'、'.join(a.get('by', []))} 汇总 "
                                       + "，".join(f"{g['col']}{FN_ZH.get(g['fn'], g['fn'])}" for g in a.get("aggs", []))
                                       + f" → 新表「{a.get('output_sheet')}」{'（活公式）' if a.get('mode', 'formula') == 'formula' else '（数值）'}",
            "lookup_join": lambda: f"按「{a.get('on')}」从 {a.get('right')} 匹配带回 {'、'.join(a.get('fetch', []))}",
            "lookup_column": lambda: f"新增列「{a.get('name')}」：按「{a.get('key')}」到表「{a.get('from_sheet')}」的「{a.get('from_key')}」查找，"
                                     f"取「{a.get('return')}」，查不到填 {a.get('if_not_found', '')!r}{'（写入结果值）' if a.get('mode') == 'value' else '（XLOOKUP 活公式）'}",
            "filter_to_sheet": lambda: f"筛选 {cond(a.get('where'), True)} 的行 → 新表「{a.get('output_sheet')}」",
            "unique_to_sheet": lambda: f"「{a.get('col')}」不重复清单{'（带计数）' if a.get('with_count') else ''} → 新表「{a.get('output_sheet')}」",
            "save": lambda: f"保存为 {a.get('path')}",
            "fs_mkdir": lambda: f"建文件夹 {a.get('path_template')}",
            "fs_backup": lambda: f"备份 {a.get('path')}",
            "fs_new_workbook": lambda: f"新建工作簿 {a.get('path')}",
            "fs_merge_workbooks": lambda: f"合并 {a.get('glob')} → {a.get('output')}",
            "split_to_files": lambda: f"按「{a.get('by')}」拆分成多个文件 → {a.get('path_template')}",
        }
    else:
        d = {
            "regex_replace": lambda: f"In [{a.get('col')}], replace regex {a.get('pattern')!r} with {a.get('repl', '')!r}",
            "trim": lambda: f"Trim extra spaces in {', '.join(a.get('cols', []))}",
            "delete_rows": lambda: f"⚠ Delete rows where {cond(a.get('where'), False)}",
            "dedupe": lambda: f"⚠ Remove duplicates (by {', '.join(a['cols']) if a.get('cols') else 'whole row'}, keep {a.get('keep', 'first')})",
            "fill_empty": lambda: f"Fill blanks in [{a.get('col')}] with {a.get('value')!r}",
            "sort": lambda: "Sort by " + ", ".join(f"{b['col']} {'desc' if b.get('desc') else 'asc'}" for b in a.get("by", [])),
            "add_column": lambda: f"Add column [{a.get('name')}] = {a.get('formula')} (live formula)",
            "conditional_format": lambda: f"Conditional format: {cond(a.get('where'), False)} → {style(a.get('style', {}), False)}"
                                          f"{' (whole row)' if a.get('apply_to') == 'row' else ''}",
            "number_format": lambda: f"Number format of [{a.get('col')}] → {a.get('format')}",
            "group_aggregate": lambda: f"Group by {', '.join(a.get('by', []))}: "
                                       + ", ".join(f"{g['fn']}({g['col']})" for g in a.get("aggs", []))
                                       + f" → sheet '{a.get('output_sheet')}'",
            "lookup_join": lambda: f"Join on [{a.get('on')}] from {a.get('right')}, fetch {', '.join(a.get('fetch', []))}",
            "lookup_column": lambda: f"Add column [{a.get('name')}]: look up [{a.get('key')}] in '{a.get('from_sheet')}'.{a.get('from_key')}, "
                                     f"return {a.get('return')}, default {a.get('if_not_found', '')!r}{' (values)' if a.get('mode') == 'value' else ' (live XLOOKUP)'}",
            "filter_to_sheet": lambda: f"Filter rows where {cond(a.get('where'), False)} → sheet '{a.get('output_sheet')}'",
            "unique_to_sheet": lambda: f"Distinct values of [{a.get('col')}]{' with counts' if a.get('with_count') else ''} → sheet '{a.get('output_sheet')}'",
            "save": lambda: f"Save as {a.get('path')}",
            "fs_mkdir": lambda: f"Create folder {a.get('path_template')}",
            "fs_backup": lambda: f"Back up {a.get('path')}",
            "fs_new_workbook": lambda: f"New workbook {a.get('path')}",
            "fs_merge_workbooks": lambda: f"Merge {a.get('glob')} → {a.get('output')}",
            "split_to_files": lambda: f"Split into one file per [{a.get('by')}] → {a.get('path_template')}",
        }
    try:
        return sheet + d[op]()
    except Exception:  # noqa: BLE001
        return f"{op} {a}"


def is_risky(st: dict, src_file: str) -> bool:
    return st.get("op") in RISKY_OPS or (st.get("op") == "save" and st.get("args", {}).get("path") == src_file)
