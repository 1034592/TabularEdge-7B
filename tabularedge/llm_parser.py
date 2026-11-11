"""第二层：模型 → Action Plan（llama.cpp server，OpenAI 兼容接口，流式）。

流程：表格画像 + 指令 → 流式解码（summary 字段一完整就回调给 UI，做到“1 秒出确认摘要”）
     → normalize + sanitize（schema 驱动的宽容清洗：字段别名、类型转换、颜色名、丢弃未知非核心字段）
     → 完整 schema + 语义校验 → 失败则把具体错误回喂模型自修复（≤2 轮，保留错误最少的一版）
     → 仍失败：只错在外观类步骤就剔除它们照常给出计划；否则返回草稿 + 逐条中文说明（不执行），
       样本记入 hard_cases.jsonl。不再退回规则解析器。

style="sft"     微调后的模型：短系统提示（tabularedge.prompting），输出 summary/lang/steps/ambiguities
style="fewshot" 未微调基座：长系统提示（op 速查）+ 1 个示例

注意：llama.cpp 的 json_schema→GBNF 不支持 if/then，所以 grammar=True 时只约束“骨架”，
参数级约束由解码后的完整校验负责。
"""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path

import requests

from .prompting import complete, messages
from .sanitize import sanitize
from .protocol import OPS, TableProfile, validate

ENDPOINT = os.environ.get("TABULAREDGE_LLM", "http://127.0.0.1:8080")
HARD_CASES = Path(os.environ.get("TABULAREDGE_HARDCASES", "outputs/hard_cases.jsonl"))

SKELETON = {
    "type": "object", "required": ["summary", "lang", "steps"],
    "properties": {
        "summary": {"type": "string"}, "lang": {"enum": ["zh", "en", "mixed"]},
        "steps": {"type": "array", "minItems": 1, "items": {
            "type": "object", "required": ["id", "op", "args"],
            "properties": {"id": {"type": "string"}, "op": {"enum": sorted(OPS)}, "sheet": {"type": "string"},
                           "args": {"type": "object"}, "depends_on": {"type": "array", "items": {"type": "string"}}}}},
        "ambiguities": {"type": "array", "items": {"type": "object",
            "properties": {"step_id": {"type": "string"}, "question": {"type": "string"},
                           "alternatives": {"type": "array", "items": {"type": "object"}}}, "required": ["step_id", "question"]}},
    },
}

SYSTEM = """你是 TabularEdge 表格指令编译器。把用户的中文/英文需求编译成 Action Plan JSON（version 0.1），只输出 JSON。
规则：
1. 只能使用给定 op；列名必须来自 columns 列表（用户说英文时映射到对应中文表头）。
2. 不确定的解读写进 ambiguities，不要猜；清洗步骤先于汇总，汇总 depends_on 清洗步骤。
3. 永远不要覆盖源文件：末尾用 save 另存。
op 速查：regex_replace{col,pattern,repl} trim{cols} delete_rows{where} dedupe{cols?} fill_empty{col,value}
sort{by:[{col,desc}]} add_column{name,formula: "=[单价]*[数量]"} conditional_format{where,apply_to:cell|row,style{font_color|fill_color:"RRGGBB"}}
number_format{col,format} group_aggregate{by[],aggs[{col,fn:sum|count|avg|min|max}],output_sheet,mode:formula|value}
lookup_join{right,on,right_on?,fetch[]} save{path} fs_mkdir{path_template,expand} fs_backup{path} fs_new_workbook{path,headers}
fs_merge_workbooks{glob,output} split_to_files{by,path_template:"out/{value}.xlsx"}
where 条件：{col,cmp:==|!=|<|<=|>|>=|contains|regex|is_empty|...,value} 或 {all:[...]} / {any:[...]}"""

# ---------------------------------------------------------------- 输出规整（确定性修复，不改语义）
_FN_ALIAS = {"REGEX_EXTRACT": "REGEXEXTRACT", "REGEX_REPLACE": "REGEXREPLACE", "REGEX_TEST": "REGEXTEST", "REGEXMATCH": "REGEXTEST",
             "PERCENT_RANK": "PERCENTRANK.INC", "PERCENTRANKINC": "PERCENTRANK.INC", "ISO_WEEKNUM": "ISOWEEKNUM",
             "WEEKNUMISO": "ISOWEEKNUM", "DATE_VALUE": "DATEVALUE", "TEXT_BEFORE": "TEXTBEFORE", "TEXT_AFTER": "TEXTAFTER"}


def _fix_formula(f: str) -> str:
    """Excel 语法常见笔误：== → =，'文本' → "文本"（不动 'Sheet'! 这种表名引用），函数名别名。"""
    parts = re.split(r'("(?:[^"]|"")*")', f)
    for i in range(0, len(parts), 2):
        p = parts[i]
        p = re.sub(r"(?<![<>!=])==(?!=)", "=", p)
        p = re.sub(r"'([^'!]*)'(?!!)", lambda m: '"' + m.group(1).replace('"', '""') + '"', p)
        p = re.sub(r"(?<![A-Za-z0-9_.])([A-Za-z_][A-Za-z0-9_.]*)(?=\()",
                   lambda m: _FN_ALIAS.get(m.group(1).upper(), m.group(1)), p)
        parts[i] = p
    return "".join(parts)


def _repair_json(raw: str) -> str:
    """修复公式字符串里没转义的双引号：把 "formula":"…" 按其结束边界 "} 重新转义。"""
    def fix(m):
        body = m.group(2)
        body = re.sub(r'(?<!\\)"', r'\\"', body)
        return m.group(1) + body + m.group(3)
    return re.sub(r'("formula"\s*:\s*")(.*?)("\s*})', fix, raw)


def normalize(obj: dict) -> dict:
    """去掉 null / 不认识的枚举值（让默认值生效），规整公式写法。"""
    for st in obj.get("steps") or []:
        a = st.get("args")
        if not isinstance(a, dict):
            continue
        for k in [k for k, v in a.items() if v is None]:
            del a[k]
        if a.get("mode") not in (None, "formula", "value"):
            del a["mode"]
        if isinstance(a.get("formula"), str):
            a["formula"] = _fix_formula(a["formula"])
            if not a["formula"].startswith("="):
                a["formula"] = "=" + a["formula"]
        if st.get("op") == "conditional_format" and a.get("apply_to") not in (None, "cell", "row"):
            del a["apply_to"]
    return obj


_SUMMARY_RX = re.compile(r'"summary"\s*:\s*"((?:[^"\\]|\\.)*)"')


_DEMO_TEXT = "把带括号的备注删了，金额负数的标红，按部门汇总新表"
_DEMO_PROFILE = TableProfile("流水", ["日期", "部门", "金额", "备注"], {"日期": "text", "部门": "text", "金额": "num", "备注": "text"}, 5)
# 一个静态示例计划（含一条歧义追问），只给未微调的基座做 few-shot 用；运行时管线里没有任何规则解析
_DEMO_PLAN = '{"lang":"zh","summary":"把带括号的备注删了，金额负数的标红，按部门汇总新表","steps":[{"op":"regex_replace","args":{"col":"备注","pattern":"[\\\\(（][^\\\\)）]*[\\\\)）]","repl":""},"id":"s1","sheet":"流水"},{"op":"conditional_format","args":{"where":{"col":"金额","cmp":"<","value":0},"apply_to":"cell","style":{"font_color":"FF0000"}},"id":"s2","sheet":"流水"},{"op":"group_aggregate","args":{"by":["部门"],"aggs":[{"col":"金额","fn":"sum"}],"output_sheet":"按部门汇总","mode":"formula"},"id":"s3","sheet":"流水","depends_on":["s1"]},{"id":"s4","op":"save","args":{"path":"ledger_处理后.xlsx"}}],"ambiguities":[{"question":"“把带括号的备注删了”是指删除【备注】里括号及括号中的文字，还是删除含括号的整行？","alternatives":[{"op":"delete_rows","args":{"where":{"col":"备注","cmp":"regex","value":"[\\\\(（][^\\\\)）]*[\\\\)）]"}},"id":"s1","sheet":"流水"}],"step_id":"s1"}]}'


def _fewshot_msgs(profile: TableProfile, file: str, text: str):
    return [{"role": "system", "content": SYSTEM},
            {"role": "user", "content": f"文件：ledger.xlsx\n{_DEMO_PROFILE.to_prompt()}\n指令：{_DEMO_TEXT}"},
            {"role": "assistant", "content": _DEMO_PLAN},
            {"role": "user", "content": f"文件：{file}\n{profile.to_prompt()}\n指令：{text}"}]


def _stream(msgs, on_summary=None, grammar=False, max_tokens=1200, url=ENDPOINT, on_delta=None, stop=None, temperature=0.0):
    body = {"messages": msgs, "temperature": temperature, "seed": 7, "max_tokens": max_tokens, "stream": True, "cache_prompt": True}
    if grammar:
        body["response_format"] = {"type": "json_schema", "json_schema": {"name": "plan", "schema": SKELETON}}
    t0 = time.time()
    buf, t_first, t_summary, said = "", None, None, False
    finish = None
    with requests.post(url.rstrip("/") + "/v1/chat/completions", json=body, stream=True, timeout=600) as r:
        r.raise_for_status()
        # 必须按字节切行再 UTF-8 解码：iter_lines(decode_unicode=True) 在无 charset 时按 ISO-8859-1 解码，
        # 中文 UTF-8 里的 0x85 字节会变成 NEL（U+0085），被 splitlines 当换行，把一条 SSE 切成两半
        for raw_line in r.iter_lines(delimiter=b"\n"):
            line = raw_line.decode("utf-8").rstrip("\r")
            if not line or not line.startswith("data: ") or line == "data: [DONE]":
                continue
            ch = json.loads(line[6:])["choices"][0]
            if ch.get("finish_reason"):
                finish = ch["finish_reason"]
            delta = ch.get("delta", {}).get("content") or ""
            if delta and t_first is None:
                t_first = time.time() - t0
            buf += delta
            if on_delta and delta:
                on_delta(delta)
            if stop is not None and stop.is_set():
                break
            if not said and (m := _SUMMARY_RX.search(buf)):
                said, t_summary = True, time.time() - t0
                if on_summary:
                    on_summary(json.loads(f'"{m.group(1)}"'))
    return buf, {"first_token_s": t_first, "summary_s": t_summary, "total_s": time.time() - t0, "truncated": finish == "length"}


def prefill(profile: TableProfile, file: str, extra_profile: str = "", url: str = ENDPOINT, timeout: float = 120):
    """用户选中表格后、还在打字时调用：把“系统提示 + 文件名 + 表格画像”先算进 llama-server 的 KV 缓存（cache_prompt）。
    真正解析时提示词前缀完全相同，只需再处理指令那几十个 token，首字从 1.5~2.5s 降到零点几秒。"""
    ptxt = profile.to_prompt() + ("\n" + extra_profile if extra_profile else "")
    body = {"messages": messages(file, ptxt, ""), "max_tokens": 1, "temperature": 0.0, "cache_prompt": True}
    t0 = time.time()
    requests.post(url.rstrip("/") + "/v1/chat/completions", json=body, timeout=timeout).raise_for_status()
    return time.time() - t0


COSMETIC_OPS = {"conditional_format", "number_format"}  # 只影响外观、不改数据：修不好可以剔除后照常执行
_OP_ZH = {"lookup_column": "跨表查找", "lookup_join": "跨表合并", "add_column": "新增列", "conditional_format": "条件格式",
          "number_format": "数字格式", "filter_to_sheet": "筛选", "group_aggregate": "分组汇总", "delete_rows": "删除行",
          "dedupe": "去重", "sort": "排序", "save": "保存", "regex_replace": "正则替换", "trim": "去空格", "fill_empty": "填充空值",
          "unique_to_sheet": "不重复清单", "split_to_files": "拆分文件"}


def _err_step(err: str, plan: dict | None):
    """错误文本 → 出错步骤的 id（schema 错误给的是下标 steps/3/…，语义错误直接写 s4）。"""
    m = re.match(r"schema: steps/(\d+)", err)
    if m and plan and int(m.group(1)) < len(plan.get("steps", [])):
        return plan["steps"][int(m.group(1))].get("id")
    m = re.match(r"semantic: (s\d+)\b", err)
    return m.group(1) if m else None


def explain(errs: list[str], plan: dict | None) -> list[str]:
    """校验错误 → 给用户看的具体说明（哪一步、什么问题、怎么改）。"""
    out = []
    steps = {s.get("id"): s for s in (plan or {}).get("steps", []) if isinstance(s, dict)}
    for e in errs:
        sid = _err_step(e, plan)
        st = steps.get(sid) or {}
        head = f"第 {sid} 步（{_OP_ZH.get(st.get('op'), st.get('op') or '?')}）" if sid else "计划"
        if e.startswith("invalid json"):
            out.append("模型输出不是完整的 JSON（多半是指令太长、计划被截断）。建议把指令拆成两三句，分步执行。")
        elif m := re.search(r"unknown column '([^']+)' in sheet '([^']+)' \((\w+)\); available: \[(.*?)\]", e):
            avail = "、".join(x.strip(" '") for x in m.group(4).split(",")[:12])
            out.append(f"{head}：表「{m.group(2)}」里没有「{m.group(1)}」这一列（现有列：{avail}…）。可以直接说列字母，如“F列”。")
        elif m := re.search(r"unknown column '([^']+)'; available: \[(.*)\]", e):
            avail = "、".join(x.strip(" '") for x in m.group(2).split(",")[:12])
            out.append(f"{head}：表里没有「{m.group(1)}」这一列（现有列：{avail}…）。请用表里的列名重说。")
        elif m := re.search(r"unknown sheet '([^']+)' in \w+; available: \[(.*)\]", e):
            avail = "、".join(x.strip(" '") for x in m.group(2).split(","))
            out.append(f"{head}：找不到表「{m.group(1)}」（可用的表：{avail}）。请在指令里写出参考文件的文件名。")
        elif m := re.search(r"'([^']+)' is a required property", e):
            out.append(f"{head}：缺少必填项 {m.group(1)}。请在指令里说清楚这一步要用哪一列 / 哪张表。")
        elif "points at one fixed cell" in e:
            out.append(f"{head}：公式固定指向一个单元格，每一行都会得到同一个值。要引用同一行的某列，请写 [列名]。")
        elif "is just a column letter" in e:
            out.append(f"{head}：新列的名字只是一个列字母，看起来没理解要做什么。请直接说新列叫什么、填什么。")
        elif "truncated" in e:
            out.append("模型输出超出长度上限被截断。请把指令拆成两三句，分步执行。")
        else:
            out.append(f"{head}：{e.split(': ', 1)[-1][:160]}")
    return list(dict.fromkeys(out))


def _salvage(plan: dict, errs: list[str], profile):
    """只有“外观类”步骤（条件格式 / 数字格式）修不好时：剔除这些步骤，核心数据操作照常给出。"""
    bad = {_err_step(e, plan) for e in errs}
    if None in bad:
        return None, []
    ops = {s["id"]: s.get("op") for s in plan.get("steps", []) if isinstance(s, dict) and "id" in s}
    if not bad or any(ops.get(b) not in COSMETIC_OPS for b in bad):
        return None, []
    p = dict(plan)
    dropped = [s for s in plan["steps"] if s.get("id") in bad]
    p["steps"] = [dict(s, depends_on=[d for d in s.get("depends_on", []) if d not in bad]) if s.get("depends_on") else s
                  for s in plan["steps"] if s.get("id") not in bad]
    p["ambiguities"] = [a for a in plan.get("ambiguities") or [] if a.get("step_id") not in bad]
    if not p["ambiguities"]:
        p.pop("ambiguities")
    if not p["steps"] or validate(p, profile):
        return None, []
    return p, dropped


def _wrap_regexextract(f: str) -> str:
    """REGEXEXTRACT 找不到时是 #N/A；“提取……没有的留空”是普遍预期 → 没包 IFERROR 的一律包上。"""
    from .compat import _find_call
    if "REGEXEXTRACT" not in f.upper():
        return f
    out, pos = f, 0
    while (hit := _find_call(out, "REGEXEXTRACT", pos)):
        st, en, _ = hit
        before = out[:st].rstrip().upper()
        if before.endswith("IFERROR(") or before.endswith("IFNA("):
            pos = en
            continue
        call = out[st:en]
        out = out[:st] + f'IFERROR({call},"")' + out[en:]
        pos = st + len(call) + 13
    return out


def _post_normalize(plan: dict) -> list[str]:
    """与指令措辞无关的结构性规整：提取类公式（REGEXEXTRACT）找不到时是 #N/A → 一律包 IFERROR(…,"")。"""
    out = []
    for st in plan.get("steps", []):
        a = st.get("args") or {}
        if st.get("op") == "add_column" and isinstance(a.get("formula"), str):
            g = _wrap_regexextract(a["formula"])
            if g != a["formula"]:
                out.append(f"{st.get('id')}.formula: REGEXEXTRACT wrapped in IFERROR(…,\"\")")
                a["formula"] = g
    return out


def _known_sheets(ptxt: str) -> list[str]:
    return re.findall(r"\[sheet=(.+?) rows=\d+\]", ptxt)


def _sheet_fields(plan):
    for st in plan.get("steps", []):
        a = st.get("args") or {}
        for k in ("from_sheet", "right"):
            if isinstance(a.get(k), str):
                yield st, a, k


def _fix_sheet_refs(plan: dict, ptxt: str) -> list[str]:
    """跨表引用写成了文件名（带 .xlsx / 带序号前缀 / 全角书名号）→ 唯一对得上的表名就改过来。"""
    known, out = _known_sheets(ptxt), []
    for st, a, k in _sheet_fields(plan):
        v = a[k]
        if v in known:
            continue
        core_ = re.sub(r"\.xlsx?$", "", v.strip().strip("《》“”\"'"), flags=re.I)
        cands = [s for s in known if s == core_ or s == re.sub(r"^\d+[_\-. ]*", "", core_) or (len(s) >= 4 and (s in core_ or core_ in s))]
        if len(cands) == 1:
            a[k] = cands[0]
            out.append(f"{st.get('id')}.{k}: {v!r} → {cands[0]!r}")
    return out


def _check_sheet_refs(plan: dict, ptxt: str) -> list[str]:
    known = _known_sheets(ptxt)
    return [f"semantic: {st.get('id')} unknown sheet {a[k]!r} in {k}; available: {known}"
            for st, a, k in _sheet_fields(plan) if a[k] not in known]


def _ref_headers(ptxt: str) -> dict[str, list[str]]:
    """提示词里每张表的列名：[sheet=X rows=N] columns: a:text, b:num …"""
    out = {}
    for m in re.finditer(r"\[sheet=(.+?) rows=\d+\] columns: (.*)", ptxt):
        out[m.group(1)] = [c.rsplit(":", 1)[0] for c in m.group(2).split(", ") if c]
    return out


def _col_ok(ref, heads) -> bool:
    return isinstance(ref, str) and (ref in heads or bool(re.fullmatch(r"[A-Z]{1,3}", ref)))


_REF_COLS = {"lookup_column": ("from_sheet", ("from_key", "return")), "lookup_join": ("right", ("right_on",))}


def _check_ref_cols(plan: dict, ptxt: str) -> list[str]:
    heads, errs = _ref_headers(ptxt), []
    for st in plan.get("steps", []):
        spec = _REF_COLS.get(st.get("op"))
        a = st.get("args") or {}
        if not spec or a.get(spec[0]) not in heads:
            continue
        for k in spec[1]:
            if k in a and not _col_ok(a[k], heads[a[spec[0]]]):
                errs.append(f"semantic: {st.get('id')} unknown column {a[k]!r} in sheet {a[spec[0]]!r} ({k}); "
                            f"available: {heads[a[spec[0]]]} or a column letter like 'F'")
    return errs


def _check_sense(plan: dict) -> list[str]:
    """能被 schema 接受、但执行出来必然不对的写法：逐行公式钉死在一个单元格（=H2，每行都一样）；新列名就是个列字母。"""
    errs = []
    for st in plan.get("steps", []):
        a = st.get("args") or {}
        if st.get("op") == "add_column" and isinstance(a.get("formula"), str):
            body = re.sub(r'"(?:[^"]|"")*"', "", a["formula"])
            body = re.sub(r"\[[^\]]+\]", "", body)
            if re.search(r"(?<![A-Za-z$!:])[A-Z]{1,3}\d+(?![\d:(])", body) and "{r}" not in body:
                errs.append(f"semantic: {st.get('id')} formula {a['formula']!r} points at one fixed cell, so every row gets the same "
                            "value; use [column name] to refer to the same row")
        if st.get("op") in ("add_column", "lookup_column") and re.fullmatch(r"[A-Z]{1,3}", str(a.get("name", ""))):
            errs.append(f"semantic: {st.get('id')} new column name {a['name']!r} is just a column letter; give it a meaningful name")
    return errs


def budget(prompt_chars: int, ctx: int = 4096) -> int:
    """生成长度上限：上下文总长减去提示词（中文约 1.5 字/token，按保守估计），至少 600，至多 2400。"""
    return max(600, min(2400, ctx - int(prompt_chars / 1.5) - 64))


def parse(text: str, profile: TableProfile, file: str, style: str = "sft", on_summary=None,
          grammar: bool = False, max_repair: int = 2, url: str = ENDPOINT, extra_profile: str = "",
          on_delta=None, on_repair=None, stop=None, ctx: int = 4096) -> tuple[dict | None, dict]:
    """返回 (plan, info)。

    管道：模型输出 → JSON（含引号/截断修补）→ normalize → sanitize（schema 驱动的宽容清洗）→ validate
          → 仍有错：把具体错误回喂模型自我修正（最多 max_repair 轮），保留错误最少的一版
          → 仍有错：只错在外观类步骤 → 剔除这些步骤后照常给出计划（info["dropped"]）；
                    否则返回草稿 + info["failed"]=True + info["explain"]（逐条说明哪一步什么问题），界面据此提示，不执行。
    不再退回规则解析器（它会把长指令切碎成一堆“没能理解”的追问）。"""
    ptxt = profile.to_prompt() + ("\n" + extra_profile if extra_profile else "")
    msgs = messages(file, ptxt, text) if style == "sft" else _fewshot_msgs(profile, file, text)
    trace, best = [], None
    for attempt in range(max_repair + 1):
        if attempt and on_repair:
            on_repair(attempt, explain(trace[-1]["errors"], trace[-1].get("plan")))
        # 第 1 次贪心解码；修复轮稍加温度，避免模型逐字重复同一个错误
        raw, tm = _stream(msgs, on_summary if attempt == 0 else None, grammar, url=url, on_delta=on_delta, stop=stop,
                          temperature=0.0 if attempt == 0 else 0.4, max_tokens=budget(sum(len(m["content"]) for m in msgs), ctx))
        if stop is not None and stop.is_set():
            raise InterruptedError("cancelled")
        plan, fixes = None, []
        try:
            try:
                obj = json.loads(raw)
            except json.JSONDecodeError:
                obj = json.loads(_repair_json(raw))
            if not isinstance(obj, dict):
                raise json.JSONDecodeError("top level is not an object", raw, 0)
            obj, fixes = sanitize(normalize(obj))
            obj = normalize(obj)
            if not obj.get("steps") and obj.get("ambiguities"):
                # 只追问、不给步骤：这是合法回答（指令太含糊），交给界面展示问题
                plan = complete({k: v for k, v in obj.items() if k not in ("version", "inputs")}, file, profile.sheet, getattr(profile, "header_row", 1))
                return plan, {"attempts": attempt + 1, "fallback": False, "ask_only": True, "fixes": fixes, **tm}
            plan = complete({k: v for k, v in obj.items() if k not in ("version", "inputs")}, file, profile.sheet, getattr(profile, "header_row", 1))
            fixes += _post_normalize(plan)
            fixes += _fix_sheet_refs(plan, ptxt)
            errs = validate(plan, profile) or _check_sheet_refs(plan, ptxt) or _check_ref_cols(plan, ptxt) or _check_sense(plan)
        except json.JSONDecodeError as e:
            errs = [f"invalid json: {e}" + (" (output truncated at max_tokens)" if tm.get("truncated") else "")]
        rec = {"attempt": attempt, **tm, "errors": errs, "raw": raw, "plan": plan, "fixes": fixes}
        trace.append(rec)
        if not errs:
            return plan, {"attempts": attempt + 1, "fallback": False, "fixes": fixes, **tm}
        if plan is not None and (best is None or best["plan"] is None or len(errs) < len(best["errors"])):
            best = rec
        elif best is None:
            best = rec
        if tm.get("truncated"):
            break  # 截断不是模型“写错了”，再修也一样长；直接给出拆分建议
        msgs = msgs + [{"role": "assistant", "content": raw},
                       {"role": "user", "content": "校验失败，请只修正以下问题后重新输出完整 JSON：\n- " + "\n- ".join(errs[:8])}]
    HARD_CASES.parent.mkdir(parents=True, exist_ok=True)
    with HARD_CASES.open("a", encoding="utf-8") as f:
        f.write(json.dumps({"text": text, "profile": ptxt, "trace": [{k: v for k, v in t.items() if k != "plan"} for t in trace]},
                           ensure_ascii=False) + "\n")
    info = {"attempts": len(trace), "fallback": False, "fixes": best["fixes"],
            **{k: best[k] for k in ("first_token_s", "summary_s", "total_s", "truncated") if k in best}}
    if best["plan"] is not None:
        salvaged, dropped = _salvage(best["plan"], best["errors"], profile)
        if salvaged is not None:
            info["dropped"] = [f"{d['id']} {_OP_ZH.get(d.get('op'), d.get('op'))}" for d in dropped]
            info["dropped_explain"] = explain(best["errors"], best["plan"])
            return salvaged, info
    info.update(failed=True, errors=best["errors"], explain=explain(best["errors"], best["plan"]))
    return best["plan"], info
