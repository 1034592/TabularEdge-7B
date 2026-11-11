"""宽容清洗 + 解析管道（不再退回规则解析器）+ 多参考表提示词精简 的回归测试。模型输出用假流替代。"""
import json
import sys
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from openpyxl import Workbook  # noqa: E402

import tabularedge.llm_parser as LP  # noqa: E402
from tabularedge.protocol import TableProfile, validate  # noqa: E402
from tabularedge.prompting import complete  # noqa: E402
from tabularedge.sanitize import color, sanitize  # noqa: E402

PROF = TableProfile("名单", ["序号", "统一社会信用代码", "单位名称"], {"序号": "num"}, 10)


def _lookup(i, ret):
    return {"id": f"s{i}", "op": "lookup_column", "sheet": "名单",
            "args": {"name": ret, "key": "统一社会信用代码", "from_sheet": f"表{i}", "from_key": "统一社会信用代码",
                     "return": ret, "if_not_found": "未匹配"}}


def _fake(monkeypatch, outputs):
    """依次返回给定的模型输出（字符串），记录调用次数。"""
    calls = []

    def stream(msgs, *a, **kw):
        calls.append(msgs)
        out = outputs[min(len(calls) - 1, len(outputs) - 1)]
        truncated = out.startswith("TRUNC:")
        return out.removeprefix("TRUNC:"), {"first_token_s": 0.1, "summary_s": 0.2, "total_s": 1.0, "truncated": truncated}
    monkeypatch.setattr(LP, "_stream", stream)
    monkeypatch.setattr(LP, "HARD_CASES", Path("/tmp") / "sr_test_hard_cases.jsonl")
    return calls


def test_style_aliases_and_types():
    raw = {"lang": "zh", "steps": [{"id": "s1", "op": "conditional_format", "args": {
        "where": {"col": "单位名称", "cmp": "等于", "value": "未匹配"}, "apply_to": "整行",
        "style": {"font_weight": "bold", "font-color": "red", "backgroundColor": "#ffeb9c", "underline": True}}}]}
    out, fixes = sanitize(raw)
    a = out["steps"][0]["args"]
    assert a["style"] == {"bold": True, "font_color": "FF0000", "fill_color": "FFEB9C"}
    assert a["where"]["cmp"] == "==" and a["apply_to"] == "row"
    assert any("underline" in f for f in fixes)
    assert validate(complete(out, "a.xlsx", "名单"), PROF) == []
    for v, want in [("weight", True)]:
        assert sanitize({"lang": "zh", "steps": [{"id": "s1", "op": "conditional_format", "args": {
            "where": {"col": "序号", "cmp": ">", "value": 1}, "style": {v: "700"}}}]})[0]["steps"][0]["args"]["style"] == {"bold": want}


def test_color_forms():
    assert color("Red") == "FF0000" and color("#f00") == "FF0000" and color("FFFF0000") == "FF0000"
    assert color("rgb(198, 239, 206)") == "C6EFCE" and color("浅绿") == "C6EFCE" and color("not a color") is None


def test_sanitizer_is_noop_on_valid_plans():
    p = {"lang": "zh", "summary": "x", "steps": [_lookup(1, "单位名称"), {"id": "s2", "op": "save", "args": {"path": "o.xlsx"}}]}
    out, fixes = sanitize(p)
    assert out == p and fixes == []


def test_structural_repairs():
    raw = {"lang": "zh", "actions": [
        {"id": "step1", "op": "highlight", "where": {"column": "序号", "operator": "gte", "val": "3"},
         "style": {"font": {"bold": "true", "color": "rgb(255,0,0)"}}},
        {"id": "step2", "op": "sort", "args": {"by": {"col": "序号", "desc": "yes"}}, "depends_on": "step1"}]}
    out, _ = sanitize(raw)
    assert [s["id"] for s in out["steps"]] == ["s1", "s2"] and out["steps"][1]["depends_on"] == ["s1"]
    assert out["steps"][0]["op"] == "conditional_format"
    assert out["steps"][0]["args"]["style"] == {"bold": True, "font_color": "FF0000"}
    assert out["steps"][1]["args"]["by"] == [{"col": "序号", "desc": True}]
    assert validate(complete(out, "a.xlsx", "名单"), PROF) == []


def test_parse_accepts_six_lookups_with_bad_style_first_try(monkeypatch):
    """6 个跨表查找全对，只有样式写成 font_weight → 必须一次通过、不重试、不回退。"""
    steps = [_lookup(i, f"字段{i}") for i in range(1, 7)]
    steps.append({"id": "s7", "op": "conditional_format", "sheet": "名单",
                  "args": {"where": {"col": "字段1", "cmp": "==", "value": "未匹配"}, "style": {"font_weight": "bold", "fill_color": "yellow"}}})
    steps.append({"id": "s8", "op": "save", "args": {"path": "名单_处理后.xlsx"}})
    calls = _fake(monkeypatch, [json.dumps({"summary": "六表匹配", "lang": "zh", "steps": steps}, ensure_ascii=False)])
    extra = "\n".join(f"[sheet=表{i} rows=5] columns: 统一社会信用代码:text, 字段{i}:text" for i in range(1, 7))
    plan, info = LP.parse("六表匹配……", PROF, "名单.xlsx", extra_profile=extra)
    assert len(calls) == 1 and not info.get("failed") and not info.get("fallback")
    assert [s["op"] for s in plan["steps"]].count("lookup_column") == 6
    assert plan["steps"][6]["args"]["style"] == {"bold": True, "fill_color": "FFFF00"}
    assert info["fixes"]


def test_parse_drops_unfixable_cosmetic_step_keeps_data_steps(monkeypatch):
    bad = {"summary": "x", "lang": "zh", "steps": [
        _lookup(1, "单位名称"),
        {"id": "s2", "op": "conditional_format", "args": {"style": {"bold": True}}},  # 缺 where，修不好
        {"id": "s3", "op": "save", "args": {"path": "o.xlsx"}}]}
    calls = _fake(monkeypatch, [json.dumps(bad, ensure_ascii=False)])
    plan, info = LP.parse("匹配单位名称并标黄", PROF, "名单.xlsx", max_repair=2)
    assert len(calls) == 3  # 先让模型自我修正两轮
    assert [s["op"] for s in plan["steps"]] == ["lookup_column", "save"]
    assert info["dropped"] == ["s2 条件格式"] and not info.get("failed")


def test_parse_unrecoverable_gives_specific_explanation_no_rule_fallback(monkeypatch):
    bad = {"summary": "x", "lang": "zh", "steps": [
        {"id": "s1", "op": "add_column", "args": {"name": "y", "formula": "=[不存在的列]*2"}},
        {"id": "s2", "op": "save", "args": {"path": "o.xlsx"}}]}
    _fake(monkeypatch, [json.dumps(bad, ensure_ascii=False)])
    assert not hasattr(LP, "rule_parser")  # 运行时管线里没有规则解析器
    plan, info = LP.parse("加一列", PROF, "名单.xlsx", max_repair=1)
    assert info["failed"] and plan["steps"][0]["op"] == "add_column"
    assert any("不存在的列" in x and "s1" in x for x in info["explain"])


def test_parse_truncated_output_stops_and_advises_split(monkeypatch):
    calls = _fake(monkeypatch, ['TRUNC:{"summary":"x","lang":"zh","steps":[{"id":"s1","op":"lookup_column","args":{"name":"a'])
    plan, info = LP.parse("超长指令", PROF, "名单.xlsx")
    assert len(calls) == 1 and info["failed"] and plan is None
    assert any("拆" in x for x in info["explain"])


def test_budget():
    assert LP.budget(300) == 2400 and LP.budget(4500) == 4096 - 3000 - 64 and LP.budget(9000) == 600


def test_prompt_compaction_with_many_reference_files(tmp_path):
    import client.core as core

    def book(name, headers, row):
        wb = Workbook()
        ws = wb.active
        ws.title = name
        ws.append(headers)
        ws.append(row)
        wb.save(tmp_path / f"{name}.xlsx")
    book("名单", ["序号", "统一社会信用代码"], [1, "911"])
    wide = [f"无关字段{i}" for i in range(30)]
    for n in ("工商表", "税务表", "社保表", "统计表", "专利表"):
        book(n, ["统一社会信用代码", "单位名称", *wide], ["911", "甲公司", *range(30)])
    text = "从工商表、税务表、社保表、统计表、专利表里按统一社会信用代码匹配单位名称"
    prof, extra, refs = core.prompt_context(tmp_path, "名单.xlsx", "名单", text)
    assert len(refs) == 5
    assert "无关字段" not in extra and extra.count("单位名称") == 5 and extra.count("统一社会信用代码") == 5
    # 3 个以内的参考文件：保持完整画像（与训练格式一致）
    _, extra3, _ = core.prompt_context(tmp_path, "名单.xlsx", "名单", "从工商表、税务表里按统一社会信用代码匹配单位名称")
    assert "无关字段29" in extra3


def test_sheet_refs_file_names_mapped_and_unknown_rejected(monkeypatch):
    extra = "[sheet=2026Q3_供应商名录 rows=9704] columns: 年:text\n[sheet=R07——000001 rows=329] columns: 单位详细名称:text"
    good = {"summary": "x", "lang": "zh", "steps": [
        {"id": "s1", "op": "lookup_column", "args": {"name": "甲类", "key": "单位名称", "from_sheet": "2026Q3_供应商名录.xlsx",
                                                     "from_key": "F", "return": "F", "if_not_found": ""}},
        {"id": "s2", "op": "lookup_column", "args": {"name": "乙类", "key": "单位名称", "from_sheet": "《R07——000001》",
                                                     "from_key": "B", "return": "B", "if_not_found": ""}}]}
    _fake(monkeypatch, [json.dumps(good, ensure_ascii=False)])
    plan, info = LP.parse("x", PROF, "名单.xlsx", extra_profile=extra)
    assert not info.get("failed") and [s["args"]["from_sheet"] for s in plan["steps"]] == ["2026Q3_供应商名录", "R07——000001"]
    bad = {**good, "steps": [{**good["steps"][0], "args": {**good["steps"][0]["args"], "from_sheet": "客户名录"}}]}
    _fake(monkeypatch, [json.dumps(bad, ensure_ascii=False)])
    plan, info = LP.parse("x", PROF, "名单.xlsx", extra_profile=extra, max_repair=0)
    assert info["failed"] and any("找不到表「客户名录」" in x for x in info["explain"])


def test_header_row_and_phantom_columns(tmp_path):
    """第 1 行是合并到第 239 列的大标题、第 2 行才是表头，整片刷过格式：画像只看真实的 7 列，新增列落在 H 列。"""
    from openpyxl.styles import Border, Side
    from client.core import RefExecutor, prompt_context
    from tabularedge.executor import Workspace
    from openpyxl import load_workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "单位名录"
    ws["A1"] = "示例单位基本信息"
    ws.merge_cells("A1:IE1")
    heads = ["序号", "校名", "办学内容", "学校地址", "举办者", "法定代表人姓名", "校长姓名"]
    for j, h in enumerate(heads, 1):
        ws.cell(2, j, h)
    for i in range(1, 24):
        for j, v in enumerate([i, f"学校{i}", "小学", "地址", "举办者", "张三", "李四"], 1):
            ws.cell(2 + i, j, v)
    thin = Side(style="thin")
    for r in range(1, 40):
        for c in range(1, 240):
            ws.cell(r, c).border = Border(left=thin)
    wb.save(tmp_path / "学校.xlsx")
    prof, extra, _ = prompt_context(tmp_path, "学校.xlsx", None, "按校名匹配")
    assert prof.header_row == 2 and prof.headers == heads and prof.n_rows == 23
    plan = {"version": "0.2", "lang": "zh", "inputs": [{"alias": "a", "path": "学校.xlsx", "sheet": "单位名录", "header_row": 2}],
            "steps": [{"id": "s1", "op": "add_column", "sheet": "单位名录", "args": {"name": "统计类型", "formula": '="其他"'}},
                      {"id": "s2", "op": "save", "args": {"path": "out.xlsx"}}]}
    RefExecutor(Workspace(tmp_path), None, []).run(plan)
    out = load_workbook(tmp_path / "out.xlsx")["单位名录"]
    assert out["H2"].value == "统计类型" and out["H3"].value == '="其他"' and out.max_column == 8 and out.max_row == 25
    assert [str(r) for r in out.merged_cells.ranges] == ["A1:G1"]  # 标题合并区收缩到数据宽度
