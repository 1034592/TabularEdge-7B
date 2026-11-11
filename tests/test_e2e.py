import json

import pytest
from openpyxl import Workbook, load_workbook

from tabularedge.executor import Executor, SandboxError, Workspace
from tabularedge.protocol import PlanError, check, profile_sheet
from benchmarks.rule_baseline import parse

ROWS = [
    ["日期", "部门", "金额", "备注"],
    ["2026-10-01", "销售部", 1200, "客户A（首付）"],
    ["2026-10-02", "财务部", -300, "退款(重复扣款)"],
    ["2026-10-03", "销售部", 800, "客户B"],
    ["2026-10-05", "研发部", -50.5, "  报销（打车） "],
    ["2026-10-06", "财务部", 1000, None],
]


@pytest.fixture
def ws_root(tmp_path):
    wb = Workbook()
    wb.active.title = "流水"
    for r in ROWS:
        wb.active.append(r)
    wb.save(tmp_path / "ledger.xlsx")
    return tmp_path


def _profile(ws_root):
    return profile_sheet(load_workbook(ws_root / "ledger.xlsx")["流水"])


def test_zh_flagship_instruction(ws_root):
    prof = _profile(ws_root)
    plan = parse("把带括号的备注删了，金额负数的标红，按部门汇总新表", prof, "ledger.xlsx")
    check(plan, prof)
    ops = [s["op"] for s in plan["steps"]]
    assert ops == ["regex_replace", "conditional_format", "group_aggregate", "save"]
    assert plan["ambiguities"][0]["alternatives"][0]["op"] == "delete_rows"  # 歧义被显式暴露
    assert plan["steps"][2]["depends_on"] == ["s1"]

    log = Executor(Workspace(ws_root)).run(plan)
    out = load_workbook(ws_root / "ledger_处理后.xlsx")
    notes = [c.value for c in out["流水"]["D"][1:]]
    assert notes == ["客户A", "退款", "客户B", "报销", None]
    rule = out["流水"].conditional_formatting
    assert any("$C2<0" in r.formula[0] for cf in rule for r in cf.rules)
    summ = out["按部门汇总"]
    assert [c.value for c in summ[1]] == ["部门", "sum_金额"]
    assert summ["B2"].value.startswith("=SUMIFS('流水'!$C:$C,'流水'!$B:$B,$A2")
    # 源文件未被改动
    assert load_workbook(ws_root / "ledger.xlsx")["流水"]["D2"].value == "客户A（首付）"
    assert len(log) == 4


def test_en_colloquial_on_chinese_headers(ws_root):
    prof = _profile(ws_root)
    plan = parse("strip the stuff in parentheses from the notes, make negative amounts red "
                 "and then total amount by department", prof, "ledger.xlsx")
    check(plan, prof)
    assert plan["lang"] == "en"
    assert [s["op"] for s in plan["steps"]][:3] == ["regex_replace", "conditional_format", "group_aggregate"]
    assert plan["steps"][2]["args"]["by"] == ["部门"]


def test_value_mode_numbers(ws_root):
    prof = _profile(ws_root)
    plan = parse("按部门汇总金额", prof, "ledger.xlsx")
    plan["steps"][0]["args"]["mode"] = "value"
    Executor(Workspace(ws_root)).run(plan)
    s = load_workbook(ws_root / "ledger_处理后.xlsx")["按部门汇总"]
    got = {r[0]: r[1] for r in s.iter_rows(min_row=2, values_only=True)}
    assert got == {"财务部": 700, "研发部": -50.5, "销售部": 2000}


def test_semantic_errors_are_specific(ws_root):
    prof = _profile(ws_root)
    bad = {"version": "0.1", "lang": "zh", "steps": [
        {"id": "s1", "op": "sort", "args": {"by": [{"col": "客户"}]}},
        {"id": "s2", "op": "regex_replace", "args": {"col": "备注", "pattern": "(("}, "depends_on": ["s9"]}]}
    with pytest.raises(PlanError) as e:
        check(bad, prof)
    msg = " ".join(e.value.errors)
    assert "unknown column '客户'" in msg and "bad regex" in msg and "depends_on s9" in msg


def test_schema_rejects_unknown_op_and_bad_args():
    with pytest.raises(PlanError):
        check({"version": "0.1", "lang": "en", "steps": [{"id": "s1", "op": "rm_rf", "args": {}}]})
    with pytest.raises(PlanError):
        check({"version": "0.1", "lang": "en", "steps": [{"id": "s1", "op": "trim", "args": {"col": "A"}}]})


def test_sandbox_blocks_escape(ws_root):
    plan = {"version": "0.1", "lang": "en", "inputs": [{"alias": "src", "path": "ledger.xlsx"}],
            "steps": [{"id": "s1", "op": "save", "args": {"path": "../../etc/evil.xlsx"}}]}
    with pytest.raises(SandboxError):
        Executor(Workspace(ws_root)).run(plan)


def test_file_ops(ws_root):
    ex = Executor(Workspace(ws_root))
    plan = {"version": "0.1", "lang": "zh", "inputs": [{"alias": "src", "path": "ledger.xlsx"}], "steps": [
        {"id": "s1", "op": "fs_backup", "args": {"path": "ledger.xlsx"}},
        {"id": "s2", "op": "fs_mkdir", "args": {"path_template": "2026/{mm}/{dept}",
                                                "expand": {"mm": ["10月", "11月"], "dept": ["销售部", "财务部"]}}},
        {"id": "s3", "op": "split_to_files", "args": {"by": "部门", "path_template": "2026/10月/{value}/明细.xlsx"}},
        {"id": "s4", "op": "fs_merge_workbooks", "args": {"glob": "2026/10月/*/明细.xlsx", "output": "合并/10月.xlsx"}},
        {"id": "s5", "op": "fs_new_workbook", "args": {"path": "2026/11月/销售部/明细.xlsx", "headers": ["日期", "部门", "金额", "备注"]}},
    ]}
    ex.run(plan)
    assert len(list((ws_root / "_backup").glob("ledger.*.xlsx"))) == 1
    assert (ws_root / "2026/11月/财务部").is_dir()
    merged = load_workbook(ws_root / "合并/10月.xlsx").active
    assert merged.max_row == 1 + 5 and merged["A1"].value == "来源文件"
    with pytest.raises(FileExistsError):
        ex.op_fs_new_workbook({}, {"path": "2026/11月/销售部/明细.xlsx", "headers": ["x"]})


def test_confirm_hook_can_veto_writes(ws_root):
    prof = _profile(ws_root)
    plan = parse("金额负数的标红", prof, "ledger.xlsx")
    log = Executor(Workspace(ws_root), confirm=lambda s: False).run(plan)
    assert "skipped" in log[-1] and not (ws_root / "ledger_处理后.xlsx").exists()


def test_sort_mixed_types_and_empty_last(tmp_path):
    wb = Workbook()
    wb.active.append(["k", "v"])
    for r in [["a", 3], ["b", None], ["c", "x"], ["d", 10], ["e", ""]]:
        wb.active.append(r)
    wb.save(tmp_path / "m.xlsx")
    plan = {"version": "0.1", "lang": "en", "inputs": [{"alias": "src", "path": "m.xlsx"}], "steps": [
        {"id": "s1", "op": "sort", "args": {"by": [{"col": "v", "desc": True}]}},
        {"id": "s2", "op": "save", "args": {"path": "o.xlsx"}}]}
    Executor(Workspace(tmp_path)).run(plan)
    ks = [r[0] for r in load_workbook(tmp_path / "o.xlsx").active.iter_rows(min_row=2, values_only=True)]
    assert ks[:3] == ["c", "d", "a"] and set(ks[3:]) == {"b", "e"}  # 文本>数字(降序)，空值永远在最后


def test_conditional_format_after_all_rows_deleted(ws_root):
    plan = {"version": "0.1", "lang": "zh", "inputs": [{"alias": "src", "path": "ledger.xlsx"}], "steps": [
        {"id": "s1", "op": "delete_rows", "args": {"where": {"col": "部门", "cmp": "not_empty"}}},
        {"id": "s2", "op": "conditional_format", "args": {"where": {"col": "金额", "cmp": "<", "value": 0}, "style": {"font_color": "FF0000"}}}]}
    log = Executor(Workspace(ws_root)).run(plan)
    assert "skipped" in log[-1]
