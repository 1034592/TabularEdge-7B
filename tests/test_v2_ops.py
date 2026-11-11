"""Action Plan v0.2：新函数前缀、跨表 XLOOKUP、FILTER / UNIQUE 到新表、破坏性操作确认。"""
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.formula import ArrayFormula

from tabularedge.executor import Executor, Workspace, xl


def _book(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "流水"
    for r in [["部门", "金额", "备注"], ["销售部", 100, "a"], ["财务部", -5, None], ["销售部", 2000, "b"], ["研发部", 50, "c"]]:
        ws.append(r)
    dim = wb.create_sheet("部门预算")
    for r in [["部门", "预算", "负责人"], ["销售部", 5000, "张伟"], ["财务部", 800, "王芳"]]:
        dim.append(r)
    wb.save(tmp_path / "t.xlsx")


def _plan(*steps):
    return {"version": "0.2", "lang": "zh", "inputs": [{"alias": "src", "path": "t.xlsx", "sheet": "流水"}],
            "steps": [*steps, {"id": "s9", "op": "save", "args": {"path": "out.xlsx"}}]}


def test_xlfn_prefix_idempotent_and_ignores_strings():
    assert xl('=XLOOKUP(A2,B:B,C:C,"XLOOKUP(")') == '=_xlfn.XLOOKUP(A2,B:B,C:C,"XLOOKUP(")'
    assert xl("=_xlfn.XLOOKUP(A2,B:B,C:C)") == "=_xlfn.XLOOKUP(A2,B:B,C:C)"
    assert xl("=SORTBY(A2:A9,B2:B9)+SORT(A2:A9)") == "=_xlfn.SORTBY(A2:A9,B2:B9)+_xlfn._xlws.SORT(A2:A9)"
    assert xl("=SUMIFS(A:A,B:B,1)+MAXIFS(A:A,B:B,1)") == "=SUMIFS(A:A,B:B,1)+_xlfn.MAXIFS(A:A,B:B,1)"
    assert xl("=TEXTJOIN(\"-\",TRUE,A2,B2)") == "=_xlfn.TEXTJOIN(\"-\",TRUE,A2,B2)"


def test_lookup_column_formula_and_value(tmp_path):
    _book(tmp_path)
    Executor(Workspace(tmp_path)).run(_plan({"id": "s1", "op": "lookup_column", "sheet": "流水", "args": {
        "name": "预算", "key": "部门", "from_sheet": "部门预算", "from_key": "部门", "return": "预算", "if_not_found": "未找到"}}))
    ws = load_workbook(tmp_path / "out.xlsx")["流水"]
    assert ws["D1"].value == "预算"
    assert ws["D2"].value == "=_xlfn.XLOOKUP(A2,'部门预算'!$A:$A,'部门预算'!$B:$B,\"未找到\")"
    _book(tmp_path)
    Executor(Workspace(tmp_path)).run(_plan({"id": "s1", "op": "lookup_column", "sheet": "流水", "args": {
        "name": "负责人", "key": "部门", "from_sheet": "部门预算", "from_key": "部门", "return": "负责人",
        "if_not_found": "未找到", "mode": "value"}}))
    ws = load_workbook(tmp_path / "out.xlsx")["流水"]
    assert [ws.cell(r, 4).value for r in range(2, 6)] == ["张伟", "王芳", "张伟", "未找到"]


def test_filter_to_sheet_live_array(tmp_path):
    _book(tmp_path)
    log = Executor(Workspace(tmp_path)).run(_plan({"id": "s1", "op": "filter_to_sheet", "sheet": "流水", "args": {
        "where": {"col": "金额", "cmp": ">", "value": 60}, "output_sheet": "大额"}}))
    assert "2 rows" in log[0] and "live FILTER" in log[0]
    out = load_workbook(tmp_path / "out.xlsx")["大额"]
    assert [c.value for c in out[1]] == ["部门", "金额", "备注"]
    f = out["A2"].value
    assert isinstance(f, ArrayFormula) and f.ref == "A2:C3"
    assert f.text == "=_xlfn._xlws.FILTER('流水'!$A$2:$C$5,(ISNUMBER('流水'!$B$2:$B$5)*('流水'!$B$2:$B$5>60)),\"\")"


def test_filter_regex_falls_back_to_values(tmp_path):
    _book(tmp_path)
    log = Executor(Workspace(tmp_path)).run(_plan({"id": "s1", "op": "filter_to_sheet", "sheet": "流水", "args": {
        "where": {"col": "备注", "cmp": "regex", "value": "^[ab]$"}, "output_sheet": "ab"}}))
    assert "(values)" in log[0]
    out = load_workbook(tmp_path / "out.xlsx")["ab"]
    assert [r for r in out.iter_rows(min_row=2, values_only=True)] == [("销售部", 100, "a"), ("销售部", 2000, "b")]


def test_unique_to_sheet_with_count(tmp_path):
    _book(tmp_path)
    Executor(Workspace(tmp_path)).run(_plan({"id": "s1", "op": "unique_to_sheet", "sheet": "流水", "args": {
        "col": "部门", "output_sheet": "部门清单", "with_count": True}}))
    out = load_workbook(tmp_path / "out.xlsx")["部门清单"]
    assert [c.value for c in out[1]] == ["部门", "计数"]
    assert out["A2"].value.ref == "A2:A4"
    assert out["B3"].value == "=COUNTIF('流水'!$A:$A,A3)"


def test_risky_ops_need_confirmation(tmp_path):
    _book(tmp_path)
    asked = []
    log = Executor(Workspace(tmp_path), confirm=lambda st: asked.append(st["op"]) or st["op"] != "delete_rows").run(
        _plan({"id": "s1", "op": "delete_rows", "sheet": "流水", "args": {"where": {"col": "金额", "cmp": "<", "value": 0}}}))
    assert asked == ["delete_rows", "save"]
    assert "skipped (user declined)" in log[0]
    assert load_workbook(tmp_path / "out.xlsx")["流水"].max_row == 5
