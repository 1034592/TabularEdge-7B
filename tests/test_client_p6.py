"""P6（桌面客户端）新增模块的回归测试：公式求值、Excel 2019 改写、删行/排序的公式引用调整、[列] 模板展开、模型输出规整。"""
import bisect
import pytest
import datetime as dt
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from openpyxl import Workbook, load_workbook  # noqa: E402

from client.core import tidy_save_paths  # noqa: E402
from tabularedge.compat import incompatible, to_2019  # noqa: E402
from tabularedge.executor import Executor, Workspace, adjust_formula, expand_template  # noqa: E402
from tabularedge.formula_eval import Evaluator  # noqa: E402
from tabularedge.llm_parser import _fix_formula, _repair_json, normalize  # noqa: E402


def _wb():
    wb = Workbook()
    ws = wb.active
    ws.title = "明细"
    for r in [["工号", "销售额", "日期", "备注"], ["E01", 120000, dt.datetime(2026, 3, 5), "退款(SO123)"],
              ["E02", 50000, "2026年3月9日", "SO777 正常"], ["E01", 80000, "2026/3/16", None]]:
        ws.append(r)
    t = wb.create_sheet("档位")
    for r in [["下限", "比例"], [0, 0.01], [60000, 0.02], [100000, 0.03]]:
        t.append(r)
    return wb, ws


def test_eval_core_functions():
    wb, ws = _wb()
    E = Evaluator(wb)
    assert E.eval("=XLOOKUP(B2,档位!$A:$A,档位!$B:$B,0,-1)", ws) == 0.03
    assert E.eval("=VLOOKUP(B3,档位!A:B,2,TRUE)", ws) == 0.01
    assert E.eval("=SUMIFS($B:$B,$A:$A,A2)", ws) == 200000
    assert E.eval("=ISOWEEKNUM(DATEVALUE(C3))", ws) == 11
    assert E.eval('=TEXT(C2,"yyyy-mm")', ws) == "2026-03"
    assert E.eval('=IFERROR(REGEXEXTRACT(D2,"SO\\d+"),"")', ws) == "SO123"
    assert E.eval('=IF(B2=MAXIFS($B:$B,$A:$A,A2),"保留","归档")', ws) == "保留"
    assert E.eval("=PERCENTRANK.INC($B$2:$B$4,B3)", ws) == 0.0
    assert E.eval('=IFERROR(VLOOKUP("x",A:B,2,FALSE),"无")', ws) == "无"


def test_compat_2019_rewrite_equivalent():
    wb, ws = _wb()
    E = Evaluator(wb)
    for f in ["=XLOOKUP(B2,档位!$A:$A,档位!$B:$B,0,-1)", '=XLOOKUP(A3,$A:$A,$B:$B,"无")', '=XLOOKUP("E9",$A:$A,$B:$B,"无")']:
        g = to_2019(f)
        assert not incompatible(g), g
        assert E.eval(f, ws) == E.eval(g, ws), (f, g)
    assert incompatible("=TEXTBEFORE(A2,\"-\")") == {"TEXTBEFORE"}


def test_adjust_formula_on_delete():
    dels = [3, 5]

    def rm(r, end=None):
        if r in dels:
            if end == "start":
                while r in dels:
                    r += 1
            elif end == "end":
                while r in dels:
                    r -= 1
            else:
                return None
        return r - bisect.bisect_left(dels, r)
    assert adjust_formula('=MID(E6,SEARCH("SO",E6),13)', "s", "s", rm) == '=MID(E4,SEARCH("SO",E4),13)'
    assert adjust_formula("=SUM($B$2:$B$10)", "s", "s", rm) == "=SUM($B$2:$B$8)"
    assert adjust_formula("=XLOOKUP(A7,档位!$A:$A,档位!$B:$B)", "s", "s", rm) == "=XLOOKUP(A5,档位!$A:$A,档位!$B:$B)"
    assert adjust_formula('=IF(E3="A6",1,0)', "s", "s", rm) == '=IF(#REF!="A6",1,0)'


def test_delete_rows_keeps_formulas_pointing_at_own_row(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "t"
    ws.append(["k", "v"])
    for i in range(1, 7):
        ws.append([f"r{i}", i])
    wb.save(tmp_path / "a.xlsx")
    ex = Executor(Workspace(tmp_path))
    ex.run({"version": "0.2", "lang": "zh", "inputs": [{"alias": "src", "path": "a.xlsx", "sheet": "t"}], "steps": [
        {"id": "s1", "op": "add_column", "sheet": "t", "args": {"name": "x2", "formula": "=[v]*2"}},
        {"id": "s2", "op": "delete_rows", "sheet": "t", "args": {"where": {"col": "v", "cmp": "<=", "value": 3}}},
        {"id": "s3", "op": "save", "args": {"path": "b.xlsx"}}]})
    out = load_workbook(tmp_path / "b.xlsx")["t"]
    ev = Evaluator(out.parent)
    assert [ev.cell(out, r, 3) for r in range(2, out.max_row + 1)] == [8, 10, 12]


def test_expand_template_range_args():
    H = ["订单号", "客户", "金额", "状态", "更新时间", "备注"]
    assert expand_template('=IF([更新时间]=MAXIFS([更新时间],[订单号],[订单号]),"保留","归档")', H) == \
        '=IF(E{r}=MAXIFS($E:$E,$A:$A,A{r}),"保留","归档")'
    assert expand_template("=[金额]/SUM([金额])", H) == "=C{r}/SUM($C:$C)"
    assert expand_template("=SUM([金额],[金额])", H) == "=SUM(C{r},C{r})"
    assert expand_template('=REGEXEXTRACT([备注],"SO[\\dA-Z-]+")', H) == '=REGEXEXTRACT(F{r},"SO[\\dA-Z-]+")'


def test_normalize_model_output():
    assert _fix_formula("=IF([回款状态]=='未回款',0,1)") == '=IF([回款状态]="未回款",0,1)'
    assert _fix_formula("=XLOOKUP(A2,'提成档位'!A:A,'提成档位'!B:B)") == "=XLOOKUP(A2,'提成档位'!A:A,'提成档位'!B:B)"
    assert _fix_formula('=REGEX_EXTRACT(A1,"x")') == '=REGEXEXTRACT(A1,"x")'
    import json
    raw = '{"steps":[{"id":"s3","op":"add_column","args":{"name":"年月","formula":"=TEXT([d],"yyyy-mm")"}}]}'
    assert json.loads(_repair_json(raw))["steps"][0]["args"]["formula"] == '=TEXT([d],"yyyy-mm")'
    obj = normalize({"steps": [{"op": "lookup_column", "args": {"if_not_found": None, "mode": "lookup_max", "name": "a"}}]})
    assert obj["steps"][0]["args"] == {"name": "a"}


def test_tidy_save_paths(tmp_path):
    (tmp_path / "X_处理后.xlsx").touch()
    p = {"steps": [{"op": "save", "args": {"path": "X_处理后_处理后_处理后.xlsx"}}]}
    assert tidy_save_paths(p, "X_处理后.xlsx", tmp_path)["steps"][0]["args"]["path"] == "X_处理后(2).xlsx"
    p = {"steps": [{"op": "save", "args": {"path": "X_处理后.xlsx"}}]}  # 明确覆盖当前文件：不动
    assert tidy_save_paths(p, "X_处理后.xlsx", tmp_path)["steps"][0]["args"]["path"] == "X_处理后.xlsx"


# ---------------------------------------------------------------- 大表（流式读取，不经 openpyxl 全量载入）
def _mixed_wb(path, n=60):
    wb = Workbook()
    ws = wb.active
    ws.title = "单位"
    ws.append(["代码", "名称", "人数", "成立日期", "备注"])
    for i in range(1, n + 1):
        ws.append([f"C{i:04d}" if i % 3 else i, f"公司&<{i}>", i * 1.5 if i % 2 else i, dt.datetime(2020, 1, 1) + dt.timedelta(days=i),
                   None if i % 4 else "  有空格 "])
    wb.create_sheet("说明").append(["说明"])
    wb.save(path)


def test_stream_profile_matches_openpyxl(tmp_path):
    from client.core import profile_eval, stream_profile
    from tabularedge.xlsx_stream import XlsxStream
    _mixed_wb(tmp_path / "m.xlsx")
    wb = load_workbook(tmp_path / "m.xlsx")
    with XlsxStream(tmp_path / "m.xlsx") as xs:
        assert xs.sheet_names == ["单位", "说明"]
        assert xs.dims("单位") == (61, 5)
        for ws in wb.worksheets:
            a, b = profile_eval(ws, Evaluator(wb)), stream_profile(xs, ws.title)
            assert (a.headers, a.types, a.n_rows, a.samples) == (b.headers, b.types, b.n_rows, b.samples)


def test_shared_string_variants():
    from tabularedge.xlsx_stream import _si_text
    u = lambda x: x.encode("utf-8")  # noqa: E731
    assert _si_text(b"<t>a&amp;b</t>") == "a&b"
    assert _si_text(u('<r><t>示例</t></r><r><rPr/><t xml:space="preserve"> 区</t></r>')) == "示例 区"
    assert _si_text(u('<t>北京</t><rPh sb="0" eb="2"><t>ペキン</t></rPh>')) == "北京"
    assert _si_text(b"<t>a_x000D_b</t>") == "a\rb"


def test_big_reference_streamed_lookup_equals_normal(tmp_path, monkeypatch):
    import client.core as core
    from client.core import RefExecutor, find_refs
    _mixed_wb(tmp_path / "大参考表.xlsx", 300)
    wb = Workbook()
    ws = wb.active
    ws.title = "名单"
    ws.append(["序号", "代码"])
    for k, c in enumerate(["C0001", "C0002", 3, "3", " C0005 ", "X9", 300, None], 1):
        ws.append([k, c])
    wb.save(tmp_path / "名单.xlsx")
    plan = {"version": "0.2", "lang": "zh", "inputs": [{"alias": "a", "path": "名单.xlsx", "sheet": "名单"}], "steps": [
        {"id": "s1", "op": "lookup_column", "sheet": "名单", "args": {"name": "名称", "key": "代码", "from_sheet": "单位",
                                                                    "from_key": "代码", "return": "名称", "if_not_found": "未匹配",
                                                                    "mode": "value"}},
        {"id": "s2", "op": "lookup_join", "sheet": "名单", "args": {"right": "单位", "on": "代码", "fetch": ["人数", "成立日期"]}},
        {"id": "s3", "op": "save", "args": {"path": "out.xlsx"}}]}
    refs = find_refs(tmp_path, "名单.xlsx", "从大参考表里匹配名称", ["名单"])
    assert refs == [("单位", str(tmp_path / "大参考表.xlsx"), "单位"), ("说明", str(tmp_path / "大参考表.xlsx"), "说明")]
    monkeypatch.setattr(core, "BIG_REF_BYTES", 0)
    log = RefExecutor(Workspace(tmp_path), None, refs).run(plan)
    assert any("streamed lookup only" in x for x in log)
    out = list(load_workbook(tmp_path / "out.xlsx")["名单"].iter_rows(values_only=True))
    assert load_workbook(tmp_path / "out.xlsx").sheetnames == ["名单"]  # 大文件的表都没被复制进输出
    assert [r[2] for r in out[1:]] == ["公司&<1>", "公司&<2>", "公司&<3>", "公司&<3>", "公司&<5>", "未匹配", "公司&<300>", "未匹配"]
    assert out[1][3:] == (1.5, dt.datetime(2020, 1, 2))
    # 公式里直接引用大表 → 明确报错，而不是把 76 万行塞进输出文件
    bad = {**plan, "steps": [{"id": "s1", "op": "add_column", "sheet": "名单",
                              "args": {"name": "x", "formula": "=VLOOKUP([代码],单位!A:B,2,FALSE)"}}]}
    with pytest.raises(ValueError, match="太大"):
        RefExecutor(Workspace(tmp_path), None, refs).run(bad)


def test_big_main_table_guard(tmp_path, monkeypatch):
    import client.core as core
    from client.core import RefExecutor
    _mixed_wb(tmp_path / "m.xlsx")
    monkeypatch.setattr(core, "BIG_MAIN_CELLS", 100)
    with pytest.raises(ValueError, match="逐格编辑"):
        RefExecutor(Workspace(tmp_path), None, []).run({"version": "0.2", "inputs": [{"alias": "a", "path": "m.xlsx"}], "steps": []})
