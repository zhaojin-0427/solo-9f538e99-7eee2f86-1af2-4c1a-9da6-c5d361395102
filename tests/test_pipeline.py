import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from openpyxl import Workbook, load_workbook

from app.engine.formula import parse_formula, unparse
from app.engine.model import WorkbookModel
from app.engine.tracer import parse_a1, parse_rect, trace
from app.engine.extractor import build_extract


@pytest.fixture()
def wb_path(tmp_path):
    def _build(name="case.xlsx"):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "表1"
        ws2 = wb.create_sheet("表2")
        # 表1：A1:A3 数据，B5 指标（四则 + 跨表 + 绝对地址 + SUM 区域含空格）
        ws1["A1"] = 10
        ws1["A2"] = 20
        ws1["A3"] = 30
        ws1["C1"] = 2
        ws1["C2"] = "=A1*$C$1+A2"           # 10*2+20 = 40
        ws1["D1"] = "=SUM(A1:A3)"           # 60（区域含空格也算引用位置）
        ws1["B5"] = "=C2+D1+表2!B2-A3"      # 40+60+5-30 = 75
        # 表2
        ws2["A1"] = 3
        ws2["A2"] = 2
        ws2["B2"] = "=A1+A2"                # 5
        path = tmp_path / name
        wb.save(path)
        return str(path)
    return _build


def test_parse_and_unparse_keeps_absolute():
    pf = parse_formula("=A1*$C$1+表2!$B$2", {"表2"})
    col_map = {("s", 1): 4, ("s", 3): 2, ("表2", 2): 1}
    row_map = {("s", 1): 1, ("s", 2): 5, ("表2", 2): 9}
    out = unparse(pf.ast, col_map, row_map, "s")
    assert "$A$9" in out          # 表2!$B$2 -> 表2!$A$9，绝对标记保留
    assert "表2!" in out
    assert "$B$1" in out          # $C$1 -> $B$1，绝对标记保留
    assert "D1" in out            # A1 -> D1（相对引用）


def test_unsupported_function_and_string():
    with pytest.raises(Exception):
        parse_formula("=IF(A1>1,2,3)", {"s"})
    with pytest.raises(Exception):
        parse_formula('=A1&"x"', {"s"})


def test_trace_layers_and_required(wb_path):
    model = WorkbookModel(wb_path())
    ind = parse_a1("表1!B5", model.sheet_names)
    result = trace(model, [ind], [])
    assert not result.errors
    required = result.required
    # SUM 区域 A1:A3 的空格位置也算引用（此处无空格，验证全覆盖）
    for r in range(1, 4):
        assert ("表1", 1, r) in required
    assert ("表2", 2, 2) in required
    assert ("表1", 3, 1) in required      # $C$1
    assert result.layers[("表1", 3, 2)] >= 1


def test_forbidden_chain(wb_path):
    model = WorkbookModel(wb_path())
    ind = parse_a1("表1!B5", model.sheet_names)
    rect = parse_rect("表1!A1:A1", model.sheet_names)
    result = trace(model, [ind], [rect])
    assert result.errors
    e = result.errors[0]
    assert e.kind == "forbidden"
    assert e.indicator_text == "表1!B5"
    assert "表1!A1" in e.chain_text
    assert "表1!B5" in e.chain_text


def test_cycle_detection(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "=B1+1"
    ws["B1"] = "=A1+1"
    path = tmp_path / "cyc.xlsx"
    wb.save(path)
    model = WorkbookModel(str(path))
    result = trace(model, [("S", 1, 1)], [])
    assert len(result.errors) == 1
    assert result.errors[0].kind == "cycle"
    assert "循环引用" in result.errors[0].message


def test_text_direct_reference(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "文本"
    ws["A2"] = "=A1+1"
    path = tmp_path / "t.xlsx"
    wb.save(path)
    model = WorkbookModel(str(path))
    result = trace(model, [("Sheet", 1, 2)], [])
    assert result.errors and result.errors[0].kind == "text"


def test_extract_rewrite_and_recompute(wb_path):
    model = WorkbookModel(wb_path())
    ind = parse_a1("表1!B5", model.sheet_names)
    result = trace(model, [ind], [])
    out = wb_path("out.xlsx").replace("case.xlsx", "out.xlsx")
    er = build_extract(model, result, out)
    assert not er.verify_errors
    assert all(row[3] for row in er.verify)
    assert er.verify[0][1] == 75 and er.verify[0][2] == 75
    # 新文件行列被压缩
    wb2 = load_workbook(out)
    ws1 = wb2["表1"]
    assert ws1.max_row < 5 or ws1.max_column < 4
    # 新文件公式仍是受支持公式，且可被重新解析复算
    found_sum = False
    for row in ws1.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                parse_formula(c.value, set(wb2.sheetnames))
                if "SUM(" in c.value:
                    found_sum = True
    assert found_sum
    # 跨表引用仍带表名
    cross = [c.value for row in ws1.iter_rows() for c in row
             if isinstance(c.value, str) and "表2!" in c.value]
    assert cross


def test_sum_blank_cell_is_required(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 1
    # A2 故意留空
    ws["A3"] = 3
    ws["C1"] = "=SUM(A1:A3)"
    path = tmp_path / "blank.xlsx"
    wb.save(path)
    model = WorkbookModel(str(path))
    result = trace(model, [("S", 3, 1)], [])
    assert ("S", 1, 2) in result.required       # 空格也算引用位置
    er = build_extract(model, result, str(tmp_path / "blank_out.xlsx"))
    assert er.verify[0][1] == 4


def test_merged_cells_kept(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 5
    ws.merge_cells("A1:B1")
    ws["D2"] = "=A1+1"
    path = tmp_path / "m.xlsx"
    wb.save(path)
    model = WorkbookModel(str(path))
    result = trace(model, [("S", 4, 2)], [])
    assert not result.errors
    er = build_extract(model, result, str(tmp_path / "m_out.xlsx"))
    wb2 = load_workbook(er.output_path)
    assert len(wb2["S"].merged_cells.ranges) == 1


def test_upload_limits(tmp_path):
    from app import create_app
    from app.engine.model import WorkbookLoadError
    # 7 张表 -> 拒绝
    wb = Workbook()
    for i in range(6):
        wb.create_sheet(f"S{i}")
    p = tmp_path / "seven.xlsx"
    wb.save(p)
    with pytest.raises(WorkbookLoadError):
        WorkbookModel(str(p))
    # 201 行 -> 拒绝
    wb = Workbook()
    wb.active.cell(row=201, column=1, value=1)
    p2 = tmp_path / "big.xlsx"
    wb.save(p2)
    with pytest.raises(WorkbookLoadError):
        WorkbookModel(str(p2))
    app = create_app(instance_path=str(tmp_path / "inst1"))
    client = app.test_client()
    with open(p, "rb") as fh:
        rv = client.post("/upload", data={"workbook": (fh, "seven.xlsx")},
                         content_type="multipart/form-data")
    assert rv.status_code == 400


def test_division_by_zero_and_multi_sum(tmp_path):
    from app.engine.formula import FormulaError
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 1
    ws["A2"] = "=A1/0"
    p = tmp_path / "z.xlsx"
    wb.save(p)
    model = WorkbookModel(str(p))
    result = trace(model, [("S", 1, 2)], [])
    # 追溯不报结构错误，但复算核对应报除零
    assert not result.errors
    er = build_extract(model, result, str(tmp_path / "z_out.xlsx"))
    assert any("除以零" in m for m in er.verify_errors)
    with pytest.raises(FormulaError):
        parse_formula("=SUM(A1:A2,B1:B2)", {"S"})


def test_multiple_indicators_and_layout(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 3
    ws["A2"] = 4
    ws["B1"] = "=A1*2"     # 6
    ws["B2"] = "=A2+A1"    # 7
    ws["Z9"] = 99          # 用不到的远格，应被行列裁剪掉
    p = tmp_path / "m2.xlsx"
    wb.save(p)
    model = WorkbookModel(str(p))
    result = trace(model, [("S", 2, 1), ("S", 2, 2)], [])
    er = build_extract(model, result, str(tmp_path / "m2_out.xlsx"))
    assert len(er.indicator_mapping) == 2
    wb2 = load_workbook(er.output_path)
    ws2 = wb2["S"]
    assert ws2.max_column == 2 and ws2.max_row == 2   # Z 列与第 9 行被删
    vals = {c.coordinate: c.value for row in ws2.iter_rows() for c in row}
    assert vals == {"A1": 3, "B1": "=(A1*2)", "A2": 4, "B2": "=(A2+A1)"}
    assert {v[1] for v in er.verify} == {6, 7}


def test_web_flow(wb_path, tmp_path):
    from app import create_app
    app = create_app(instance_path=str(tmp_path / "instance"))
    client = app.test_client()
    # 上传
    with open(wb_path(), "rb") as fh:
        rv = client.post("/upload", data={"workbook": (fh, "case.xlsx")},
                         content_type="multipart/form-data")
    assert rv.status_code == 200
    # 正常分析
    rv = client.post("/analyze", data={
        "ind_sheet": ["表1"], "ind_addr": ["B5"],
        "fb_sheet": ["表1"], "fb_addr": [""],
    })
    assert rv.status_code == 200
    assert "复算核对".encode() in rv.data
    rv = client.get("/download")
    assert rv.status_code == 200
    assert rv.data[:2] == b"PK"

    # 禁发场景：重新上传
    with open(wb_path(), "rb") as fh:
        client.post("/upload", data={"workbook": (fh, "case2.xlsx")},
                    content_type="multipart/form-data")
    rv = client.post("/analyze", data={
        "ind_sheet": ["表1"], "ind_addr": ["B5"],
        "fb_sheet": ["表1"], "fb_addr": ["A1:A1"],
    })
    assert "追溯未通过".encode() in rv.data
    rv = client.get("/download")
    assert rv.status_code == 404


def test_out_of_dimension_blank_reference(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 10
    ws["A2"] = "=A1+D1"      # D1 在数据区域外，按空格 0
    ws["C1"] = "=A2*2"
    p = tmp_path / "oob.xlsx"
    wb.save(p)
    model = WorkbookModel(str(p))
    result = trace(model, [("S", 3, 1)], [])
    assert not result.errors
    er = build_extract(model, result, str(tmp_path / "oob_out.xlsx"))
    assert not er.verify_errors and er.verify[0][3]
    wb2 = load_workbook(er.output_path)
    # 改写后的公式仍可解析且引用到新文件内的真实空格
    for row in wb2.active.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                parse_formula(c.value, set(wb2.sheetnames))


def test_web_verify_failure_blocks_download(tmp_path):
    from app import create_app
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 1
    ws["A2"] = "=A1/0"
    p = tmp_path / "z.xlsx"
    wb.save(p)
    app = create_app(instance_path=str(tmp_path / "inst2"))
    client = app.test_client()
    with open(p, "rb") as fh:
        client.post("/upload", data={"workbook": (fh, "z.xlsx")},
                    content_type="multipart/form-data")
    rv = client.post("/analyze", data={
        "ind_sheet": ["S"], "ind_addr": ["A2"],
        "fb_sheet": ["S"], "fb_addr": [""],
    })
    assert "复算核对未通过".encode() in rv.data
    assert client.get("/download").status_code == 404
