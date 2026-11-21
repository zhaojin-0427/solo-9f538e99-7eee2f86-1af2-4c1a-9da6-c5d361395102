"""针对“同一底稿多次/交错分析”缺陷的回归测试。

覆盖：
- 同一份底稿先后选不同指标生成两份审计件，结果页与下载件各自对应本次提交；
- 交错提交（一次通过、一次失败、再次通过）互不覆盖、互不删除；
- 复算失败的提交不开放下载，且页面展示本次失败原因；
- 结果页地址对照/复算值与下载文件内容一致；
- 服务重启（用同一 instance_path 重新 create_app）后旧链接仍可下载原文件；
- 猜测/伪造的提交 ID 一律 404。
"""
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from openpyxl import Workbook, load_workbook

from app import create_app


def _build_workbook(path):
    """两个独立指标：B1=A1*2（用 A 列），D1=C1+10（用 C 列），互不依赖。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = 3
    ws["B1"] = "=A1*2"          # 6
    ws["C1"] = 7
    ws["D1"] = "=C1+10"         # 17
    ws["A2"] = 1
    ws["B2"] = "=A2/0"          # 除零：复算必失败
    wb.save(path)


@pytest.fixture()
def env(tmp_path):
    src = tmp_path / "src.xlsx"
    _build_workbook(str(src))
    instance = tmp_path / "instance"
    app = create_app(instance_path=str(instance))
    client = app.test_client()
    with open(src, "rb") as fh:
        rv = client.post("/upload", data={"workbook": (fh, "src.xlsx")},
                         content_type="multipart/form-data")
    assert rv.status_code == 200
    return client, instance


def _analyze(client, addr, forbidden=""):
    rv = client.post("/analyze", data={
        "ind_sheet": ["S"], "ind_addr": [addr],
        "fb_sheet": ["S"], "fb_addr": [forbidden],
    }, follow_redirects=True)
    assert rv.status_code == 200
    sub_id = rv.request.path.split("/")[-1]
    return sub_id, rv


def _download(client, sub_id):
    rv = client.get(f"/download/{sub_id}")
    assert rv.status_code == 200
    assert rv.data[:2] == b"PK"
    return rv.data


def _xlsx_indicator_value(blob):
    wb = load_workbook(io.BytesIO(blob))
    ws = wb["S"]
    # 压缩后指标都在首行；读出全部公式/数值，复算用独立求值器
    from app.engine.formula import parse_formula
    from app.engine.evaluator import evaluate
    vals = {}

    def get_value(sheet, col, row):
        v = wb[sheet].cell(row=row, column=col).value
        if isinstance(v, str) and v.startswith("="):
            return evaluate(parse_formula(v, set(wb.sheetnames)).ast,
                            get_value, sheet)
        return v

    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                vals[cell.coordinate] = get_value("S", cell.column, cell.row)
    return vals


def test_sequential_analyses_keep_distinct_files(env):
    client, _instance = env
    # 第一次：指标 B1
    sub1, page1 = _analyze(client, "B1")
    assert "复算核对".encode() in page1.data
    assert "S!B1".encode() in page1.data
    blob1 = _download(client, sub1)
    vals1 = _xlsx_indicator_value(blob1)
    assert 6.0 in vals1.values()
    assert 17.0 not in vals1.values()       # 第一份不含 D1 的结果

    # 第二次：同一上传，改选指标 D1
    sub2, page2 = _analyze(client, "D1")
    assert sub1 != sub2
    assert "S!D1".encode() in page2.data
    blob2 = _download(client, sub2)
    vals2 = _xlsx_indicator_value(blob2)
    assert 17.0 in vals2.values()
    assert 6.0 not in vals2.values()

    # 旧结果页链接仍指向第一份文件，未被第二次覆盖
    page1_again = client.get(f"/result/{sub1}")
    assert page1_again.status_code == 200
    assert "S!B1".encode() in page1_again.data
    assert _download(client, sub1) == blob1
    # 两份文件确实不同
    assert blob1 != blob2


def test_interleaved_failure_does_not_delete_passed_file(env):
    client, _instance = env
    sub_ok1, _ = _analyze(client, "B1")
    blob_ok1 = _download(client, sub_ok1)

    # 交错提交：一次除零失败（复算不通过）
    sub_bad, page_bad = _analyze(client, "B2")
    assert "复算核对未通过".encode() in page_bad.data
    assert "除以零".encode() in page_bad.data
    # 失败提交无下载
    assert client.get(f"/download/{sub_bad}").status_code == 404
    # 失败提交的结果页可回看，且标明本次原因
    again = client.get(f"/result/{sub_bad}")
    assert again.status_code == 200
    assert "除以零".encode() in again.data

    # 已通过的另一份审计件仍可原样下载，未被删除/覆盖
    assert _download(client, sub_ok1) == blob_ok1

    # 再交错一次通过的提交，三者互不影响
    sub_ok2, _ = _analyze(client, "D1")
    assert _download(client, sub_ok1) == blob_ok1
    blob_ok2 = _download(client, sub_ok2)
    assert blob_ok2 != blob_ok1
    assert client.get(f"/download/{sub_bad}").status_code == 404


def test_trace_error_submission_has_no_file_and_isolates(env):
    client, _instance = env
    sub_ok, _ = _analyze(client, "B1")
    blob_ok = _download(client, sub_ok)
    # 追溯失败：指标格直接引用一个不存在的越界格不报错，这里用禁发触发
    sub_forbidden, page = _analyze(client, "B1", forbidden="A1:A1")
    assert "追溯未通过".encode() in page.data
    assert "禁发".encode() in page.data
    assert client.get(f"/download/{sub_forbidden}").status_code == 404
    # 已通过件不受影响
    assert _download(client, sub_ok) == blob_ok


def test_result_page_matches_download_content(env):
    client, _instance = env
    sub_id, page = _analyze(client, "B1")
    # 结果页展示的复算值
    assert b"6" in page.data
    blob = _download(client, sub_id)
    vals = _xlsx_indicator_value(blob)
    assert vals  # 非空
    assert all(v == 6.0 for v in vals.values())


def test_links_survive_service_restart(env):
    client, instance = env
    sub1, page1 = _analyze(client, "B1")
    sub2, _ = _analyze(client, "D1")
    blob1 = _download(client, sub1)
    blob2 = _download(client, sub2)

    # 模拟服务重启：同一 instance_path 重新构建应用（新进程视角）
    del client
    app2 = create_app(instance_path=str(instance))
    client2 = app2.test_client()

    rv1 = client2.get(f"/result/{sub1}")
    rv2 = client2.get(f"/result/{sub2}")
    assert rv1.status_code == 200 and "S!B1".encode() in rv1.data
    assert rv2.status_code == 200 and "S!D1".encode() in rv2.data
    assert client2.get(f"/download/{sub1}").data == blob1
    assert client2.get(f"/download/{sub2}").data == blob2


def test_forged_or_unknown_submission_ids_404(env):
    client, _instance = env
    for bad in ("deadbeef" * 4, "zzz", "../etc", "a" * 65, ""):
        rv = client.get(f"/download/{bad}")
        assert rv.status_code == 404
        rv = client.get(f"/result/{bad}")
        assert rv.status_code == 404
