"""公式区域搬移预演测试：平移语义、错误阻断、重叠快照与 TSV 接口。"""
import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook
from openpyxl.workbook.defined_name import DefinedName

from app.graph import WorkbookGraph
from app.main import app
from app.relocate import (
    MAX_RELOCATE_CELLS,
    RelocateInputError,
    parse_coord,
    parse_rect,
    preview_to_tsv,
    relocate_preview,
)

client = TestClient(app)


def build(wb: Workbook) -> tuple[WorkbookGraph, object]:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wbf = load_workbook(buf, data_only=False)
    wbv = load_workbook(buf, data_only=True)
    return WorkbookGraph.from_workbooks(wbf, wbv), wbf


def base_wb() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["B1"] = "=A1*2"                  # 纯相对
    ws["B2"] = "=$A$1+B1"               # 绝对 + 相对
    ws["B3"] = "=$A1+C$2"               # 混合
    ws["B4"] = "=SUM(A1:A2)"            # 区域
    ws["B5"] = "=SUM(B:B)+SUM(2:2)"     # 整列 / 整行
    ws["B6"] = "=S2!D5*2"               # 跨表
    ws["B7"] = "=本地名+全局名"          # 名称
    wb.create_sheet("S2")
    wb.defined_names["全局名"] = DefinedName("全局名", attr_text="S1!$A$1")
    ws.defined_names["本地名"] = DefinedName("本地名", attr_text="S1!$A$1")
    return wb


# ------------------------------------------------------------ 输入解析
def test_parse_coord_and_rect():
    assert (parse_coord("$B$3").col, parse_coord("B3").row) == (2, 3)
    (tl, br) = parse_rect("D5:B2")
    assert (tl.col, tl.row, br.col, br.row) == (2, 2, 4, 5)
    (tl, br) = parse_rect("B2")
    assert (tl.col, tl.row, br.col, br.row) == (2, 2, 2, 2)
    for bad in ("", "XFD", "A", "A0", "AAAA1", "B9999999"):
        try:
            parse_coord(bad) if bad != "" else parse_rect(bad)
        except RelocateInputError:
            pass
        else:
            raise AssertionError(f"应拒绝 {bad!r}")


# ------------------------------------------------------------ 同表平移
def test_same_sheet_relative_shift():
    g, wbf = build(base_wb())
    # B1 -> D5：+2 列 +4 行
    p = relocate_preview(g, wbf, "S1", "B1", "S1", "D5")
    e = p["entries"][0]
    assert e["original_formula"] == "=A1*2"
    assert e["translated_formula"] == "=C5*2"
    assert p["exportable"] and p["formula_count"] == 1 and p["non_formula_count"] == 0


def test_absolute_and_mixed_refs():
    g, wbf = build(base_wb())
    p = relocate_preview(g, wbf, "S1", "B2:B3", "S1", "D5")
    by_src = {e["source_coord"]: e for e in p["entries"]}
    # 源区左上角 B2 -> D5（+2 列 +3 行）；$A$1 全固定，B1 相对 -> D4
    assert by_src["B2"]["translated_formula"] == "=$A$1+D4"
    # B3 -> D6（+2/+3）：$A1 列锁定行 +3 -> $A4；C$2 行锁定列 +2 -> E$2
    assert by_src["B3"]["translated_formula"] == "=$A4+E$2"


def test_range_and_full_col_row():
    g, wbf = build(base_wb())
    p = relocate_preview(g, wbf, "S1", "B4:B5", "S1", "D5")
    by_src = {e["source_coord"]: e for e in p["entries"]}
    # B4 -> D5（+2/+1）
    assert by_src["B4"]["translated_formula"] == "=SUM(C2:C3)"
    # B5 -> D6（+2/+2）：B:B 列相对 +2 -> D:D；2:2 行相对 +2 -> 4:4
    # 注意源区为 B4:B5，B5 实际映射 D6
    assert by_src["B5"]["translated_formula"] == "=SUM(D:D)+SUM(3:3)"


def test_cross_sheet_ref_keeps_sheet_qualifier():
    g, wbf = build(base_wb())
    # B6 -> D5（+2 列 -1 行）：S2!D5 -> S2!F4
    p = relocate_preview(g, wbf, "S1", "B6", "S1", "D5")
    assert p["entries"][0]["translated_formula"] == "=S2!F4*2"


# ------------------------------------------------------------ 跨表搬移
def test_move_to_other_sheet_qualifies_local_refs():
    g, wbf = build(base_wb())
    # 从 S1 搬到 S2（D5，同样 +2 列 +4 行）：本表引用全部补 S1! 前缀
    p = relocate_preview(g, wbf, "S1", "B1", "S2", "D5")
    assert p["entries"][0]["translated_formula"] == "=S1!C5*2"

    # 从 S1 搬到 S2（B2 -> D5，+2 列 +3 行）：本表引用补 S1! 前缀，B1 -> D4
    p2 = relocate_preview(g, wbf, "S1", "B2", "S2", "D5")
    assert p2["entries"][0]["translated_formula"] == "=S1!$A$1+S1!D4"

    # 跨表引用 S2!D5（B6 -> D5，+2/-1）搬到 S2 后变成同表 -> F4 且去限定
    p3 = relocate_preview(g, wbf, "S1", "B6", "S2", "D5")
    assert p3["entries"][0]["translated_formula"] == "=F4*2"


def test_sheet_scoped_name_gets_qualified_when_cross_sheet():
    g, wbf = build(base_wb())
    # 同表：名称原样保留
    p = relocate_preview(g, wbf, "S1", "B7", "S1", "D5")
    assert p["entries"][0]["translated_formula"] == "=本地名+全局名"
    # 跨表：源表的工作表级名称改写为 S1!本地名，工作簿级名称不动
    p2 = relocate_preview(g, wbf, "S1", "B7", "S2", "D5")
    assert p2["entries"][0]["translated_formula"] == "=S1!本地名+全局名"


# ------------------------------------------------------------ 致命错误阻断
def fatal_wb() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["B1"] = "=#REF!+A1"            # 失效引用
    ws["B2"] = "=幽灵!A1"             # 不存在工作表
    ws["B3"] = "=A1048576"            # 下移 1 行即越界
    ws["B4"] = "=不存在的名称+1"       # 未定义名称
    return wb


def test_fatal_refs_make_preview_not_exportable():
    g, wbf = build(fatal_wb())
    p = relocate_preview(g, wbf, "S1", "B1:B4", "S1", "C2")
    assert p["exportable"] is False
    codes = {e["code"] for e in p["errors"]}
    assert codes == {"broken_ref", "missing_sheet", "out_of_bounds",
                     "undefined_name"}
    for entry in p["entries"]:
        assert entry["translated_formula"] is None and entry["errors"]
    # 不可导出时 TSV 直接拒绝
    try:
        preview_to_tsv(p)
    except RelocateInputError:
        pass
    else:
        raise AssertionError("不可导出的预演不应生成 TSV")


def test_column_overflow_is_fatal():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["XFD1"] = "=A1"  # 右移 1 列：A1 -> B1 合法，但目标区域止于 XFE 越界先在输入层拦截
    g, wbf = build(wb)
    try:
        relocate_preview(g, wbf, "S1", "XFD1", "S1", "XFE1")
    except RelocateInputError:
        pass
    else:
        raise AssertionError("目标区域越过 XFD 应报错")
    # 引用自身平移越界（源 A1，目标 XFD1：A1 -> XFD1 合法；改用 XFC1 格内引用 XFD1）
    ws["XFC1"] = "=XFD1"
    g, wbf = build(wb)
    p = relocate_preview(g, wbf, "S1", "XFC1", "S1", "XFD1")
    assert p["exportable"] is False
    assert p["errors"][0]["code"] == "out_of_bounds"


# ------------------------------------------------------------ 目标冲突与重叠快照
def test_target_conflict_and_overlap_snapshot():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=1+1"   # 公式
    ws["B1"] = 5        # 常量
    ws["D4"] = "旧内容"
    g, wbf = build(wb)

    # 非重叠：源 A1 -> 目标 D4（单格），D4 已有常量 -> occupied
    p = relocate_preview(g, wbf, "S1", "A1", "S1", "D4")
    assert p["conflict_count"] == 1 and p["occupied_count"] == 1
    c = p["conflicts"][0]
    assert c["type"] == "occupied" and c["target_coord"] == "D4"
    assert c["existing_kind"] == "value" and c["existing_text"] == "旧内容"

    # 重叠：源 A1:B1 -> B1:C1；B1 的旧值 5 必须来自源快照（公式条目行可见冲突）
    p2 = relocate_preview(g, wbf, "S1", "A1:B1", "S1", "B1")
    assert p2["overlap"] and p2["overlap_count"] == 1
    entry = next(e for e in p2["entries"] if e["source_coord"] == "A1")
    assert entry["conflict"] == "overlap_source"
    assert entry["target_existing_text"] == "5"
    # 快照语义：引擎不修改工作簿
    assert wbf["S1"]["A1"].value == "=1+1" and wbf["S1"]["B1"].value == 5


def test_non_formula_cells_skipped_but_conflicts_listed():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 10
    ws["A2"] = "=A1"
    ws["B2"] = "x"
    g, wbf = build(wb)
    p = relocate_preview(g, wbf, "S1", "A1:A2", "S1", "B1")
    assert p["formula_count"] == 1 and p["non_formula_count"] == 1
    assert p["total_cells"] == 2
    # 源 A1:A2 与目标 B1:B2 几何上不相交（不同列），B2 原有 "x" 属于
    # occupied 冲突（映射到 B2 的源格 A2 是公式，其条目带该冲突）
    p = relocate_preview(g, wbf, "S1", "A1:A2", "S1", "B1")
    assert p["formula_count"] == 1 and p["non_formula_count"] == 1
    assert p["total_cells"] == 2
    assert p["overlap"] is True and p["overlap_count"] == 0
    assert {c["target_coord"]: c["type"] for c in p["conflicts"]} == {"B2": "occupied"}
    entry = next(e for e in p["entries"] if e["source_coord"] == "A2")
    assert entry["conflict"] == "occupied"
    assert entry["target_existing_text"] == "x"


def test_out_of_used_range_is_notice_not_error():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=B2"
    g, wbf = build(wb)
    p = relocate_preview(g, wbf, "S1", "A1", "S1", "D10")
    assert p["exportable"] is True
    assert p["notice_count"] >= 1
    assert p["notices"][0]["code"] == "out_of_used_range"


def test_source_too_large_rejected():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    g, wbf = build(wb)
    try:
        relocate_preview(g, wbf, "S1", "A1:XFD100", "S1", "A1")
    except RelocateInputError as e:
        assert str(MAX_RELOCATE_CELLS) in str(e)
    else:
        raise AssertionError("超大区域应被拒绝")


# ------------------------------------------------------------ TSV
def test_tsv_content():
    g, wbf = build(base_wb())
    p = relocate_preview(g, wbf, "S1", "B1:B2", "S1", "D5")
    tsv = preview_to_tsv(p)
    assert tsv.startswith("﻿目标位置\t拟写入公式\n")
    lines = tsv.strip("\n").split("\n")
    # B1 -> D5（+2/+4）；B2 -> D6（+2/+4）
    assert lines[1] == "S1!D5\t=C5*2"
    assert lines[2] == "S1!D6\t=$A$1+D5"


# ------------------------------------------------------------ HTTP API
def upload(wb: Workbook) -> str:
    buf = io.BytesIO()
    wb.save(buf)
    r = client.post(
        "/api/upload",
        files={"file": ("t.xlsx", buf.getvalue(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
    )
    assert r.status_code == 200
    return r.json()["workbook_id"]


def test_relocate_preview_api_ok():
    wid = upload(base_wb())
    r = client.post(f"/api/workbooks/{wid}/relocate/preview", json={
        "source_sheet": "S1", "source_range": "B1:B3",
        "target_sheet": "S2", "target_anchor": "D5",
    })
    assert r.status_code == 200
    d = r.json()
    assert d["exportable"] is True and d["formula_count"] == 3
    assert d["delta"] == {"rows": 4, "cols": 2}
    first = next(e for e in d["entries"] if e["source_coord"] == "B1")
    assert first["translated_formula"] == "=S1!C5*2"


def test_relocate_preview_api_errors():
    wid = upload(fatal_wb())
    # 逐格错误：200 + exportable=false
    r = client.post(f"/api/workbooks/{wid}/relocate/preview", json={
        "source_sheet": "S1", "source_range": "B1:B4",
        "target_sheet": "S1", "target_anchor": "C2",
    })
    assert r.status_code == 200 and r.json()["exportable"] is False
    # TSV 被拒绝（400）
    r = client.post(f"/api/workbooks/{wid}/relocate/tsv", json={
        "source_sheet": "S1", "source_range": "B1:B4",
        "target_sheet": "S1", "target_anchor": "C2",
    })
    assert r.status_code == 400
    # 非法坐标 400；不存在工作表 404；不存在工作簿 404
    r = client.post(f"/api/workbooks/{wid}/relocate/preview", json={
        "source_sheet": "S1", "source_range": "BAD",
        "target_sheet": "S1", "target_anchor": "C2",
    })
    assert r.status_code == 400
    r = client.post(f"/api/workbooks/{wid}/relocate/preview", json={
        "source_sheet": "Nope", "source_range": "B1",
        "target_sheet": "S1", "target_anchor": "C2",
    })
    assert r.status_code == 404
    assert client.post("/api/workbooks/deadbeefab/relocate/preview", json={
        "source_sheet": "S1", "source_range": "B1",
        "target_sheet": "S1", "target_anchor": "C2",
    }).status_code == 404


def test_relocate_tsv_download():
    wid = upload(base_wb())
    r = client.post(f"/api/workbooks/{wid}/relocate/tsv", json={
        "source_sheet": "S1", "source_range": "B1",
        "target_sheet": "S1", "target_anchor": "D5",
    })
    assert r.status_code == 200
    assert "text/tab-separated-values" in r.headers["content-type"]
    assert "attachment" in r.headers["content-disposition"]
    assert "relocate_" in r.headers["content-disposition"]
    assert r.content.startswith("﻿目标位置".encode("utf-8"))
