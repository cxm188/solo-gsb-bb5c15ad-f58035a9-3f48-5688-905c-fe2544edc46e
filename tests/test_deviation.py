"""同列公式偏离检查测试：模板归一、区段划分、多数规则与 HTTP 接口。"""
import io

from fastapi.testclient import TestClient
from openpyxl import Workbook, load_workbook

from app.deviation import MIN_SEGMENT_SIZE, detect_formula_deviations, normalize_formula
from app.graph import WorkbookGraph
from app.main import app

client = TestClient(app)


def load_graph(wb: Workbook) -> WorkbookGraph:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    wbf = load_workbook(buf, data_only=False)
    buf.seek(0)
    wbv = load_workbook(buf, data_only=True)
    return WorkbookGraph.from_workbooks(wbf, wbv)


# ---------------------------------------------------------------- 模板归一
def test_normalize_relative_and_absolute():
    # C2 的 =A2+$B$1：A2 为相对引用（列 -2、行 0），$B$1 绝对保留
    assert normalize_formula("=A2+$B$1", 3, 2) == "=R[0]C[-2]+R1C2"
    # C3 的 =A3+$B$1 归一结果相同 -> 同一模板
    assert normalize_formula("=A3+$B$1", 3, 3) == "=R[0]C[-2]+R1C2"


def test_normalize_mixed_reference():
    # $A2 列绝对行相对；B$2 列相对行绝对
    assert normalize_formula("=$A2+B$2", 3, 2) == "=R[0]C1+R2C[-1]"


def test_normalize_range_cols_rows():
    assert normalize_formula("=SUM(A2:A4)", 3, 5) == "=SUM(R[-3]C[-2]:R[-1]C[-2])"
    assert normalize_formula("=SUM(B:B)", 3, 5) == "=SUM(C[-1]:C[-1])"
    assert normalize_formula("=SUM(2:2)", 3, 5) == "=SUM(R[-3]:R[-3])"
    assert normalize_formula("=$A$1:$C$3", 1, 1) == "=R1C1:R3C3"


def test_normalize_cross_sheet_qualifier_kept():
    assert normalize_formula("=参数!B2", 3, 2) == "=参数!R[0]C[-1]"
    # 带空格表名规范化为引号写法
    assert normalize_formula("='月报 2026'!B2", 3, 2) == "='月报 2026'!R[0]C[-1]"


def test_normalize_keeps_strings_names_and_text():
    # 字符串字面量、函数名、定义名称、错误字面量与非引用文本原样保留
    assert normalize_formula('=IF(A2="达标",B2,B3)', 4, 2) == '=IF(R[0]C[-3]="达标",R[0]C[-2],R[1]C[-2])'
    assert normalize_formula("=单价*A2", 3, 2) == "=单价*R[0]C[-2]"
    assert normalize_formula("=#REF!+A2", 3, 2) == "=#REF!+R[0]C[-2]"
    # 不同字符串 / 名称 -> 不同模板
    assert normalize_formula('=IF(A2="x",1,2)', 4, 2) != normalize_formula('=IF(A2="y",1,2)', 4, 3)


# ---------------------------------------------------------------- 区段检测
def build_deviation_wb() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "数据"
    ws["B1"] = 100
    for r in range(2, 9):                    # A2:A8 输入
        ws.cell(row=r, column=1, value=r)
    for r in range(2, 9):                    # C2:C8 同模板 =A{r}+$B$1
        ws.cell(row=r, column=3, value=f"=A{r}+$B$1")
    ws["C5"] = "=A4+$B$1"                    # 偏离：引用了上一行
    ws["E2"] = "=A2*2"                       # E2:E5 两种模板各半（平票，不报）
    ws["E3"] = "=A3*2"
    ws["E4"] = "=A4+1"
    ws["E5"] = "=A5+1"
    ws["G2"] = "=A2"                         # G2:G3 仅 2 格（不足 3 格，不检查）
    ws["G3"] = "=B1"
    return wb


def test_detects_deviating_formula():
    g = load_graph(build_deviation_wb())
    res = detect_formula_deviations(g)
    assert res["anomaly_count"] == 1
    assert res["segments_with_anomalies"] == 1
    seg = res["segments"][0]
    assert seg["sheet"] == "数据" and seg["column"] == "C"
    assert seg["range"] == "C2:C8"
    assert seg["display"] == "数据!C2:C8"
    assert seg["cell_count"] == 7
    assert seg["majority"]["count"] == 6
    assert seg["majority"]["example_formula"] == "=A2+$B$1"
    assert seg["majority"]["example_cell"] == "数据!C2"
    assert seg["majority"]["template"] == "=R[0]C[-2]+R1C2"
    assert len(seg["anomalies"]) == 1
    anom = seg["anomalies"][0]
    assert anom["cell"] == "数据!C5" and anom["coord"] == "C5"
    assert anom["formula"] == "=A4+$B$1"
    assert anom["template"] == "=R[-1]C[-2]+R1C2"


def test_tie_and_short_segments_not_reported():
    g = load_graph(build_deviation_wb())
    res = detect_formula_deviations(g)
    # E 列 2:2 平票、G 列不足 3 格 -> 都不出现在结果里
    assert {s["column"] for s in res["segments"]} == {"C"}
    # 但 E 列区段（4 格）计入了已检查区段数；G 列（2 格）不计
    assert res["segments_checked"] == 2
    assert res["min_segment_size"] == MIN_SEGMENT_SIZE == 3


def test_exactly_half_is_not_majority():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = "=B1"
    ws["A2"] = "=B2"
    ws["A3"] = "=C3"                         # 4 格 2:2 平票 -> 不报
    ws["A4"] = "=C4"
    g = load_graph(wb)
    assert detect_formula_deviations(g)["anomaly_count"] == 0
    # 3 格 2:1 -> 严格过半，报 1 个
    ws["A3"] = "=B3"
    ws["A4"] = "=C4"
    ws["A5"] = "=B5"
    g = load_graph(wb)
    res = detect_formula_deviations(g)
    assert res["anomaly_count"] == 1
    assert res["segments"][0]["anomalies"][0]["cell"] == "S1!A4"


def test_segments_split_by_row_gaps():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for r in (2, 3, 4):                      # C2:C4 一致
        ws.cell(row=r, column=3, value=f"=A{r}")
    for r in (6, 7, 8, 9):                   # C6:C9 含一个偏离
        ws.cell(row=r, column=3, value=f"=A{r}")
    ws["C8"] = "=B8"
    g = load_graph(wb)
    res = detect_formula_deviations(g)
    assert res["segments_checked"] == 2
    assert res["segments_with_anomalies"] == 1
    seg = res["segments"][0]
    assert seg["range"] == "C6:C9"
    assert [a["coord"] for a in seg["anomalies"]] == ["C8"]


def test_columns_checked_independently():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for r in (2, 3, 4):
        ws.cell(row=r, column=2, value=f"=A{r}*2")   # B 列模板 1
        ws.cell(row=r, column=4, value=f"=C{r}+1")   # D 列模板 2
    g = load_graph(wb)
    # 两列各自一致 -> 无偏离（若跨列混为一谈则会误报）
    assert detect_formula_deviations(g)["anomaly_count"] == 0


def test_absolute_reference_difference_detected():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    ws["A1"] = 1
    ws["A2"] = 2
    for r in (2, 3, 4, 5):
        ws.cell(row=r, column=3, value=f"=$A$1+B{r}")
    ws["C4"] = "=$A$2+B4"                    # 绝对行不同 -> 不同模板
    g = load_graph(wb)
    res = detect_formula_deviations(g)
    assert res["anomaly_count"] == 1
    assert res["segments"][0]["anomalies"][0]["cell"] == "S1!C4"


def test_sheet_filter():
    wb = build_deviation_wb()
    ws2 = wb.create_sheet("第二张")
    for r in (2, 3, 4, 5):
        ws2.cell(row=r, column=2, value=f"=A{r}")
    ws2["B4"] = "=A1"
    g = load_graph(wb)
    res = detect_formula_deviations(g, sheet="第二张")
    assert res["filters"]["sheet"] == "第二张"
    assert {s["sheet"] for s in res["segments"]} == {"第二张"}
    assert res["anomaly_count"] == 1
    assert res["segments"][0]["anomalies"][0]["cell"] == "第二张!B4"
    # 限定「数据」表时只报 C5
    res = detect_formula_deviations(g, sheet="数据")
    assert res["anomaly_count"] == 1
    assert res["segments"][0]["anomalies"][0]["cell"] == "数据!C5"


def test_clean_column_no_anomaly():
    wb = Workbook()
    ws = wb.active
    ws.title = "S1"
    for r in range(1, 6):
        ws.cell(row=r, column=2, value=f"=A{r}+$C$1")
    g = load_graph(wb)
    res = detect_formula_deviations(g)
    assert res["segments_checked"] == 1
    assert res["anomaly_count"] == 0
    assert res["segments"] == []


# ---------------------------------------------------------------- HTTP 接口
def make_deviation_xlsx() -> bytes:
    buf = io.BytesIO()
    build_deviation_wb().save(buf)
    return buf.getvalue()


def upload() -> str:
    r = client.post("/api/upload", files={"file": ("d.xlsx", make_deviation_xlsx(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert r.status_code == 200
    return r.json()["workbook_id"]


def test_deviation_endpoint():
    wid = upload()
    r = client.get(f"/api/workbooks/{wid}/deviation")
    assert r.status_code == 200
    d = r.json()
    assert d["workbook_id"] == wid
    assert d["anomaly_count"] == 1
    assert d["filters"] == {"sheet": None}
    seg = d["segments"][0]
    assert {"sheet", "column", "range", "display", "cell_count",
            "template_count", "majority", "anomalies"} <= set(seg)
    assert {"template", "count", "example_cell", "example_display",
            "example_formula"} <= set(seg["majority"])
    assert {"cell", "coord", "display", "formula", "template"} <= set(seg["anomalies"][0])


def test_deviation_endpoint_sheet_filter():
    wid = upload()
    r = client.get(f"/api/workbooks/{wid}/deviation", params={"sheet": "数据"})
    assert r.status_code == 200
    d = r.json()
    assert d["filters"]["sheet"] == "数据"
    assert all(s["sheet"] == "数据" for s in d["segments"])
    assert d["anomaly_count"] == 1
    # 筛到没有偏离的表
    r = client.get(f"/api/workbooks/{wid}/deviation", params={"sheet": "不存在"})
    assert r.status_code == 404


def test_deviation_missing_workbook_is_404():
    assert client.get("/api/workbooks/deadbeefab/deviation").status_code == 404
