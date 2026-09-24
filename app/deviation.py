"""同列公式偏离检查：同列相邻公式格的模板一致性审计。

直觉：同一列中上下相邻的公式格通常是「同一公式向下填充」的结果，例如
C2 的 ``=A2+$B$1`` 与 C3 的 ``=A3+$B$1`` 应属同一模板。本模块基于
``scan_formula()`` 的带字符区间扫描，把每个公式按其所在格归一为
**模板串**——引用改写为相对位移（R1C1 风格），绝对行 / 列、跨表限定、
字符串字面量、定义名称与其余非引用文本原样保留——再按
（工作表、列、相邻行区段）聚类比较：

- 只检查公式格数 >= ``MIN_SEGMENT_SIZE``（3）的区段；
- 某模板数量**严格超过一半**时，区段内其余模板的公式格被标记为偏离；
  平票或没有任何模板过半时整个区段不报；
- 结果给出异常坐标与原公式、主流（多数）公式示例与区段范围。
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict

from openpyxl.utils import column_index_from_string

from .formula_parser import RawRef, scan_formula
from .graph import display_id, split_node

# 区段内公式格数达到该值才参与偏离检查
MIN_SEGMENT_SIZE = 3

# 内部节点 id 的坐标部分（构图时统一大写、无 $）
_COORD_RE = re.compile(r"^([A-Z]{1,3})([0-9]{1,7})$")


def _dim(value: int, base: int, absolute: bool) -> str:
    """单个维度归一：绝对维保留原值，相对维改写为相对公式所在格的位移。"""
    return str(value) if absolute else f"[{value - base}]"


def _qualify(sheet: str) -> str:
    """跨表限定的规范写法：简单名不引，含空格 / 特殊字符加引号（与搬移模块一致）。"""
    if sheet.replace("_", "").isalnum():
        return f"{sheet}!"
    return "'{}'!".format(sheet.replace("'", "''"))


def _normalize_ref(ref: RawRef, col: int, row: int) -> str:
    """把一条引用归一为模板片段（R1C1 风格：绝对维为数值，相对维为 [位移]）。

    跨表限定保留；行 / 列的 ``$`` 锁定状态决定该维保留原值还是改写为位移，
    因此 ``=A2+$B$1`` 在 C2 与 ``=A3+$B$1`` 在 C3 归一结果相同。
    """
    if ref.kind == "cell":
        body = f"R{_dim(ref.row1, row, ref.row1_abs)}" \
               f"C{_dim(ref.col1, col, ref.col1_abs)}"
    elif ref.kind == "range":
        body = (
            f"R{_dim(ref.row1, row, ref.row1_abs)}"
            f"C{_dim(ref.col1, col, ref.col1_abs)}"
            f":R{_dim(ref.row2, row, ref.row2_abs)}"
            f"C{_dim(ref.col2, col, ref.col2_abs)}"
        )
    elif ref.kind == "cols":
        body = f"C{_dim(ref.col1, col, ref.col1_abs)}" \
               f":C{_dim(ref.col2, col, ref.col2_abs)}"
    else:  # rows
        body = f"R{_dim(ref.row1, row, ref.row1_abs)}" \
               f":R{_dim(ref.row2, row, ref.row2_abs)}"
    if ref.sheet is not None:
        return _qualify(ref.sheet) + body
    return body


def normalize_formula(formula: str, col: int, row: int) -> str:
    """把公式按所在格（col, row）归一为模板串。

    引用改写为相对位移；字符串字面量、函数名、定义名称、``#REF!`` 等
    非引用文本原样保留（利用 scan_formula 的字符区间逐字拷贝未命中片段）。
    """
    body = formula[1:] if formula.startswith("=") else formula
    prefix = "=" if formula.startswith("=") else ""
    parts: list[str] = []
    cursor = 0
    for sp in scan_formula(formula):
        parts.append(body[cursor:sp.start])
        parts.append(_normalize_ref(sp.ref, col, row) if sp.kind == "ref"
                     else sp.text)
        cursor = sp.end
    parts.append(body[cursor:])
    return prefix + "".join(parts)


def _column_segments(
    cells: list[tuple[int, str, str]],
) -> list[list[tuple[int, str, str]]]:
    """把一列中按行排序的公式格切分为相邻行区段（行号连续为同一段）。"""
    segments: list[list[tuple[int, str, str]]] = []
    current: list[tuple[int, str, str]] = []
    for cell in cells:
        if current and cell[0] != current[-1][0] + 1:
            segments.append(current)
            current = []
        current.append(cell)
    if current:
        segments.append(current)
    return segments


def _check_segment(
    sheet_name: str,
    col_letter: str,
    col_idx: int,
    segment: list[tuple[int, str, str]],
) -> dict | None:
    """检查单个相邻公式格区段，返回偏离结果（无多数模板或无偏离时返回 None）。"""
    templates = [normalize_formula(f, col_idx, row) for row, _, f in segment]
    counts = Counter(templates)
    majority_template, majority_count = counts.most_common(1)[0]
    # 严格过半才算主流；平票 / 无多数不报
    if majority_count * 2 <= len(segment):
        return None
    anomalies = [
        {
            "cell": node,
            "coord": f"{col_letter}{row}",
            "display": display_id(sheet_name, f"{col_letter}{row}"),
            "formula": formula,
            "template": template,
        }
        for (row, node, formula), template in zip(segment, templates)
        if template != majority_template
    ]
    if not anomalies:
        return None
    example_row, example_node, example_formula = next(
        cell for cell, template in zip(segment, templates)
        if template == majority_template
    )
    first_row, last_row = segment[0][0], segment[-1][0]
    return {
        "sheet": sheet_name,
        "column": col_letter,
        "range": f"{col_letter}{first_row}:{col_letter}{last_row}",
        "display": display_id(
            sheet_name, f"{col_letter}{first_row}:{col_letter}{last_row}"),
        "cell_count": len(segment),
        "template_count": len(counts),
        "majority": {
            "template": majority_template,
            "count": majority_count,
            "example_cell": example_node,
            "example_display": display_id(sheet_name,
                                          f"{col_letter}{example_row}"),
            "example_formula": example_formula,
        },
        "anomalies": anomalies,
    }


def detect_formula_deviations(graph, sheet: str | None = None) -> dict:
    """扫描工作簿公式，找出同列区段中偏离多数模板的公式格。

    返回汇总计数与逐区段结果（仅列出存在偏离的区段）；sheet 非空时
    只检查该工作表。结果按工作表、列、区段起始行排序，区段内异常按行排序。
    """
    # 按（工作表, 列）收集公式格：(行号, 节点 id, 公式)
    columns: dict[tuple[str, str], list[tuple[int, str, str]]] = defaultdict(list)
    for node, formula in graph.formulas.items():
        s, coord = split_node(node)
        if sheet is not None and s != sheet:
            continue
        m = _COORD_RE.match(coord)
        if m is None:
            continue
        columns[(s, m.group(1))].append((int(m.group(2)), node, formula))

    results: list[dict] = []
    segments_checked = 0
    anomaly_count = 0
    for (sheet_name, col_letter), cells in sorted(
            columns.items(),
            key=lambda kv: (kv[0][0], column_index_from_string(kv[0][1]))):
        cells.sort()
        col_idx = column_index_from_string(col_letter)
        for segment in _column_segments(cells):
            if len(segment) < MIN_SEGMENT_SIZE:
                continue
            segments_checked += 1
            found = _check_segment(sheet_name, col_letter, col_idx, segment)
            if found is not None:
                anomaly_count += len(found["anomalies"])
                results.append(found)

    return {
        "segments_checked": segments_checked,
        "segments_with_anomalies": len(results),
        "anomaly_count": anomaly_count,
        "min_segment_size": MIN_SEGMENT_SIZE,
        "filters": {"sheet": sheet},
        "segments": results,
    }
