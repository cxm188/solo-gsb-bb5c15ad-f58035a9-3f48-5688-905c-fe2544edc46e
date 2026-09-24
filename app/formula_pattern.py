"""同列公式模板归一化与相邻区段偏离检测。

把每条公式按**其所在格**归一为 R1C1 风格模板：相对引用写成相对该格的
行列位移（``R[0]C[-2]``），绝对引用保留绝对行列号（``R1C2``），混合引用
逐维度处理；跨表限定、整行 / 整列区域、定义名称、字符串与其他非引用文本
原样保留。于是同列向下复制的公式（如 C2 的 ``=A2+$B$1`` 与 C3 的
``=A3+$B$1``）归一结果相同，属于同一模板。

检测时按工作表、列把**行号相邻**的公式格切成区段，只检查至少
``MIN_SEGMENT_CELLS`` 格的区段；区段内某模板数量**严格超过一半**时，
其余模板的公式标记为偏离；平票（如 2:2）或无多数（多模板且无一过半）时
不报任何问题。
"""
from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass

from openpyxl.utils import coordinate_to_tuple, get_column_letter, quote_sheetname

from .formula_parser import RawRef, scan_formula
from .graph import display_id, split_node

# 只检查不少于该格数的相邻公式区段
MIN_SEGMENT_CELLS = 3


def _qualify(sheet: str, body: str) -> str:
    """跨表限定：简单名不引，含空格 / 特殊字符的表名按 Excel 写法加引号。"""
    if sheet.replace("_", "").isalnum():
        return f"{sheet}!{body}"
    return f"{quote_sheetname(sheet)}!{body}"


# ------------------------------------------------------------ 模板归一化
def _col_text(col: int, owner_col: int, abs_flag: bool) -> str:
    """列维度的 R1C1 文本：绝对 C{n}，相对 C[位移]（相对本列即 C[0]）。"""
    if abs_flag:
        return f"C{col}"
    return f"C[{col - owner_col}]"


def _row_text(row: int, owner_row: int, abs_flag: bool) -> str:
    """行维度的 R1C1 文本：绝对 R{n}，相对 R[位移]（相对本行即 R[0]）。"""
    if abs_flag:
        return f"R{row}"
    return f"R[{row - owner_row}]"


def _cell_text(
    col: int, row: int, col_abs: bool, row_abs: bool,
    owner_col: int, owner_row: int,
) -> str:
    # R1C1 约定行在前、列在后，如 R2C3 / R[0]C[-2]
    return _row_text(row, owner_row, row_abs) + _col_text(col, owner_col, col_abs)


def _render_ref(ref: RawRef, owner_col: int, owner_row: int) -> str:
    """把一条引用渲染为相对公式所在格归一化后的 R1C1 文本。

    跨表限定保留（简单表名不引、含特殊字符的表名加引号）；``$`` 绝对标记
    决定对应维度写绝对号还是相对位移。
    """
    if ref.kind == "cell":
        body = _cell_text(ref.col1, ref.row1, ref.col1_abs, ref.row1_abs,
                          owner_col, owner_row)
    elif ref.kind == "range":
        body = (
            _cell_text(ref.col1, ref.row1, ref.col1_abs, ref.row1_abs,
                       owner_col, owner_row)
            + ":"
            + _cell_text(ref.col2, ref.row2, ref.col2_abs, ref.row2_abs,
                         owner_col, owner_row)
        )
    elif ref.kind == "cols":
        body = (
            _col_text(ref.col1, owner_col, ref.col1_abs)
            + ":"
            + _col_text(ref.col2, owner_col, ref.col2_abs)
        )
    else:  # rows
        body = (
            _row_text(ref.row1, owner_row, ref.row1_abs)
            + ":"
            + _row_text(ref.row2, owner_row, ref.row2_abs)
        )
    if ref.sheet is not None:
        return _qualify(ref.sheet, body)
    return body


def normalize_formula(formula: str, owner_col: int, owner_row: int) -> str:
    """把公式按所在格 (owner_col, owner_row) 归一为 R1C1 模板字符串。

    引用片段按相对 / 绝对标记重写；定义名称（含跨表限定）、``#REF!`` 错误、
    字符串字面量、函数名、结构化引用、数字与运算符等全部原样保留。
    """
    body0 = formula[1:] if formula.startswith("=") else formula
    parts: list[str] = []
    cursor = 0
    for sp in scan_formula(formula):
        parts.append(body0[cursor:sp.start])
        if sp.kind == "ref":
            parts.append(_render_ref(sp.ref, owner_col, owner_row))
        else:
            # name（定义名称及跨表限定）/ error（#REF! 等）按原文保留
            parts.append(sp.text)
        cursor = sp.end
    parts.append(body0[cursor:])
    return "=" + "".join(parts)


# ------------------------------------------------------------ 区段分析
@dataclass(frozen=True)
class _FormulaCell:
    row: int
    node: str
    formula: str


def _split_runs(cells: list[_FormulaCell]) -> list[list[_FormulaCell]]:
    """把同列公式格按行号相邻切成连续区段（行号断开即新区段）。"""
    runs: list[list[_FormulaCell]] = []
    cur: list[_FormulaCell] = []
    for cell in cells:
        if cur and cell.row != cur[-1].row + 1:
            runs.append(cur)
            cur = []
        cur.append(cell)
    if cur:
        runs.append(cur)
    return runs


def analyze_deviations(
    formulas: dict[str, str],
    sheet_order: list[str],
    sheet: str | None = None,
) -> dict:
    """按工作表 / 列切相邻区段并统计同列公式模板偏离。

    返回包含 ``segments``（存在严格多数模板且有偏离格的区段）、各工作表
    汇总与筛选条件的字典。``sheet`` 非 None 时只分析 / 返回该表的区段，
    但 ``sheets`` 汇总始终覆盖整本工作簿（供页面筛选器展示各表计数）。
    """
    order_of = {name: i for i, name in enumerate(sheet_order)}
    # (工作表, 列号) -> 该列内的公式格
    columns: dict[tuple[str, int], list[_FormulaCell]] = defaultdict(list)
    for node, formula in formulas.items():
        owner_sheet, coord = split_node(node)
        row, col = coordinate_to_tuple(coord)
        columns[(owner_sheet, col)].append(_FormulaCell(row, node, formula))

    sheet_stats: dict[str, dict[str, int]] = {
        name: {"checked_segments": 0, "anomalous_segments": 0,
               "deviation_cells": 0}
        for name in sheet_order
    }
    segments: list[dict] = []

    def inspect_run(owner_sheet: str, col: int,
                    run: list[_FormulaCell]) -> None:
        stats = sheet_stats.setdefault(
            owner_sheet,
            {"checked_segments": 0, "anomalous_segments": 0,
             "deviation_cells": 0})
        stats["checked_segments"] += 1
        sized = len(run)
        patterns = [
            (normalize_formula(cell.formula, col, cell.row), cell)
            for cell in run
        ]
        counter = Counter(pattern for pattern, _ in patterns)
        majority, majority_count = counter.most_common(1)[0]
        # 严格超过一半才算多数；平票或多模板均不过半时整段不报
        if majority_count * 2 <= sized:
            return

        deviations = [
            {"cell": cell.node,
             "sheet": owner_sheet,
             "row": cell.row,
             "coord": split_node(cell.node)[1],
             "display": display_id(owner_sheet, split_node(cell.node)[1]),
             "formula": cell.formula,
             "pattern": pattern}
            for pattern, cell in patterns
            if pattern != majority
        ]
        deviations.sort(key=lambda d: d["row"])
        if not deviations:
            return

        example_cell = next(
            cell for pattern, cell in patterns if pattern == majority)
        pattern_counts = [
            {"pattern": pattern, "count": count,
             "example_cell": next(cell.node for pat, cell in patterns
                                  if pat == pattern),
             "example_formula": next(cell.formula for pat, cell in patterns
                                     if pat == pattern)}
            for pattern, count in counter.most_common()
        ]
        letter = get_column_letter(col)
        start_row, end_row = run[0].row, run[-1].row
        start_coord, end_coord = f"{letter}{start_row}", f"{letter}{end_row}"
        stats["anomalous_segments"] += 1
        stats["deviation_cells"] += len(deviations)
        # 带引号的表名同样包住区段范围，如 '月报 2026'!B2:B6
        prefix = display_id(owner_sheet, start_coord).rsplit("!", 1)[0]
        segments.append({
            "sheet": owner_sheet,
            "column": letter,
            "column_index": col,
            "segment": f"{start_coord}:{end_coord}",
            "segment_display": f"{prefix}!{start_coord}:{end_coord}",
            "segment_start": start_coord,
            "segment_end": end_coord,
            "start_row": start_row,
            "segment_rows": sized,
            "majority_pattern": majority,
            "majority_count": majority_count,
            "majority_example_cell": example_cell.node,
            "majority_example_display": display_id(
                owner_sheet, split_node(example_cell.node)[1]),
            "majority_example_formula": example_cell.formula,
            "deviation_count": len(deviations),
            "deviations": deviations,
            "pattern_counts": pattern_counts,
        })

    for (owner_sheet, col) in sorted(columns,
                                     key=lambda k: (order_of.get(k[0], 1 << 30),
                                                    k[1])):
        run_cells = sorted(columns[(owner_sheet, col)], key=lambda c: c.row)
        for run in _split_runs(run_cells):
            if len(run) < MIN_SEGMENT_CELLS:
                continue
            inspect_run(owner_sheet, col, run)

    segments.sort(key=lambda s: (order_of.get(s["sheet"], 1 << 30),
                                 s["column_index"], s["start_row"]))
    # 全局计数始终基于整本工作簿；筛选只影响返回的区段列表
    checked_total = sum(st["checked_segments"] for st in sheet_stats.values())
    anomalous_total = sum(st["anomalous_segments"] for st in sheet_stats.values())
    deviation_total = sum(st["deviation_cells"] for st in sheet_stats.values())
    visible = (
        [s for s in segments if s["sheet"] == sheet] if sheet is not None
        else segments
    )
    sheets = [
        {"name": name, **stats}
        for name, stats in sheet_stats.items()
        if stats["checked_segments"] or stats["anomalous_segments"]
    ]
    return {
        "segments": visible,
        "total": len(visible),
        "deviation_cell_count": sum(s["deviation_count"] for s in visible),
        "checked_segment_count": checked_total,
        "anomalous_segment_count": anomalous_total,
        "total_deviation_cell_count": deviation_total,
        "formula_cell_count": len(formulas),
        "min_segment_cells": MIN_SEGMENT_CELLS,
        "sheets": sheets,
        "filters": {"sheet": sheet},
    }
