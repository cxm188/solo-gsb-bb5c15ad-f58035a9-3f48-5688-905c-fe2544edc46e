"""公式区域搬移预演（不写工作簿，仅生成拟写入公式清单）。

用户选择源工作表的矩形区域、目标工作表与左上角坐标，按 **Excel 复制 /
粘贴语义**把区域内公式映射到目标位置：

- 相对引用（``A1``）随源格 -> 目标格的行列位移整体平移；
- 绝对引用（``$A$1``）固定不动；混合引用（``$A1`` / ``A$1``）只平移
  未锁定的维度；
- 跨表引用：表名前缀保持，表内坐标仍按上述规则平移；同表引用搬到另一
  张表时自动补上源表限定符，跨表引用搬回该表时按 Excel 习惯去掉冗余限定；
- 整行 / 整列引用（``B:B`` / ``2:5``）沿可平移维度平移，锁定维度不动；
- 定义名称按复制语义保留（名称属于工作簿，复制不改定义）；公式所在表的
  **工作表级名称**搬到别的表时改写为「源表!名称」限定写法，避免在目标表
  命中另一个同名本地名称；
- ``#REF!``、指向不存在工作表的引用 / 名称、平移后越过工作表物理边界
  （第 1048576 行 / 第 XFD 列）的引用属于**无法平移**：整次预演报错且
  不生成可导出 TSV。

预演严格基于**源区域与目标区域的一次性快照**计算：源 / 目标重叠时也只做
纯函数映射，绝不边读边改（应用本身也从不修改已上传工作簿）。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, replace

from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    quote_sheetname,
)

from .formula_parser import scan_formula
from .names import (
    EXTERNAL as NAME_EXTERNAL,
    INVALID as NAME_INVALID,
    UNDEFINED as NAME_UNDEFINED,
)
from .graph import display_id

# Excel 工作表物理边界
MAX_ROW = 1_048_576
MAX_COL = 16_384
# 单次预演允许的区域格子数上限（防止超大矩形把内存 / 响应打爆）
MAX_RELOCATE_CELLS = 50_000

_COORD_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


class RelocateInputError(ValueError):
    """请求参数非法（区域 / 坐标格式错误、区域过大等），映射为 HTTP 400。"""


class RelocateSheetError(LookupError):
    """指定的源 / 目标工作表不存在，映射为 HTTP 404。"""


@dataclass(frozen=True)
class _Point:
    col: int
    row: int


# ------------------------------------------------------------ 输入解析
def parse_coord(text: str) -> _Point:
    """解析 ``B3`` / ``$B$3`` 形态的单个坐标。"""
    m = _COORD_RE.match((text or "").strip().upper())
    if not m:
        raise RelocateInputError(f"坐标格式错误: {text!r}，应为如 B3 / $B$3")
    col = column_index_from_string(m.group(1))
    row = int(m.group(2))
    if not (1 <= col <= MAX_COL and 1 <= row <= MAX_ROW):
        raise RelocateInputError(f"坐标超出工作表边界: {text}")
    return _Point(col, row)


def parse_rect(text: str) -> tuple[_Point, _Point]:
    """解析单格（``B3``）或矩形区域（``B2:D5``，允许带 $ 与反向书写）。"""
    raw = (text or "").strip()
    if not raw:
        raise RelocateInputError("源区域不能为空，形如 B2:D5")
    parts = raw.split(":")
    if len(parts) > 2:
        raise RelocateInputError(f"源区域格式错误: {text!r}，应为单格或 B2:D5")
    pts = [parse_coord(p) for p in parts]
    p1, p2 = pts[0], pts[-1]
    return (
        _Point(min(p1.col, p2.col), min(p1.row, p2.row)),
        _Point(max(p1.col, p2.col), max(p1.row, p2.row)),
    )


def coord_text(col: int, row: int) -> str:
    return f"{get_column_letter(col)}{row}"


# ------------------------------------------------------------ 单元格快照
def _describe_value(value) -> dict | None:
    """把 openpyxl 单元格值描述为 {kind, text}；空单元格返回 None。"""
    if value is None:
        return None
    if isinstance(value, str) and value.startswith("="):
        text = value
        kind = "formula"
    elif isinstance(value, bool):
        kind, text = "value", "TRUE" if value else "FALSE"
    elif hasattr(value, "isoformat") and not isinstance(value, (int, float, str)):
        kind, text = "value", value.isoformat(sep=" ")
    else:
        kind, text = "value", str(value)
    if len(text) > 200:
        text = text[:200] + "…"
    return {"kind": kind, "text": text}


def _snapshot(ws, r1: int, r2: int, c1: int, c2: int) -> dict[tuple[int, int], dict]:
    """一次性读取矩形区域内的非空单元格内容（基于读取时刻的快照）。"""
    snap: dict[tuple[int, int], dict] = {}
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=c1, max_col=c2):
        for cell in row:
            desc = _describe_value(cell.value)
            if desc is not None:
                snap[(cell.column, cell.row)] = desc
    return snap


# ------------------------------------------------------------ 公式平移
def _qualify(sheet: str, body: str) -> str:
    """公式中的工作表限定：简单名不引，含空格 / 特殊字符才加引号（Excel 写法）。"""
    if sheet.replace("_", "").isalnum():
        return f"{sheet}!{body}"
    return f"{quote_sheetname(sheet)}!{body}"


def _shift_dim(value: int, delta: int, locked: bool) -> int:
    return value if locked else value + delta


def _render_cell(col: int, row: int, col_abs: bool, row_abs: bool) -> str:
    return f"{'$' if col_abs else ''}{get_column_letter(col)}" \
           f"{'$' if row_abs else ''}{row}"


def _render_ref(ref, effective_sheet: str, target_sheet: str) -> str:
    """按平移后的坐标与原始 $ 标记重新拼出引用文本。"""
    if ref.kind == "cell":
        body = _render_cell(ref.col1, ref.row1, ref.col1_abs, ref.row1_abs)
    elif ref.kind == "range":
        body = (
            _render_cell(ref.col1, ref.row1, ref.col1_abs, ref.row1_abs)
            + ":"
            + _render_cell(ref.col2, ref.row2, ref.col2_abs, ref.row2_abs)
        )
    elif ref.kind == "cols":
        body = f"{'$' if ref.col1_abs else ''}{get_column_letter(ref.col1)}:" \
               f"{'$' if ref.col2_abs else ''}{get_column_letter(ref.col2)}"
    else:  # rows
        body = f"{'$' if ref.row1_abs else ''}{ref.row1}:" \
               f"{'$' if ref.row2_abs else ''}{ref.row2}"
    if effective_sheet != target_sheet:
        return _qualify(effective_sheet, body)
    return body


def _translate_formula(
    formula: str,
    source_sheet: str,
    source_cell: _Point,
    target_sheet: str,
    target_cell: _Point,
    sheet_names: set[str],
    used_bounds: dict[str, tuple[int, int]],
    name_resolver,
) -> tuple[str | None, list[dict], list[dict]]:
    """把一条公式按 Excel 复制语义平移到目标单元格。

    返回 (拟写入公式, 错误列表, 提示列表)。出现任何错误时拟写入公式为
    None（整次预演不可导出）。
    """
    errors: list[dict] = []
    notices: list[dict] = []
    delta_col = target_cell.col - source_cell.col
    delta_row = target_cell.row - source_cell.row

    def err(code: str, message: str, ref: str) -> None:
        errors.append({"code": code, "message": message, "ref": ref})

    def notice(code: str, message: str, ref: str) -> None:
        notices.append({"code": code, "message": message, "ref": ref})

    body0 = formula[1:] if formula.startswith("=") else formula
    prefix = "=" if formula.startswith("=") else ""
    out_parts: list[str] = []
    cursor = 0
    for sp in scan_formula(formula):
        out_parts.append(body0[cursor:sp.start])
        replacement = sp.text

        if sp.kind == "error":
            err("broken_ref",
                f"失效引用 {sp.text} 无法随搬移调整（Excel 中即代表引用已丢失）",
                sp.text)
        elif sp.kind == "ref":
            ref = sp.ref
            eff_sheet = ref.sheet if ref.sheet is not None else source_sheet
            if eff_sheet not in sheet_names:
                err("missing_sheet",
                    f"引用指向不存在的工作表 {eff_sheet!r}（{sp.text}），无法平移",
                    sp.text)
            else:
                # 仅未锁定（相对）的维度参与平移
                nc1 = _shift_dim(ref.col1, delta_col, ref.col1_abs) \
                    if ref.col1 is not None else None
                nr1 = _shift_dim(ref.row1, delta_row, ref.row1_abs) \
                    if ref.row1 is not None else None
                nc2 = _shift_dim(ref.col2, delta_col, ref.col2_abs) \
                    if ref.col2 is not None else None
                nr2 = _shift_dim(ref.row2, delta_row, ref.row2_abs) \
                    if ref.row2 is not None else None
                # 越界判定：列向坐标 <= MAX_COL、行向坐标 <= MAX_ROW
                col_vals = [v for v in (nc1, nc2) if v is not None]
                row_vals = [v for v in (nr1, nr2) if v is not None]
                if any(not 1 <= v <= MAX_COL for v in col_vals) or \
                        any(not 1 <= v <= MAX_ROW for v in row_vals):
                    err("out_of_bounds",
                        f"引用 {sp.text} 平移后越过工作表边界"
                        f"（最大 {get_column_letter(MAX_COL)}{MAX_ROW}），无法平移",
                        sp.text)
                else:
                    nref = replace(ref, sheet=None, col1=nc1, row1=nr1,
                                   col2=nc2, row2=nr2)
                    replacement = _render_ref(nref, eff_sheet, target_sheet)
                    # 平移后落在目标表已用范围之外：不阻断（Excel 允许），仅提示
                    bounds = used_bounds.get(eff_sheet)
                    if bounds is not None and nref.kind in ("cell", "range"):
                        max_r, max_c = bounds
                        if nr2 > max_r or nc2 > max_c:
                            notices.append({
                                "code": "out_of_used_range",
                                "message": f"引用 {sp.text} 平移为 {replacement}，"
                                           f"超出工作表「{eff_sheet}」当前已用范围"
                                           f"（{max_r} 行 × {max_c} 列），粘贴后可能为空",
                                "ref": sp.text,
                            })
        else:  # name
            replacement = _translate_name(
                sp, source_sheet, target_sheet, sheet_names,
                name_resolver, err, notice)

        out_parts.append(replacement)
        cursor = sp.end
    out_parts.append(body0[cursor:])

    translated = None if errors else prefix + "".join(out_parts)
    return translated, errors, notices


def _translate_name(
    sp,
    source_sheet: str,
    target_sheet: str,
    sheet_names: set[str],
    name_resolver,
    err,
    notice,
) -> str:
    """按复制语义处理定义名称片段，返回替换文本（出错时原样返回并记错误）。"""
    rn = sp.name
    if rn.qualifier is not None and rn.qualifier not in sheet_names:
        err("missing_sheet",
            f"名称限定的工作表 {rn.qualifier!r} 不存在（{sp.text}），无法平移",
            sp.text)
        return sp.text
    if name_resolver is None:
        return sp.text
    res = name_resolver.resolve(rn, source_sheet)
    if res.status == NAME_UNDEFINED:
        err("undefined_name",
            f"定义名称 {sp.text} 未在对应作用域定义，搬移后语义无法保证",
            sp.text)
        return sp.text
    if res.status == NAME_INVALID:
        err("invalid_name",
            f"定义名称 {sp.text} 的目标已失效 / 存在名称循环，无法平移",
            sp.text)
        return sp.text
    if res.status == NAME_EXTERNAL:
        notice("external_name",
               f"定义名称 {sp.text} 指向外部工作簿，搬移后保持原样（不随区域调整）",
               sp.text)
        return sp.text
    # resolved / constant / formula：跨表限定名称原样保留；仅当无限定名称
    # 命中**源表的工作表级名称**、且公式被搬到另一张表时，补上源表限定符，
    # 以免在目标表解析到另一个同名本地名称（Excel 复制语义）。
    if rn.qualifier is None and res.scope == source_sheet and \
            target_sheet != source_sheet:
        return _qualify(source_sheet, rn.name)
    return sp.text


# ------------------------------------------------------------ 预演主流程
def relocate_preview(graph, wb_formula, source_sheet: str, source_range: str,
                     target_sheet: str, target_anchor: str) -> dict:
    """生成公式区域搬移预演结果（纯函数，不修改工作簿）。"""
    sheet_names = {s.name for s in graph.sheets}
    if source_sheet not in sheet_names:
        raise RelocateSheetError(f"源工作表不存在: {source_sheet}")
    if target_sheet not in sheet_names:
        raise RelocateSheetError(f"目标工作表不存在: {target_sheet}")

    top_left, bottom_right = parse_rect(source_range)
    anchor = parse_coord(target_anchor)
    rows = bottom_right.row - top_left.row + 1
    cols = bottom_right.col - top_left.col + 1
    if rows * cols > MAX_RELOCATE_CELLS:
        raise RelocateInputError(
            f"源区域过大：{rows} × {cols} = {rows * cols} 格，"
            f"单次预演上限 {MAX_RELOCATE_CELLS} 格")

    t_r2, t_c2 = anchor.row + rows - 1, anchor.col + cols - 1
    if t_r2 > MAX_ROW or t_c2 > MAX_COL:
        raise RelocateInputError(
            f"目标区域越过工作表边界（止于 {coord_text(t_c2, t_r2)}，"
            f"最大 {get_column_letter(MAX_COL)}{MAX_ROW}）")

    ws_src = wb_formula[source_sheet]
    ws_tgt = wb_formula[target_sheet]
    # 一次性快照：后续全部计算只使用快照，重叠也不回读活动状态
    src_snap = _snapshot(ws_src, top_left.row, bottom_right.row,
                         top_left.col, bottom_right.col)
    tgt_snap = _snapshot(ws_tgt, anchor.row, t_r2, anchor.col, t_c2)

    overlap = source_sheet == target_sheet
    overlap_src_at: dict[tuple[int, int], tuple[int, int]] = {}
    if overlap:
        # 目标格与源格在几何上重合时，粘贴前内容即该位置上的源格快照值
        # （一次性快照，不读取「被改写后」的活动状态）
        for r in range(anchor.row, t_r2 + 1):
            for c in range(anchor.col, t_c2 + 1):
                if top_left.row <= r <= bottom_right.row and \
                        top_left.col <= c <= bottom_right.col:
                    overlap_src_at[(c, r)] = (c, r)

    used_bounds = {s.name: (s.max_row, s.max_col) for s in graph.sheets}
    name_resolver = graph.name_resolver

    entries: list[dict] = []
    conflicts: list[dict] = []
    all_errors: list[dict] = []
    all_notices: list[dict] = []
    formula_count = 0

    for i in range(rows):
        for j in range(cols):
            sr, sc = top_left.row + i, top_left.col + j
            tr, tc = anchor.row + i, anchor.col + j
            s_coord = coord_text(sc, sr)
            t_coord = coord_text(tc, tr)
            src_node = f"{source_sheet}!{s_coord}"
            formula = graph.formulas.get(src_node)

            existing = None
            geo_src = overlap_src_at.get((tc, tr))
            is_overlap = geo_src is not None
            if is_overlap:
                gsc, gsr = geo_src
                existing = src_snap.get((gsc, gsr))
            else:
                existing = tgt_snap.get((tc, tr))

            conflict_type = None
            if existing is not None:
                conflict_type = "overlap_source" if is_overlap else "occupied"
                conflicts.append({
                    "type": conflict_type,
                    "target_sheet": target_sheet,
                    "target_coord": t_coord,
                    "target_display": display_id(target_sheet, t_coord),
                    "existing_kind": existing["kind"],
                    "existing_text": existing["text"],
                    "overlap": is_overlap,
                    "source_sheet": source_sheet if is_overlap else None,
                    "source_coord": coord_text(gsc, gsr) if is_overlap else None,
                    "source_display": (display_id(source_sheet, coord_text(gsc, gsr))
                                       if is_overlap else None),
                })

            if formula is None:
                continue
            formula_count += 1
            translated, errs, notes = _translate_formula(
                formula, source_sheet, _Point(sc, sr),
                target_sheet, _Point(tc, tr),
                sheet_names, used_bounds, name_resolver,
            )
            cell_errors = [{**e, "cell": src_node,
                            "cell_display": display_id(source_sheet, s_coord)}
                           for e in errs]
            cell_notices = [{**n, "cell": src_node,
                             "cell_display": display_id(source_sheet, s_coord)}
                            for n in notes]
            all_errors.extend(cell_errors)
            all_notices.extend(cell_notices)
            entries.append({
                "source_sheet": source_sheet,
                "source_coord": s_coord,
                "source_display": display_id(source_sheet, s_coord),
                "target_sheet": target_sheet,
                "target_coord": t_coord,
                "target_display": display_id(target_sheet, t_coord),
                "original_formula": formula,
                "translated_formula": translated,
                "errors": cell_errors,
                "notices": cell_notices,
                "conflict": conflict_type,
                "target_existing_kind": existing["kind"] if existing else None,
                "target_existing_text": existing["text"] if existing else None,
            })

    occupied_count = sum(1 for c in conflicts if not c["overlap"])
    overlap_count = len(conflicts) - occupied_count
    return {
        "source": {
            "sheet": source_sheet,
            "range": f"{coord_text(top_left.col, top_left.row)}:"
                     f"{coord_text(bottom_right.col, bottom_right.row)}",
            "coord1": coord_text(top_left.col, top_left.row),
            "coord2": coord_text(bottom_right.col, bottom_right.row),
            "rows": rows,
            "cols": cols,
        },
        "target": {
            "sheet": target_sheet,
            "anchor": coord_text(anchor.col, anchor.row),
            "range": f"{coord_text(anchor.col, anchor.row)}:"
                     f"{coord_text(t_c2, t_r2)}",
            "coord1": coord_text(anchor.col, anchor.row),
            "coord2": coord_text(t_c2, t_r2),
            "rows": rows,
            "cols": cols,
        },
        "delta": {"rows": anchor.row - top_left.row,
                  "cols": anchor.col - top_left.col},
        "overlap": overlap,
        "total_cells": rows * cols,
        "formula_count": formula_count,
        "non_formula_count": rows * cols - formula_count,
        "exportable": not all_errors,
        "error_count": len(all_errors),
        "notice_count": len(all_notices),
        "conflict_count": len(conflicts),
        "occupied_count": occupied_count,
        "overlap_count": overlap_count,
        "entries": entries,
        "conflicts": conflicts,
        "errors": all_errors,
        "notices": all_notices,
    }


def preview_to_tsv(preview: dict) -> str:
    """把可导出的预演结果渲染为 TSV（目标位置 + 拟写入公式）。"""
    if not preview["exportable"]:
        raise RelocateInputError("存在无法平移的引用，本次预演不可导出")

    def clean(s: str) -> str:
        return s.replace("\t", " ").replace("\r", " ").replace("\n", " ")

    lines = ["目标位置\t拟写入公式"]
    for e in preview["entries"]:
        location = _qualify(e["target_sheet"], e["target_coord"])
        lines.append(f"{clean(location)}\t{clean(e['translated_formula'])}")
    # UTF-8 BOM，便于 Excel 直接打开识别编码
    return "﻿" + "\n".join(lines) + "\n"
