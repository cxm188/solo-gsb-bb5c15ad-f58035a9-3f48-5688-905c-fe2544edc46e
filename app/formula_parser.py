"""Excel 公式引用解析。

从公式文本中抽取单元格引用 / 矩形区域引用（含跨工作表、带引号工作表名、
整行/整列区域）以及**定义名称**（名称管理器中的名称，含跨表限定
``Sheet!Name``），输出未展开的 RawRef / RawName 列表，由 graph 模块结合
工作表实际尺寸与定义名称表展开为具体单元格。
"""
from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string

# 工作表名前缀：'月报 2026'! 或 Sheet1! 或 参数!
_SHEET = r"(?:'(?P<qsheet>[^']+(?:''[^']*)*)'|(?P<usheet>[A-Za-z0-9_.一-鿿]+))!"
# 限定名称处的第二份表名前缀（命名组不能重名）
_SHEET2 = r"(?:'(?P<qsheet2>[^']+(?:''[^']*)*)'|(?P<usheet2>[A-Za-z0-9_.一-鿿]+))!"
_CELL = r"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}"
_COL = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?[0-9]{1,7}"
# 名称主体：首字符为字母 / 下划线 / 反斜杠 / CJK，后续可含 . 等；不含 $
_NAME_HEAD = r"(?:[A-Za-z_\\一-鿿][A-Za-z0-9_.一-鿿]*)"

_TOKEN_RE = re.compile(
    r"""
      (?P<string>"(?:[^"]|"")*")                        # 字符串字面量，忽略
    | (?P<error>\#[A-Za-z0-9!/?]+)                      # #REF! 等错误字面量，忽略
    | (?P<sheetref>{sheet}(?:{cell}(?::{cell})?|{col}:{col}|{row}:{row})(?![A-Za-z0-9_.一-鿿$]))
    | (?P<sheetname>{sheet2}{name})                     # 跨表限定的定义名称
    | (?P<structured>{name}?\[(?:[^\[\]]|\[[^\[\]]*\])*\])  # 结构化引用 表[列] / [[#头],[列]]
    | (?P<range>(?<![0-9.]){cell}(?![A-Za-z0-9_.一-鿿$]):{cell}(?![A-Za-z0-9_.一-鿿$])
                 |{col}(?![0-9$]):{col}(?![0-9$])|{row}(?![A-Za-z$]):{row}(?![A-Za-z$]))
    | (?P<func>[A-Za-z_\\一-鿿][A-Za-z0-9_.一-鿿]*(?=\s*\())  # 函数名，忽略
    | (?P<cell>(?<![0-9.]){cell})(?![A-Za-z0-9_.一-鿿$])     # 本表单单元格（排除 1E2 科学计数法）
    | (?P<name>(?<![0-9.]){name})(?!\s*\()              # 本表定义名称（不含函数调用）
    | (?P<other>.)
    """.format(sheet=_SHEET, sheet2=_SHEET2, cell=_CELL, col=_COL, row=_ROW, name=_NAME_HEAD),
    re.VERBOSE,
)
_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$")


@dataclass(frozen=True)
class RawRef:
    """一条尚未展开的单元格引用。行/列为 None 表示整列/整行区域，需按工作表尺寸截断。

    各 ``*_abs`` 标记记录引用对应端是否带 ``$``（绝对引用），供公式搬移时按
    Excel 复制语义决定该端是否随位移平移。
    """

    sheet: str | None  # None 表示公式所在工作表
    kind: str  # "cell" | "range" | "cols" | "rows"
    col1: int | None
    row1: int | None
    col2: int | None
    row2: int | None
    text: str  # 原始文本，用于展示
    col1_abs: bool = False
    row1_abs: bool = False
    col2_abs: bool = False
    row2_abs: bool = False


@dataclass(frozen=True)
class RawName:
    """公式中出现的定义名称（名称管理器名称）。

    qualifier 非 None 表示跨表限定，如 ``计算!单价``；为 None 时按公式所在
    工作表先查工作表级名称、再查工作簿级名称。
    """

    name: str
    qualifier: str | None
    text: str


@dataclass(frozen=True)
class RawToken:
    """扫描得到的单个 token：引用或名称。"""

    kind: str  # "ref" | "name"
    ref: RawRef | None = None
    name: RawName | None = None


@dataclass(frozen=True)
class RawSpan:
    """公式扫描命中片段：类型、原文与其在公式体（去掉开头 ``=``）中的字符区间。

    kind 除 ref / name 外还有 ``error``（``#REF!`` 等错误字面量，搬移时需
    拦截报错）；字符串、函数名、结构化引用等不产生 span（按原文保留）。
    """

    kind: str  # "ref" | "name" | "error"
    start: int
    end: int
    text: str
    ref: RawRef | None = None
    name: RawName | None = None


# 布尔等保留字不是定义名称，公式中直接跳过
_RESERVED_NAMES = {"TRUE", "FALSE"}

# 带绝对引用标记的单元格片段：分组 1/3 为 $（可能为 None），2/4 为列与行
_CELL_PARTS_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]{1,7})$")
_COL_PARTS_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
_ROW_PARTS_RE = re.compile(r"^(\$?)([0-9]{1,7})$")


def _split_cell(token: str) -> tuple[int, int]:
    m = _CELL_RE.match(token)
    if not m:
        raise ValueError(f"非法单元格引用: {token!r}")
    return column_index_from_string(m.group(1).upper()), int(m.group(2))


def _split_cell_parts(token: str) -> tuple[int, int, bool, bool]:
    m = _CELL_PARTS_RE.match(token)
    if not m:
        raise ValueError(f"非法单元格引用: {token!r}")
    return (column_index_from_string(m.group(2).upper()), int(m.group(4)),
            bool(m.group(1)), bool(m.group(3)))


def _split_col(token: str) -> int:
    return column_index_from_string(token.replace("$", "").upper())


def _split_row(token: str) -> int:
    return int(token.replace("$", ""))


def _build_ref(sheet: str | None, body: str, full_text: str) -> RawRef | None:
    """把引用主体（不含工作表前缀）解析为 RawRef（保留 $ 绝对引用标记）。"""
    if ":" in body:
        left, right = body.split(":", 1)
        if _CELL_RE.match(left) and _CELL_RE.match(right):
            c1, r1, c1a, r1a = _split_cell_parts(left)
            c2, r2, c2a, r2a = _split_cell_parts(right)
            # 允许反向书写（B2:A1）：归并时 $ 标记跟随其原本所属的端
            lo_c, hi_c = (c1, c1a) if c1 <= c2 else (c2, c2a), \
                         (c1, c1a) if c1 > c2 else (c2, c2a)
            lo_r, hi_r = (r1, r1a) if r1 <= r2 else (r2, r2a), \
                         (r1, r1a) if r1 > r2 else (r2, r2a)
            return RawRef(sheet, "range", lo_c[0], lo_r[0], hi_c[0], hi_r[0],
                          full_text, lo_c[1], lo_r[1], hi_c[1], hi_r[1])
        if re.fullmatch(_COL, left) and re.fullmatch(_COL, right):
            m1, m2 = _COL_PARTS_RE.match(left), _COL_PARTS_RE.match(right)
            c1, c2 = _split_col(left), _split_col(right)
            (lo, loa), (hi, hia) = ((c1, bool(m1.group(1))), (c2, bool(m2.group(1)))) \
                if c1 <= c2 else ((c2, bool(m2.group(1))), (c1, bool(m1.group(1))))
            return RawRef(sheet, "cols", lo, None, hi, None, full_text,
                          loa, False, hia, False)
        if re.fullmatch(_ROW, left) and re.fullmatch(_ROW, right):
            m1, m2 = _ROW_PARTS_RE.match(left), _ROW_PARTS_RE.match(right)
            r1, r2 = _split_row(left), _split_row(right)
            (lo, loa), (hi, hia) = ((r1, bool(m1.group(1))), (r2, bool(m2.group(1)))) \
                if r1 <= r2 else ((r2, bool(m2.group(1))), (r1, bool(m1.group(1))))
            return RawRef(sheet, "rows", None, lo, None, hi, full_text,
                          False, loa, False, hia)
        return None
    if _CELL_RE.match(body):
        c, r, ca, ra = _split_cell_parts(body)
        return RawRef(sheet, "cell", c, r, c, r, full_text, ca, ra, ca, ra)
    return None


def _sheet_of(m: re.Match) -> str | None:
    sheet = m.group("qsheet")
    if sheet is not None:
        return sheet.replace("''", "'")
    return m.group("usheet")


def scan_formula(formula: str) -> list[RawSpan]:
    """词法扫描公式，返回带字符区间的命中片段（ref / name / error）。

    区间相对去掉开头 ``=`` 后的公式体；字符串字面量、函数名、结构化引用等
    不产生 span，调用方可据区间把公式重写为搬移后的文本。
    """
    if not formula:
        return []
    text = formula[1:] if formula.startswith("=") else formula
    spans: list[RawSpan] = []
    for m in _TOKEN_RE.finditer(text):
        kind = m.lastgroup
        s, e, matched = m.start(), m.end(), m.group(kind)
        if kind == "sheetref":
            sheet = _sheet_of(m)
            body = matched[matched.index("!") + 1:]
            ref = _build_ref(sheet, body, matched)
            if ref is not None:
                spans.append(RawSpan("ref", s, e, matched, ref=ref))
        elif kind == "sheetname":
            sheet = m.group("qsheet2")
            if sheet is not None:
                sheet = sheet.replace("''", "'")
            else:
                sheet = m.group("usheet2")
            name = matched.rsplit("!", 1)[-1]
            spans.append(RawSpan("name", s, e, matched,
                                 name=RawName(name, sheet, matched)))
        elif kind == "range":
            ref = _build_ref(None, matched, matched)
            if ref is not None:
                spans.append(RawSpan("ref", s, e, matched, ref=ref))
        elif kind == "cell":
            spans.append(RawSpan("ref", s, e, matched,
                                 ref=_build_ref(None, matched, matched)))
        elif kind == "error":
            spans.append(RawSpan("error", s, e, matched))
        elif kind == "name":
            if matched.upper() in _RESERVED_NAMES:
                continue
            spans.append(RawSpan("name", s, e, matched,
                                 name=RawName(matched, None, matched)))
        else:
            continue
    return spans


def tokenize_formula(formula: str) -> list[RawToken]:
    """词法扫描公式，返回引用 / 名称 token（字符串、函数名等忽略）。"""
    return [
        RawToken(kind=sp.kind, ref=sp.ref, name=sp.name)
        for sp in scan_formula(formula)
        if sp.kind in ("ref", "name")
    ]


def parse_formula(formula: str) -> list[RawRef]:
    """解析公式文本，返回单元格引用列表（未展开，定义名称不在其中）。"""
    return [t.ref for t in tokenize_formula(formula) if t.kind == "ref"]


def parse_names(formula: str) -> list[RawName]:
    """解析公式文本，返回其中出现的定义名称。"""
    return [t.name for t in tokenize_formula(formula) if t.kind == "name"]
