"""XLSX 解析器：把工作簿转成知识库文档块（工单 S3-KB-007）。

规则（与工单四条验收逐条对应）：

1. 只读模式读取：``load_workbook(path, read_only=True, data_only=True)``。
   ``data_only=True`` 表示公式取 Excel 保存时写入的缓存值；
   openpyxl 自身不会计算公式，没有缓存值时该单元格按空值处理并记一条 warning。
2. 隐藏工作表跳过：``sheet_state`` 显式比对 ``"visible"``，
   ``hidden`` 与 ``veryHidden`` 都不产生标题、不进入任何计数。
3. 每个可见工作表先产出一个 level 1 heading（工作表名）。
4. 表头 = 首个非空行（任意单元格有值的行）；表头去首尾空白，
   空表头列命名为 ``column_{列号}``（列号从 1 开始）。
5. 数据行每 ``TABLE_ROWS_PER_BLOCK``（默认 50）行切成一个 table block，
   每个 block 的第一行重复表头。
6. 合并单元格：区域内每个格子都填成左上角的值（先展开、后切块）。
7. 行数上限：每个工作表最多输出 ``KB_XLSX_MAX_ROWS``（默认 2000）行数据；
   超出时追加一个 paragraph 写明「已截断，共 N 行」，
   N 是该表的实际数据行总数（不含表头），不是已读行数。
"""

from __future__ import annotations

import os
import zipfile
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence

from openpyxl import load_workbook

from ._xlsx_meta import MergeRange, SheetMeta, sheet_metadata, sheet_xml_paths
from .base import Block, register

__all__ = [
    "parse_xlsx",
    "DEFAULT_MAX_ROWS",
    "MAX_ROWS_ENV_VAR",
    "TABLE_ROWS_PER_BLOCK",
]

#: 每个工作表允许输出的最大数据行数默认值。
DEFAULT_MAX_ROWS = 2000

#: 行数上限的环境变量名（每次解析时读取，便于运行期调整）。
MAX_ROWS_ENV_VAR = "KB_XLSX_MAX_ROWS"

#: 一个 table block 容纳的数据行数。
TABLE_ROWS_PER_BLOCK = 50

#: 截断提示的固定文案。
TRUNCATION_NOTE = "已截断，共 {total} 行"


def max_rows_from_env() -> int:
    """读取行数上限；环境变量缺失、为空或非法时回落到默认值。"""
    raw = os.environ.get(MAX_ROWS_ENV_VAR)
    if raw is None or not raw.strip():
        return DEFAULT_MAX_ROWS
    try:
        value = int(raw.strip())
    except ValueError:
        return DEFAULT_MAX_ROWS
    return value if value > 0 else DEFAULT_MAX_ROWS


def _has_content(value: Any) -> bool:
    """判断单元格是否算「非空」。0 与 False 算有值，空串与 None 算空。"""
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() != ""
    return True


def _row_has_content(values: Sequence[Any]) -> bool:
    """判断一行是否非空。"""
    return any(_has_content(value) for value in values)


def _merge_lookup(merges: Iterable[MergeRange]) -> dict[tuple[int, int], tuple[int, int]]:
    """(行, 列) -> 该格所属合并区域的左上角 (行, 列)。"""
    lookup: dict[tuple[int, int], tuple[int, int]] = {}
    for merge in merges:
        for row in range(merge.min_row, merge.max_row + 1):
            for col in range(merge.min_col, merge.max_col + 1):
                if (row, col) == (merge.min_row, merge.min_col):
                    continue
                lookup[(row, col)] = (merge.min_row, merge.min_col)
    return lookup


def _normalize_header(values: Sequence[Any]) -> list[str]:
    """表头规范化：去首尾空白；空表头列命名为 ``column_{列号}``。"""
    names: list[str] = []
    for index, value in enumerate(values, start=1):
        if value is None:
            text = ""
        elif isinstance(value, str):
            text = value.strip()
        else:
            text = str(value).strip()
        names.append(text if text else "column_{}".format(index))
    return names


def _expand_merged(
    rows: Sequence[tuple[int, list[Any]]],
    merges: Sequence[MergeRange],
    width: int,
) -> list[list[Any]]:
    """把合并区域左上角的值填充到区域内每一个格子。

    ``rows`` 是 ``(行号, 值列表)``，行号从 1 开始。
    展开在切块之前完成，因此每个 block 内的合并区域都是完整填满的。
    """
    lookup = _merge_lookup(merges)
    if not lookup:
        return [list(values) for _number, values in rows]
    source: dict[tuple[int, int], Any] = {}
    for number, values in rows:
        for index in range(width):
            value = values[index] if index < len(values) else None
            source[(number, index + 1)] = value
    expanded: list[list[Any]] = []
    for number, values in rows:
        row = list(values)
        if len(row) < width:
            row.extend([None] * (width - len(row)))
        elif len(row) > width:
            row = row[:width]
        for index in range(width):
            if _has_content(row[index]):
                continue
            anchor = lookup.get((number, index + 1))
            if anchor is None:
                continue
            row[index] = source.get(anchor)
        expanded.append(row)
    return expanded


def _collect_rows(
    worksheet: Any,
    max_rows: int,
) -> tuple[tuple[int, list[Any]] | None, list[tuple[int, list[Any]]], int]:
    """流式读取一个工作表。

    返回 ``(表头行, 保留的数据行, 实际数据行总数)``。
    表头之前的空行跳过；表头之后完全为空的行跳过且不计数。
    数据行只保留前 ``max_rows`` 行，但总数照实统计，用于截断提示。
    """
    header: tuple[int, list[Any]] | None = None
    kept: list[tuple[int, list[Any]]] = []
    total = 0
    for number, raw in enumerate(worksheet.iter_rows(values_only=True), start=1):
        values = list(raw)
        if not _row_has_content(values):
            continue
        if header is None:
            header = (number, values)
            continue
        total += 1
        if len(kept) < max_rows:
            kept.append((number, values))
    return header, kept, total


def _sheet_width(
    worksheet: Any,
    meta: SheetMeta,
    rows: Sequence[tuple[int, list[Any]]],
) -> int:
    """估算表格列数：声明的维度、合并区域与实际读到的行三者取最大。"""
    width = 0
    declared = getattr(worksheet, "max_column", None)
    if isinstance(declared, int) and declared > 0:
        width = declared
    for merge in meta.merges:
        width = max(width, merge.max_col)
    for _number, values in rows:
        width = max(width, len(values))
    return width


def _blocks_for_sheet(
    title: str,
    header: tuple[int, list[Any]] | None,
    kept: list[tuple[int, list[Any]]],
    total: int,
    meta: SheetMeta,
    width: int,
    warnings: list[str] | None,
) -> list[Block]:
    """把一个工作表转成一串文档块。"""
    blocks: list[Block] = [Block(kind="heading", level=1, text=title, sheet=title)]
    if header is None:
        return blocks
    header_number, header_values = header
    all_rows: list[tuple[int, list[Any]]] = [(header_number, header_values)] + list(kept)
    expanded = _expand_merged(all_rows, meta.merges, width)
    header_names = _normalize_header(expanded[0])
    data_rows = expanded[1:]
    if data_rows:
        for start in range(0, len(data_rows), TABLE_ROWS_PER_BLOCK):
            chunk = data_rows[start:start + TABLE_ROWS_PER_BLOCK]
            rows = (tuple(header_names),) + tuple(tuple(row) for row in chunk)
            blocks.append(Block(kind="table", rows=rows, sheet=title))
    if total > len(data_rows):
        blocks.append(
            Block(kind="paragraph", text=TRUNCATION_NOTE.format(total=total), sheet=title)
        )
    return blocks


def _formula_warnings(
    title: str,
    meta: SheetMeta,
    values_by_cell: dict[tuple[int, int], Any],
    warnings: list[str] | None,
) -> None:
    """公式没有缓存值时记一条 warning（不静默吞掉）。"""
    if warnings is None:
        return
    for (row, col), has_cache in sorted(meta.formula_cells.items()):
        if has_cache:
            continue
        value = values_by_cell.get((row, col))
        if value is None:
            warnings.append(
                "工作表 {!r} 的公式单元格 R{}C{} 没有缓存值，已按空值处理".format(title, row, col)
            )


def _load_sheet_meta(path: Path, titles: Iterable[str]) -> dict[str, SheetMeta]:
    """从 xlsx 包里读出每个工作表的合并区域与公式单元格。"""
    metas: dict[str, SheetMeta] = {}
    wanted = set(titles)
    if not wanted:
        return metas
    try:
        with zipfile.ZipFile(path) as archive:
            paths = sheet_xml_paths(archive)
            for title, part in paths.items():
                if title not in wanted:
                    continue
                try:
                    metas[title] = sheet_metadata(archive, part)
                except KeyError:
                    metas[title] = SheetMeta()
    except (zipfile.BadZipFile, OSError):
        return metas
    return metas


def parse_xlsx(
    path: str | Path,
    *,
    warnings: list[str] | None = None,
    max_rows: int | None = None,
) -> list[Block]:
    """解析一个 xlsx 文件，返回文档块列表。

    :param path: xlsx 文件路径。
    :param warnings: 可选的警告收集列表；公式没有缓存值等情况会写进去。
    :param max_rows: 覆盖每表数据行上限；默认读 ``KB_XLSX_MAX_ROWS``。
    """
    target = Path(path)
    limit = max_rows if max_rows is not None else max_rows_from_env()
    workbook = load_workbook(target, read_only=True, data_only=True)
    try:
        worksheets = [ws for ws in workbook.worksheets]
        metas = _load_sheet_meta(target, [ws.title for ws in worksheets])
        blocks: list[Block] = []
        for worksheet in worksheets:
            if worksheet.sheet_state != "visible":
                continue
            title = worksheet.title
            meta = metas.get(title, SheetMeta())
            header, kept, total = _collect_rows(worksheet, limit)
            width = _sheet_width(worksheet, meta, ([header] if header else []) + kept)
            sheet_blocks = _blocks_for_sheet(
                title, header, kept, total, meta, width, warnings
            )
            if warnings is not None and meta.formula_cells:
                cells: dict[tuple[int, int], Any] = {}
                for number, values in ([header] if header else []) + kept:
                    for index, value in enumerate(values, start=1):
                        cells[(number, index)] = value
                _formula_warnings(title, meta, cells, warnings)
            blocks.extend(sheet_blocks)
        return blocks
    finally:
        workbook.close()


# 注册到解析器表：入库流水线按 .xlsx 后缀分发。
register(".xlsx", parse_xlsx, overwrite=True)
