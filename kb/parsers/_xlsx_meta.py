"""XLSX 包内元数据读取：工作表 XML 路径、合并区域、公式单元格。

为什么需要这一层：

- openpyxl 的只读工作表 ``ReadOnlyWorksheet`` 不会解析 ``<mergeCells>``
  （见 openpyxl 3.1.5 ``worksheet/_read_only.py``，该类没有 ``merged_cells`` 属性），
  所以只读模式下拿不到合并区域，必须从 xlsx 包（zip）里读。
- ``data_only=True`` 时 openpyxl 直接返回缓存值，不区分「常量」与「公式无缓存值」，
  为了如实告知调用方，这里额外记录哪些单元格是公式、是否带缓存值。
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass, field
from xml.etree import ElementTree as ET

from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_OFFICE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

__all__ = ["MergeRange", "SheetMeta", "sheet_xml_paths", "sheet_metadata"]


@dataclass(frozen=True)
class MergeRange:
    """一个合并区域（闭区间，行列均为 1 基）。"""

    min_row: int
    min_col: int
    max_row: int
    max_col: int

    @property
    def ref(self) -> str:
        """区域的 A1 记法，仅用于日志。"""
        return "{}:{}".format(self.min_row, self.min_col)


@dataclass(frozen=True)
class SheetMeta:
    """工作表的合并区域与公式单元格信息。"""

    merges: tuple[MergeRange, ...] = ()
    # (row, col) -> 该单元格是公式且是否带缓存值
    formula_cells: dict[tuple[int, int], bool] = field(default_factory=dict)


def sheet_xml_paths(archive: zipfile.ZipFile) -> dict[str, str]:
    """工作表名 -> zip 内 worksheet XML 的成员路径。"""
    workbook_xml = ET.fromstring(archive.read("xl/workbook.xml"))
    rels_xml = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {
        rel.get("Id"): rel.get("Target")
        for rel in rels_xml.findall("{" + NS_PKG_REL + "}Relationship")
    }
    paths: dict[str, str] = {}
    sheets = workbook_xml.find("{" + NS_MAIN + "}sheets")
    if sheets is None:
        return paths
    for sheet in sheets:
        title = sheet.get("name")
        rid = sheet.get("{" + NS_REL_OFFICE + "}id")
        target = targets.get(rid)
        if not title or not target:
            continue
        paths[title] = _resolve_part(target)
    return paths


def _resolve_part(target: str) -> str:
    """把关系里的 Target 归一化成 zip 成员路径。"""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join("xl", target))


def sheet_metadata(archive: zipfile.ZipFile, part: str) -> SheetMeta:
    """读取一个工作表 XML 的合并区域与公式单元格。"""
    merges: list[MergeRange] = []
    formulas: dict[tuple[int, int], bool] = {}
    with archive.open(part) as source:
        for _event, elem in ET.iterparse(source, events=("end",)):
            tag = elem.tag
            if tag == "{" + NS_MAIN + "}c":
                coord = elem.get("r")
                if coord and elem.find("{" + NS_MAIN + "}f") is not None:
                    value = elem.find("{" + NS_MAIN + "}v")
                    has_cache = value is not None and (value.text or "").strip() != ""
                    row, col = coordinate_to_tuple(coord)
                    formulas[(row, col)] = has_cache
            elif tag == "{" + NS_MAIN + "}mergeCell":
                ref = elem.get("ref")
                if ref:
                    min_col, min_row, max_col, max_row = range_boundaries(ref)
                    merges.append(MergeRange(min_row, min_col, max_row, max_col))
            elem.clear()
    return SheetMeta(merges=tuple(merges), formula_cells=formulas)
