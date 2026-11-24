"""测试夹具：样例都在临时目录现造，数据全部虚构。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable, Sequence

import pytest
from openpyxl import Workbook

REPO_ROOT = Path(__file__).resolve().parent.parent
SAMPLES_DIR = REPO_ROOT / "samples"
EXPECTED_DIR = SAMPLES_DIR / "expected"


def build_workbook(
    path: Path,
    rows: Sequence[Sequence[object]],
    *,
    title: str = "数据表",
    merges: Iterable[str] = (),
    extra_sheets: Iterable[tuple[str, str, Sequence[Sequence[object]]]] = (),
) -> Path:
    """写一个最小工作簿：标题、行、合并区域、附加工作表（可指定 hidden 状态）。"""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = title
    for row in rows:
        sheet.append(list(row))
    for reference in merges:
        sheet.merge_cells(reference)
    for name, state, sheet_rows in extra_sheets:
        extra = workbook.create_sheet(name)
        extra.sheet_state = state
        for row in sheet_rows:
            extra.append(list(row))
    workbook.save(path)
    workbook.close()
    return path


def table_blocks(blocks):
    """只取表格块。"""
    return [block for block in blocks if block.kind == "table"]


def heading_texts(blocks):
    """所有标题块的文本。"""
    return [block.text for block in blocks if block.kind == "heading"]


def paragraph_texts(blocks):
    """所有段落块的文本。"""
    return [block.text for block in blocks if block.kind == "paragraph"]


def data_rows(block):
    """表格块里的数据行（去掉重复的表头行）。"""
    return list(block.rows[1:])


@pytest.fixture
def merged_workbook(tmp_path: Path) -> Path:
    """横向合并 3 格 + 纵向合并 2 行。"""
    rows = [
        ["项目", "单价(元)", "数量"],
        ["横向合并三格", None, None],  # 第 2 行：A2:C2 横向合并，锚点在 A2
        ["纵向合并两行", 10, 2],
        [None, 20, 4],  # 第 4 行：A4 属于 A3:A4 合并区域
    ]
    path = tmp_path / "merged.xlsx"
    return build_workbook(
        path,
        rows,
        title="合并单元格示例",
        merges=("A2:C2", "A3:A4"),
    )


@pytest.fixture
def truncation_workbook(tmp_path: Path) -> Path:
    """25 个数据行，用于验证 KB_XLSX_MAX_ROWS=10 时的截断提示。"""
    rows = [["序号", "名称", "数值"]]
    for index in range(1, 26):
        rows.append([index, "虚构物料 {:02d}".format(index), index * 10])
    return build_workbook(tmp_path / "truncation.xlsx", rows, title="截断示例")


@pytest.fixture
def hidden_workbook(tmp_path: Path) -> Path:
    """一张可见表 + 一张 hidden + 一张 veryHidden。"""
    rows = [["编号", "名称"], [1, "可见行"]]
    hidden = [["字段", "内容"], ["底价", "SECRET-内部底价"]]
    very_hidden = [["字段", "内容"], ["旧版", "SECRET-归档旧版"]]
    path = tmp_path / "hidden.xlsx"
    return build_workbook(
        path,
        rows,
        title="可见表",
        extra_sheets=(("内部备注", "hidden", hidden), ("归档", "veryHidden", very_hidden)),
    )


@pytest.fixture
def block_workbook(tmp_path: Path) -> Path:
    """120 个数据行，应当切成 50 + 50 + 20 三个表格块。"""
    rows = [["编号", "名称", "金额(元)"]]
    for index in range(1, 121):
        rows.append(["R-{:03d}".format(index), "虚构条目 {:03d}".format(index), index])
    return build_workbook(tmp_path / "blocks.xlsx", rows, title="分块示例")


@pytest.fixture
def pricelist_path() -> Path:
    return SAMPLES_DIR / "pricelist.xlsx"


@pytest.fixture
def meeting_rooms_path() -> Path:
    return SAMPLES_DIR / "meeting_rooms.xlsx"


def load_snapshot(name: str) -> dict:
    """读取已提交的期望输出快照。"""
    return json.loads((EXPECTED_DIR / name).read_text(encoding="utf-8"))
