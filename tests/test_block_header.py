"""验收 4：每 50 行一个 table block，且每个 block 的第一行是表头。"""

from __future__ import annotations

from kb.parsers.xlsx import TABLE_ROWS_PER_BLOCK, parse_xlsx
from tests.conftest import data_rows, table_blocks


def test_120_rows_split_into_three_blocks(block_workbook):
    """120 个数据行 -> 50 + 50 + 20 三个表格块。"""
    blocks = parse_xlsx(block_workbook)
    tables = table_blocks(blocks)
    assert TABLE_ROWS_PER_BLOCK == 50
    assert len(tables) == 3
    assert [len(data_rows(table)) for table in tables] == [50, 50, 20]


def test_every_block_starts_with_header(block_workbook):
    """每个块的 rows[0] 都等于表头行，且不是数据行。"""
    tables = table_blocks(parse_xlsx(block_workbook))
    header = ("编号", "名称", "金额(元)")
    for table in tables:
        assert table.rows[0] == header
        assert "编号" not in [value for row in table.rows[1:] for value in row]
    assert tables[0].rows[1][0] == "R-001"
    assert tables[1].rows[1][0] == "R-051"
    assert tables[2].rows[1][0] == "R-101"


def test_sample_meeting_rooms_blocks(meeting_rooms_path):
    """样例会议室表同样是 3 个块，每块首行都等于表头。"""
    tables = table_blocks(parse_xlsx(meeting_rooms_path))
    assert len(tables) == 3
    assert [len(data_rows(table)) for table in tables] == [50, 50, 20]
    header = (
        "会议室编号",
        "会议室名称",
        "楼层",
        "面积(㎡)",
        "容纳人数",
        "显示设备",
        "计费标准(元/半天)",
    )
    for table in tables:
        assert table.rows[0] == header
