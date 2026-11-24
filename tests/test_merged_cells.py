"""验收 1：合并单元格展开后区域内每格都有值。"""

from __future__ import annotations

from kb.parsers.xlsx import parse_xlsx
from tests.conftest import data_rows, table_blocks


def test_horizontal_merge_fills_every_cell(merged_workbook):
    """横向合并 3 格：区域里三个格子都必须等于锚点值且非空。"""
    blocks = parse_xlsx(merged_workbook)
    tables = table_blocks(blocks)
    assert len(tables) == 1
    rows = data_rows(tables[0])
    merged_row = rows[0]
    assert merged_row == ("横向合并三格", "横向合并三格", "横向合并三格")
    for value in merged_row:
        assert value is not None
        assert str(value).strip() != ""


def test_vertical_merge_fills_every_row(merged_workbook):
    """纵向合并 2 行：上下两行的对应格子都拿到锚点值。"""
    blocks = parse_xlsx(merged_workbook)
    rows = data_rows(table_blocks(blocks)[0])
    assert rows[1][0] == "纵向合并两行"
    assert rows[2][0] == "纵向合并两行"
    assert rows[1][0] == rows[2][0]
    # 合并区域之外的格子保持原值，没有被覆盖。
    assert rows[1][1:] == (10, 2)
    assert rows[2][1:] == (20, 4)


def test_sample_pricelist_merges_are_expanded(pricelist_path):
    """样例价目表：F2:F3（纵向 2 行）与 C12:E12（横向 3 格）都要填满。"""
    blocks = parse_xlsx(pricelist_path)
    rows = data_rows(table_blocks(blocks)[0])
    # 纵向合并：第 2、3 个数据行的「区域」列相同且非空。
    assert rows[0][5] == "华东"
    assert rows[1][5] == "华东"
    # 横向合并：第 11 个数据行（工作表第 12 行）的三列相同且非空。
    promo = rows[10]
    assert promo[2] == promo[3] == promo[4] == "组合优惠：任选三件九折（已含税）"
