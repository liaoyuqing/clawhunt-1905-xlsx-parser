"""公式取缓存值：data_only=True 读取，样例文件里真实存在缓存值。"""

from __future__ import annotations

from openpyxl import Workbook, load_workbook

from kb.parsers.xlsx import parse_xlsx
from tests.conftest import data_rows, table_blocks


def test_sample_pricelist_really_has_formula_and_cache(pricelist_path):
    """样例文件本身就是「带缓存的公式文件」：有公式，也有缓存值。"""
    with_formula = load_workbook(pricelist_path, read_only=True, data_only=False)
    formula = with_formula["价目表"]["E2"].value
    with_formula.close()
    assert formula == "=C2*D2"

    with_cache = load_workbook(pricelist_path, read_only=True, data_only=True)
    cached = with_cache["价目表"]["E2"].value
    with_cache.close()
    assert cached == 1200


def test_parser_returns_cached_formula_value(pricelist_path):
    """解析结果里金额列是缓存值（数字），不是公式字符串、不是空值。"""
    rows = data_rows(table_blocks(parse_xlsx(pricelist_path))[0])
    assert rows[0][4] == 1200
    assert rows[1][4] == 4800
    amounts = [row[4] for row in rows if row[4] is not None and isinstance(row[4], (int, float))]
    assert len(amounts) == 29  # 30 个数据行里有 1 行是合并的说明行，不含公式
    for row in rows:
        if isinstance(row[4], str):
            assert not row[4].startswith("=")


def test_formula_without_cache_is_reported(tmp_path):
    """没有缓存值的公式只能读到空值——此时必须记一条 warning，而不是假装有值。"""
    path = tmp_path / "no_cache.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "无缓存"
    sheet.append(["单价", "数量", "金额"])
    sheet.append([100, 3, "=A2*B2"])
    workbook.save(path)
    workbook.close()

    warnings: list[str] = []
    blocks = parse_xlsx(path, warnings=warnings)
    rows = data_rows(table_blocks(blocks)[0])
    assert rows[0][2] is None
    assert len(warnings) == 1
    assert "没有缓存值" in warnings[0]
