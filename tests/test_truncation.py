"""验收 2：超过行数上限时输出截断提示且总行数正确。"""

from __future__ import annotations

from kb.parsers.xlsx import (
    DEFAULT_MAX_ROWS,
    MAX_ROWS_ENV_VAR,
    TRUNCATION_NOTE,
    max_rows_from_env,
    parse_xlsx,
)
from tests.conftest import data_rows, paragraph_texts, table_blocks


def test_truncates_at_ten_rows(truncation_workbook, monkeypatch):
    """KB_XLSX_MAX_ROWS=10 + 25 行数据：只输出 10 行，并追加截断说明。"""
    monkeypatch.setenv(MAX_ROWS_ENV_VAR, "10")
    blocks = parse_xlsx(truncation_workbook)
    rows = data_rows(table_blocks(blocks)[0])
    assert len(rows) == 10
    assert rows[0][0] == 1
    assert rows[-1][0] == 10
    assert "已截断，共 25 行" in paragraph_texts(blocks)


def test_truncation_note_reports_total_not_read(truncation_workbook, monkeypatch):
    """截断提示里的数字是原始总行数 25，不是已读行数 10。"""
    monkeypatch.setenv(MAX_ROWS_ENV_VAR, "10")
    blocks = parse_xlsx(truncation_workbook)
    texts = paragraph_texts(blocks)
    assert TRUNCATION_NOTE.format(total=25) == "已截断，共 25 行"
    assert "已截断，共 25 行" in texts
    assert "已截断，共 10 行" not in texts


def test_default_limit_does_not_truncate(truncation_workbook, monkeypatch):
    """默认上限 2000 行：25 行数据整体输出，不出现截断说明。"""
    monkeypatch.delenv(MAX_ROWS_ENV_VAR, raising=False)
    assert DEFAULT_MAX_ROWS == 2000
    assert max_rows_from_env() == DEFAULT_MAX_ROWS
    blocks = parse_xlsx(truncation_workbook)
    assert len(data_rows(table_blocks(blocks)[0])) == 25
    assert paragraph_texts(blocks) == []
