"""验收 3：隐藏工作表不出现在结果中。"""

from __future__ import annotations

from openpyxl import load_workbook

from kb.parsers.xlsx import parse_xlsx
from tests.conftest import heading_texts


def test_hidden_sheets_are_skipped(hidden_workbook):
    """hidden 与 veryHidden 两张表都不产生标题、不泄漏内容。"""
    workbook = load_workbook(hidden_workbook, read_only=True)
    states = {name: workbook[name].sheet_state for name in workbook.sheetnames}
    workbook.close()
    assert states == {"可见表": "visible", "内部备注": "hidden", "归档": "veryHidden"}

    blocks = parse_xlsx(hidden_workbook)
    assert heading_texts(blocks) == ["可见表"]
    assert "内部备注" not in heading_texts(blocks)
    assert "归档" not in heading_texts(blocks)
    dumped = repr([block.rows for block in blocks]) + repr([block.text for block in blocks])
    assert "SECRET" not in dumped


def test_sample_meeting_rooms_hides_internal_sheet(meeting_rooms_path):
    """样例会议室表里确实有两张隐藏表，且解析结果只含可见表标题。"""
    workbook = load_workbook(meeting_rooms_path, read_only=True)
    states = {name: workbook[name].sheet_state for name in workbook.sheetnames}
    workbook.close()
    assert states["会议室规格"] == "visible"
    assert states["内部备注"] == "hidden"
    assert states["归档"] == "veryHidden"

    blocks = parse_xlsx(meeting_rooms_path)
    assert heading_texts(blocks) == ["会议室规格"]
    dumped = repr([block.rows for block in blocks]) + repr([block.text for block in blocks])
    assert "SECRET" not in dumped
