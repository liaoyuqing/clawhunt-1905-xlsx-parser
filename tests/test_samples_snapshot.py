"""样例快照与注册表。"""

from __future__ import annotations

import subprocess
import sys

from kb.parsers import parse_file
from kb.parsers.base import get_parser, supported_suffixes
from kb.parsers.xlsx import parse_xlsx
from tests.conftest import REPO_ROOT, data_rows, heading_texts, load_snapshot, table_blocks


def _assert_snapshot(sample_name: str, snapshot_name: str) -> None:
    snapshot = load_snapshot(snapshot_name)
    blocks = parse_xlsx(REPO_ROOT / "samples" / sample_name, max_rows=2000)
    assert [block.to_dict() for block in blocks] == snapshot["blocks"]


def test_pricelist_matches_snapshot(pricelist_path):
    _assert_snapshot("pricelist.xlsx", "pricelist.json")
    snapshot = load_snapshot("pricelist.json")
    assert snapshot["max_rows"] == 2000
    assert snapshot["blocks"][0]["kind"] == "heading"
    assert snapshot["blocks"][0]["level"] == 1
    assert snapshot["blocks"][0]["text"] == "价目表"


def test_meeting_rooms_matches_snapshot(meeting_rooms_path):
    _assert_snapshot("meeting_rooms.xlsx", "meeting_rooms.json")
    blocks = parse_xlsx(meeting_rooms_path)
    assert heading_texts(blocks) == ["会议室规格"]
    assert sum(len(data_rows(table)) for table in table_blocks(blocks)) == 120


def test_generator_reproduces_snapshots() -> None:
    """tools/make_samples.py --check 重新生成样例并与提交的快照比对。"""
    result = subprocess.run(
        [sys.executable, "tools/make_samples.py", "--check"],
        cwd=str(REPO_ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert result.returncode == 0, result.stdout + result.stderr
    assert "快照校验通过" in result.stdout


def test_registry_dispatches_xlsx(pricelist_path) -> None:
    """注册表按 .xlsx 后缀分发到 parse_xlsx。"""
    assert ".xlsx" in supported_suffixes()
    assert get_parser(".xlsx") is parse_xlsx
    assert [block.to_dict() for block in parse_file(pricelist_path)] == [
        block.to_dict() for block in parse_xlsx(pricelist_path)
    ]
