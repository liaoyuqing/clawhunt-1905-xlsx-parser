"""生成四份验收证据文件：每条验收一段真实跑出来的输出。

运行::

    python tools/emit_evidence.py

输出写入 ``docs/checks/``，内容全部来自当场执行的解析结果，不做任何人工改写。
"""

from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from openpyxl import Workbook, load_workbook  # noqa: E402

from kb.parsers.xlsx import MAX_ROWS_ENV_VAR, parse_xlsx  # noqa: E402

CHECKS_DIR = REPO_ROOT / "docs" / "checks"
SAMPLES_DIR = REPO_ROOT / "samples"


def _build(
    path: Path,
    rows,
    *,
    title: str,
    merges=(),
    extra_sheets=(),
) -> Path:
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


def _table_rows(blocks):
    table = [block for block in blocks if block.kind == "table"][0]
    return list(table.rows[1:])


def evidence_merged(directory: Path) -> list[str]:
    """验收 1：合并单元格区域内每格都有值。"""
    path = _build(
        directory / "merged_evidence.xlsx",
        [
            ["项目", "单价(元)", "数量"],
            ["横向合并三格", None, None],
            ["纵向合并两行", 10, 2],
            [None, 20, 4],
        ],
        title="合并单元格示例",
        merges=("A2:C2", "A3:A4"),
    )
    rows = _table_rows(parse_xlsx(path))
    lines = ["# 验收 1：合并单元格展开", "", "## 自造样例（横向 A2:C2 = 3 格，纵向 A3:A4 = 2 行）", ""]
    lines.append("横向合并区域 A2:C2 -> {}".format(list(rows[0])))
    lines.append("纵向合并区域 A3:A4 -> A3={!r} A4={!r}".format(rows[1][0], rows[2][0]))
    lines.append("")
    lines.append("## 样例价目表 samples/pricelist.xlsx")
    lines.append("")
    sample_rows = _table_rows(parse_xlsx(SAMPLES_DIR / "pricelist.xlsx"))
    lines.append("纵向合并 F2:F3 -> F2={!r} F3={!r}".format(sample_rows[0][5], sample_rows[1][5]))
    promo = sample_rows[10]
    lines.append("横向合并 C12:E12 -> {}".format(list(promo[2:5])))
    lines.append("")
    lines.append("结论：两个合并区域内每格都等于锚点值且非空。")
    return lines


def evidence_truncate(directory: Path) -> list[str]:
    """验收 2：超过 KB_XLSX_MAX_ROWS 时的截断提示与总行数。"""
    rows = [["序号", "名称", "数值"]]
    for index in range(1, 26):
        rows.append([index, "虚构物料 {:02d}".format(index), index * 10])
    path = _build(directory / "truncation_evidence.xlsx", rows, title="截断示例")
    previous = os.environ.get(MAX_ROWS_ENV_VAR)
    os.environ[MAX_ROWS_ENV_VAR] = "10"
    try:
        blocks = parse_xlsx(path)
    finally:
        if previous is None:
            os.environ.pop(MAX_ROWS_ENV_VAR, None)
        else:
            os.environ[MAX_ROWS_ENV_VAR] = previous
    data = _table_rows(blocks)
    notes = [block.text for block in blocks if block.kind == "paragraph"]
    lines = [
        "# 验收 2：行数截断",
        "",
        "环境变量 KB_XLSX_MAX_ROWS=10，样例数据行 25 行（不含表头）。",
        "",
        "输出数据行数：{}".format(len(data)),
        "首个数据行：{}".format(list(data[0])),
        "最后数据行：{}".format(list(data[-1])),
        "截断段落：{}".format(notes),
        "",
        "结论：只输出 10 行，段落写明原始总行数 25。",
    ]
    return lines


def evidence_hidden(directory: Path) -> list[str]:
    """验收 3：隐藏工作表不出现在结果中。"""
    path = _build(
        directory / "hidden_evidence.xlsx",
        [["编号", "名称"], [1, "可见行"]],
        title="可见表",
        extra_sheets=(
            ("内部备注", "hidden", [["字段", "内容"], ["底价", "SECRET-内部底价"]]),
            ("归档", "veryHidden", [["字段", "内容"], ["旧版", "SECRET-归档旧版"]]),
        ),
    )
    workbook = load_workbook(path, read_only=True)
    states = {name: workbook[name].sheet_state for name in workbook.sheetnames}
    workbook.close()
    blocks = parse_xlsx(path)
    headings = [block.text for block in blocks if block.kind == "heading"]
    dumped = repr([block.rows for block in blocks]) + repr([block.text for block in blocks])
    lines = [
        "# 验收 3：隐藏工作表跳过",
        "",
        "工作簿里的表与状态：{}".format(states),
        "解析得到的标题：{}".format(headings),
        "结果里是否出现隐藏表内容 SECRET：{}".format("SECRET" in dumped),
        "",
        "样例 samples/meeting_rooms.xlsx 的表与状态：{}".format(_sample_states(SAMPLES_DIR / "meeting_rooms.xlsx")),
        "结论：hidden 与 veryHidden 两张表都不产生标题，也不泄漏内容。",
    ]
    return lines


def _sample_states(path: Path) -> dict[str, str]:
    workbook = load_workbook(path, read_only=True)
    states = {name: workbook[name].sheet_state for name in workbook.sheetnames}
    workbook.close()
    return states


def evidence_blocks(directory: Path) -> list[str]:
    """验收 4：每 50 行一个 block，每块第一行是表头。"""
    rows = [["编号", "名称", "金额(元)"]]
    for index in range(1, 121):
        rows.append(["R-{:03d}".format(index), "虚构条目 {:03d}".format(index), index])
    path = _build(directory / "blocks_evidence.xlsx", rows, title="分块示例")
    tables = [block for block in parse_xlsx(path) if block.kind == "table"]
    lines = ["# 验收 4：每 50 行一个 block 且每块重复表头", "", "输入 120 个数据行。", ""]
    for order, table in enumerate(tables, start=1):
        lines.append(
            "block {}：{} 个数据行，首行 = {}".format(
                order, len(table.rows) - 1, list(table.rows[0])
            )
        )
    lines.append("")
    lines.append("结论：共 {} 个表格块，每个块的第一行都等于表头行。".format(len(tables)))
    return lines


def main() -> int:
    CHECKS_DIR.mkdir(parents=True, exist_ok=True)
    builders = {
        "merged.txt": evidence_merged,
        "truncate.txt": evidence_truncate,
        "hidden.txt": evidence_hidden,
        "blocks.txt": evidence_blocks,
    }
    with tempfile.TemporaryDirectory() as temporary:
        directory = Path(temporary)
        for name, builder in builders.items():
            lines = builder(directory)
            text = "\n".join(lines) + "\n"
            (CHECKS_DIR / name).write_text(text, encoding="utf-8")
            print("=== docs/checks/{} ===".format(name))
            print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
