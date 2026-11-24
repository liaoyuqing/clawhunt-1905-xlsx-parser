"""生成两个虚构样例 xlsx 与期望输出快照（工单 S3-KB-007）。

样例全部为虚构数据，不含任何真实企业名与人名：

* ``samples/pricelist.xlsx``：会员空间价目表，含纵向 2 行合并（区域列）与
  横向 3 格合并（组合优惠说明行），金额列是公式 ``=单价*数量``。
* ``samples/meeting_rooms.xlsx``：会议室规格表，120 个数据行（切成 3 个表格块），
  另有两张隐藏表（hidden / veryHidden），解析时必须跳过。

公式缓存值：openpyxl 只写公式、不写计算结果，所以保存后用 zipfile + ElementTree
重写工作表 XML，把算好的结果写进 ``<v>``，效果等同「文件被 Excel 保存过」。
这样 ``load_workbook(..., data_only=True)`` 才能读到缓存值（见 README 说明）。

用法::

    python tools/make_samples.py          # 重新生成样例与快照
    python tools/make_samples.py --check  # 校验提交的快照是否与生成结果一致
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from openpyxl import Workbook  # noqa: E402
from openpyxl.styles import Font, Alignment  # noqa: E402

from kb.parsers._xlsx_meta import NS_MAIN, sheet_xml_paths  # noqa: E402
from kb.parsers.xlsx import DEFAULT_MAX_ROWS, parse_xlsx  # noqa: E402

SAMPLES_DIR = REPO_ROOT / "samples"
EXPECTED_DIR = SAMPLES_DIR / "expected"

PRICELIST = "pricelist.xlsx"
MEETING_ROOMS = "meeting_rooms.xlsx"

PRICELIST_SHEET = "价目表"
ROOMS_SHEET = "会议室规格"
ROOMS_HIDDEN_SHEET = "内部备注"
ROOMS_VERY_HIDDEN_SHEET = "归档"

PRICELIST_HEADER = ["产品编号", "产品名称", "单价(元)", "数量", "金额(元)", "区域", "备注"]
ROOMS_HEADER = [
    "会议室编号",
    "会议室名称",
    "楼层",
    "面积(㎡)",
    "容纳人数",
    "显示设备",
    "计费标准(元/半天)",
]

#: 价目表的虚构产品线：名称、单价、备注。
PRICE_ITEMS = [
    ("开放工位·月卡", 1200, "含水电与网络"),
    ("独立办公室·月卡", 4800, ""),
    ("会议室 4 人间·半天", 380, "含投屏与白板"),
    ("会议室 8 人间·半天", 560, ""),
    ("路演厅·整场", 3600, "含音响与灯光"),
    ("直播间·小时", 240, ""),
    ("打印额度·100 页", 30, ""),
    ("储物柜·月租", 80, ""),
    ("注册地址·年费", 6000, "含信件代收"),
    ("前台代收·月包", 200, ""),
]

PRICE_REGIONS = ["华东", "华南", "华北", "西南"]

#: 组合优惠行（横向合并 3 格）落在价目表第几行。
PROMO_ROW = 12
PROMO_TEXT = "组合优惠：任选三件九折（已含税）"

ROOM_TYPES = [
    ("小型洽谈室", 4, 12),
    ("标准会议室", 8, 28),
    ("培训教室", 16, 56),
    ("多功能厅", 40, 150),
    ("直播工作室", 6, 30),
]
ROOM_FLOORS = ["3 层", "5 层", "8 层", "12 层"]
ROOM_DISPLAYS = ["75 寸电视", "激光投影", "LED 屏", "无"]
ROOM_FEE_BY_TYPE = {
    "小型洽谈室": 380,
    "标准会议室": 620,
    "培训教室": 1180,
    "多功能厅": 2400,
    "直播工作室": 760,
}
ROOM_COUNT = 120


def _style_header(sheet, columns: int) -> None:
    for column in range(1, columns + 1):
        cell = sheet.cell(row=1, column=column)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")
    sheet.freeze_panes = "A2"


def pricelist_rows() -> list[list[object]]:
    """价目表的 30 个数据行；金额列先留空，由公式填充。"""
    rows: list[list[object]] = []
    for index in range(30):
        name, unit_price, note = PRICE_ITEMS[index % len(PRICE_ITEMS)]
        quantity = 1 + (index // 10)
        if index == 0:
            region: object = PRICE_REGIONS[0]
        elif index == 1:
            region = None  # 与上一行纵向合并
        else:
            region = PRICE_REGIONS[(index - 2) % len(PRICE_REGIONS)]
        if index + 2 == PROMO_ROW:
            rows.append([None, None, PROMO_TEXT, None, None, None, "不与其他优惠叠加"])
            continue
        rows.append([
            "P-{:04d}".format(index + 1),
            name,
            unit_price,
            quantity,
            None,  # 金额：公式单元格
            region,
            note,
        ])
    return rows


def build_pricelist(path: Path) -> dict[str, dict[str, object]]:
    """生成价目表样例，返回需要注入缓存值的公式单元格。"""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = PRICELIST_SHEET
    sheet.append(PRICELIST_HEADER)
    rows = pricelist_rows()
    for row in rows:
        sheet.append(row)
    _style_header(sheet, len(PRICELIST_HEADER))
    for column, width in zip("ABCDEFG", (12, 24, 12, 8, 14, 10, 22)):
        sheet.column_dimensions[column].width = width
    for row_number in range(2, len(rows) + 2):
        sheet.cell(row=row_number, column=3).number_format = "#,##0"
        sheet.cell(row=row_number, column=5).number_format = "#,##0"
    # 纵向合并 2 行：同一区域覆盖两个产品。
    sheet.merge_cells("F2:F3")
    # 横向合并 3 格：组合优惠说明横跨 单价/数量/金额 三列。
    sheet.merge_cells("C{}:E{}".format(PROMO_ROW, PROMO_ROW))

    cached: dict[str, float] = {}
    for offset, row in enumerate(rows):
        row_number = offset + 2
        if row_number == PROMO_ROW:
            continue
        sheet.cell(row=row_number, column=5).value = "=C{r}*D{r}".format(r=row_number)
        cached["E{}".format(row_number)] = float(row[2]) * float(row[3])
    workbook.save(path)
    return {PRICELIST_SHEET: cached}


def meeting_room_rows(count: int = ROOM_COUNT) -> list[list[object]]:
    """会议室规格表的 ``count`` 个数据行（全部虚构）。"""
    rows: list[list[object]] = []
    for index in range(count):
        room_type, capacity, area = ROOM_TYPES[index % len(ROOM_TYPES)]
        floor = ROOM_FLOORS[(index // len(ROOM_TYPES)) % len(ROOM_FLOORS)]
        number = "MR-{:03d}".format(index + 1)
        name = "{} {}".format(room_type, 101 + index)
        rows.append([
            number,
            name,
            floor,
            area + (index % 3),
            capacity + (index % 2),
            ROOM_DISPLAYS[index % len(ROOM_DISPLAYS)],
            ROOM_FEE_BY_TYPE[room_type] + 60 * (((index // len(ROOM_TYPES))) % 4),
        ])
    return rows


def build_meeting_rooms(path: Path) -> dict[str, dict[str, object]]:
    """生成会议室规格表样例（含隐藏工作表），无公式。"""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = ROOMS_SHEET
    sheet.append(ROOMS_HEADER)
    rows = meeting_room_rows()
    for row in rows:
        sheet.append(row)
    _style_header(sheet, len(ROOMS_HEADER))
    for column, width in zip("ABCDEFG", (14, 22, 10, 12, 12, 14, 22)):
        sheet.column_dimensions[column].width = width
    # 纵向合并：前三间会议室同属 3 层。
    sheet.merge_cells("C2:C4")

    hidden = workbook.create_sheet(ROOMS_HIDDEN_SHEET)
    hidden.append(["备注项", "内容"])
    hidden.append(["内部资料", "以下内容不对外报价"])
    hidden.append(["底价", "SECRET-内部底价，不要外传"])
    hidden.append(["联系人", "虚构联系人 张样例"])
    hidden.sheet_state = "hidden"

    very_hidden = workbook.create_sheet(ROOMS_VERY_HIDDEN_SHEET)
    very_hidden.append(["归档项", "内容"])
    very_hidden.append(["历史版本", "SECRET-veryHidden-2024 年旧版规格"])
    very_hidden.sheet_state = "veryHidden"

    workbook.save(path)
    return {}


def inject_cached_values(path: Path, cached: dict[str, dict[str, object]]) -> int:
    """把公式计算结果写进工作表 XML 的 ``<v>``，模拟 Excel 保存后的缓存值。"""
    if not cached:
        return 0
    temporary = path.with_name(path.name + ".tmp")
    written = 0
    with zipfile.ZipFile(path) as source:
        parts = sheet_xml_paths(source)
        wanted = {
            parts[title]: cells for title, cells in cached.items() if title in parts
        }
        members = [(info, source.read(info.filename)) for info in source.infolist()]
    ET.register_namespace("", NS_MAIN)
    with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as target:
        for info, payload in members:
            cells = wanted.get(info.filename)
            if cells:
                root = ET.fromstring(payload)
                for cell in root.iter("{%s}c" % NS_MAIN):
                    reference = cell.get("r")
                    if reference not in cells:
                        continue
                    node = cell.find("{%s}v" % NS_MAIN)
                    if node is None:
                        node = ET.SubElement(cell, "{%s}v" % NS_MAIN)
                    node.text = repr(cells[reference])
                    written += 1
                payload = ET.tostring(root, xml_declaration=True, encoding="UTF-8")
            target.writestr(info, payload)
    os.replace(temporary, path)
    return written


def snapshot_for(name: str) -> dict[str, object]:
    """解析样例并生成期望输出快照。"""
    blocks = parse_xlsx(SAMPLES_DIR / name, max_rows=DEFAULT_MAX_ROWS)
    return {
        "source": "samples/" + name,
        "parser": "kb.parsers.xlsx.parse_xlsx",
        "max_rows": DEFAULT_MAX_ROWS,
        "blocks": [block.to_dict() for block in blocks],
    }


def write_samples(directory: Path) -> list[Path]:
    """在 ``directory`` 下生成两个样例文件，返回生成的文件路径。"""
    directory.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []
    pricelist = directory / PRICELIST
    cached = build_pricelist(pricelist)
    inject_cached_values(pricelist, cached)
    written.append(pricelist)
    rooms = directory / MEETING_ROOMS
    build_meeting_rooms(rooms)
    written.append(rooms)
    return written


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="生成 XLSX 解析器的虚构样例与期望输出")
    parser.add_argument(
        "--check",
        action="store_true",
        help="只在临时目录重新生成并比对已提交的快照，不修改工作区文件",
    )
    arguments = parser.parse_args(argv)

    if arguments.check:
        with tempfile.TemporaryDirectory() as temporary:
            target = Path(temporary)
            write_samples(target)
            drift = 0
            for name in (PRICELIST, MEETING_ROOMS):
                expected_path = EXPECTED_DIR / (Path(name).stem + ".json")
                committed = json.loads(expected_path.read_text(encoding="utf-8"))
                blocks = parse_xlsx(target / name, max_rows=DEFAULT_MAX_ROWS)
                fresh = [block.to_dict() for block in blocks]
                if fresh != committed["blocks"]:
                    drift += 1
                    print("快照不一致: {}".format(expected_path.relative_to(REPO_ROOT)))
            if drift:
                return 1
            print("快照校验通过：2 个样例重新生成后与 samples/expected 完全一致")
            return 0

    written = write_samples(SAMPLES_DIR)
    EXPECTED_DIR.mkdir(parents=True, exist_ok=True)
    for path in written:
        snapshot = snapshot_for(path.name)
        snapshot_path = EXPECTED_DIR / (path.stem + ".json")
        snapshot_path.write_text(
            json.dumps(snapshot, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        tables = sum(1 for block in snapshot["blocks"] if block["kind"] == "table")
        headings = [block["text"] for block in snapshot["blocks"] if block["kind"] == "heading"]
        print(
            "{}: {} 个数据行 -> {} 个块（标题 {}，表格块 {}）".format(
                path.relative_to(REPO_ROOT),
                sum(len(block["rows"]) - 1 for block in snapshot["blocks"] if block["kind"] == "table"),
                len(snapshot["blocks"]),
                headings,
                tables,
            )
        )
        print("  快照: {}".format(snapshot_path.relative_to(REPO_ROOT)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
