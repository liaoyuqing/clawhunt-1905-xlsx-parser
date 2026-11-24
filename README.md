# clawhunt-1905-xlsx-parser

会员空间平台知识库的 XLSX 解析器（工单号 **S3-KB-007**）：把 `.xlsx` 工作簿转成知识库文档块，
并按文件后缀注册到入库流水线的解析器表。

处理的是「人做的表」里最常见的三件事：**合并单元格**（表头或分区经常横竖合并）、
**行数过多的表**（几千行的价目表不能整表塞进检索块）、**隐藏工作表**（内部底价、归档版本）。

## 依赖与环境

| 项 | 版本 | 说明 |
| --- | --- | --- |
| Python | 3.11+（实测 3.11.9） | 只用到标准库 + openpyxl |
| openpyxl | 3.1.5 | 唯一的运行期依赖，见 `requirements.txt` |
| pytest | 9.2.0（`requirements.txt` 允许 8.x–9.x） | 仅跑测试用 |

不需要：数据库、网络、Excel/WPS 可执行文件、pandas、xlrd。解析过程不发起任何网络请求，
只读取本地文件，样例数据全部虚构。只支持 `.xlsx`，`.xls` / `.xlsm` 不在本单范围。

```bash
python -m pip install -r requirements.txt
```
## 跑测试

```bash
python -m pytest -v
```

实测输出（完整日志见 [docs/checks/pytest.log](docs/checks/pytest.log)）：

```
collected 18 items
18 passed in 0.80s
```

重新生成两个样例与期望输出快照、重新生成四份验收证据：

```bash
python tools/make_samples.py          # 写 samples/*.xlsx 与 samples/expected/*.json
python tools/make_samples.py --check  # 只校验已提交的快照是否与生成结果一致
python tools/emit_evidence.py         # 写 docs/checks/{merged,truncate,hidden,blocks}.txt
```

## 用法

```python
from kb.parsers import parse_file

blocks = parse_file("samples/pricelist.xlsx")
for block in blocks:
    if block.kind == "heading":
        print("#" * block.level, block.text)      # 每个工作表名，level 恒为 1
    elif block.kind == "table":
        print(block.rows[0])                     # rows[0] 恒为表头
        for row in block.rows[1:]:
            print(row)
    else:
        print(block.text)                        # 例如「已截断，共 25 行」
```

`Block` 的字段：`kind`（`heading` / `table` / `paragraph`）、`level`、`text`、`rows`、`sheet`，
另有 `to_dict()` 用于快照比对。也可以直接用 `kb.parsers.xlsx.parse_xlsx(path)`，
或用注册表 `kb.parsers.base.get_parser(".xlsx")` 取同一个函数。
## 解析规则（与四条验收逐条对应）

1. **合并单元格展开**：区域内左上角的值填满整块区域。实现上一半靠 openpyxl、一半靠自己：
   `read_only=True` 的 `ReadOnlyWorksheet` 没有 `merged_cells` 属性（openpyxl 在只读模式下
   不解析 `<mergeCells>`），所以 `kb/parsers/_xlsx_meta.py` 用 `zipfile` + `ElementTree`
   直接从工作表 XML 里读 `mergeCell` 的区间；展开在切块之前完成，因此每个块拿到的都是填满的行。
2. **行数截断**：每个工作表最多输出 `KB_XLSX_MAX_ROWS`（默认 2000）行数据（不含表头）。
   超出时该表末尾追加一个 `paragraph`，文本为「已截断，共 N 行」，N 是该表的实际数据行总数，
   不是已读行数。上限在每次解析时读环境变量，运行期可调，例如在测试里设置
   `KB_XLSX_MAX_ROWS=10` 后再解析。
3. **隐藏工作表跳过**：`sheet_state` 显式比对 `"visible"`，`hidden` 与 `veryHidden` 都不产生
   标题、不进入行数统计，内容也不会出现在任何块里。
4. **每 50 行一个表格块，每块重复表头**：表头 = 首个非空行（任意单元格有值的行）；
   `rows[0]` 恒为表头，数据行按 50 行切块，120 个数据行就是 3 个块（50 + 50 + 20）。

补充约定：表头去掉首尾空白；空表头列命名为 `column_{列号}`（列号从 1 开始）；
表头之前和之后的完全空行跳过且不计数；空工作表只产出一个标题块。

## 公式取缓存值（重要限制）

`parse_xlsx` 用 `load_workbook(..., read_only=True, data_only=True)`。
openpyxl 不会计算公式：`data_only=True` 读的是 Excel 最后一次保存时写进文件的结果缓存。所以：

* 文件被 Excel / WPS 打开并保存过，公式单元格有缓存值，解析结果里就是那个数字；
* 文件由程序生成、从没被 Excel 保存过，公式单元格没有缓存值，openpyxl 返回空值 `None`；
  解析器不会编造数值，而是往 `warnings` 里记一条「工作表 … 的公式单元格 … 没有缓存值，已按空值处理」。

本仓库的样例是真实带缓存值的文件：`tools/make_samples.py` 先用 openpyxl 写公式，
再用 `zipfile` + `ElementTree` 重写工作表 XML，把算好的结果写进 `<v>`，效果等同 Excel 保存过一次。
`tests/test_formula_cache.py::test_sample_pricelist_really_has_formula_and_cache` 同时断言了两面：
`data_only=False` 读到公式字符串 `=C2*D2`（说明文件里确实是公式，不是硬编码的数值），
`data_only=True` 读到 `1200`。没有缓存值的场景由 `test_formula_without_cache_is_reported` 覆盖：
断言读到 `None` 且恰好产生 1 条 warning，不跳过、也不假装通过。
## 样例与期望输出

| 文件 | 内容 |
| --- | --- |
| `samples/pricelist.xlsx` | 虚构价目表「价目表」，30 个数据行；`F2:F3` 纵向合并 2 行（区域列），`C12:E12` 横向合并 3 格（组合优惠说明行），金额列是公式 `=单价*数量`，带缓存值 |
| `samples/meeting_rooms.xlsx` | 虚构会议室规格表「会议室规格」，120 个数据行（3 个表格块）；`C2:C4` 纵向合并；另有 hidden 表「内部备注」与 veryHidden 表「归档」 |
| `samples/expected/pricelist.json` | 价目表的期望输出快照（逐块全等比较） |
| `samples/expected/meeting_rooms.json` | 会议室规格表的期望输出快照（逐块全等比较） |

快照结构为 `{"source", "parser", "max_rows", "blocks"}`，`blocks` 是 `Block.to_dict()` 的列表。
测试用 `max_rows=2000`（即默认值）显式解析后与 `blocks` 全等比较，避免开发机上的环境变量影响快照。

## 目录结构

```
kb/parsers/base.py         Block 数据类 + 解析器注册表（register / get_parser / parse_file）
kb/parsers/_xlsx_meta.py   xlsx 包内 XML 读取：合并区间、公式单元格、工作表与 part 的对应
kb/parsers/xlsx.py         parse_xlsx()：只读遍历、表头识别、合并展开、切块、截断提示
tools/make_samples.py      生成两个虚构样例与期望输出快照（--check 只校验）
tools/emit_evidence.py     生成 docs/checks/ 下四份验收证据
samples/                   两个虚构样例 xlsx 与 expected/ 快照
tests/                     6 个测试文件、18 个用例
docs/checks/               每条验收一段真实输出 + pytest 完整日志
```

## 测试清单（18 个用例）

| 文件 | 用例数 | 覆盖 |
| --- | --- | --- |
| `tests/test_merged_cells.py` | 3 | 横向 3 格、纵向 2 行、样例价目表的两处合并 |
| `tests/test_truncation.py` | 3 | 上限设为 10 行 + 25 行样例，输出 10 行且提示「已截断，共 25 行」；默认 2000 行不截断 |
| `tests/test_hidden_sheet.py` | 2 | hidden 与 veryHidden 的表名不出现在任何标题里、内容不泄漏；样例会议室表同一断言 |
| `tests/test_block_header.py` | 3 | 120 行切成 3 个块（50/50/20），每块首行等于表头；样例会议室表同一断言 |
| `tests/test_formula_cache.py` | 3 | 文件里真有公式且有缓存值、解析结果取缓存值、无缓存时返回空值并记 warning |
| `tests/test_samples_snapshot.py` | 4 | 两个样例与快照全等、生成器可复现快照、注册表按后缀分发 |
## 与投标提案的差异（以验收原文为准）

| 提案写法 | 实际交付 | 原因 |
| --- | --- | --- |
| 样例放 `fixtures/` | 放 `samples/`，期望输出在 `samples/expected/` | 样例与期望输出由本仓库自有布局承载，`tools/make_samples.py` 可重生成 |
| 价目表 2103 数据行，`A1:C2` 与 `D5:D9` 合并 | 30 数据行，`F2:F3`（纵向 2 行）与 `C12:E12`（横向 3 格）合并 | `A1:C2` 会把标题行当表头，与「首个非空行是表头」冲突；改为数据区内的合并，语义更清楚。大表截断场景由截断用例用 25 行 + 上限 10 的等价形式覆盖 |
| 会议室规格表 117 数据行 | 120 数据行 | 120 = 50 + 50 + 20，正好验证「每 50 行一块」的边界 |
| 合并展开函数接收工作表、读 `ws.merged_cells` | `_expand_merged(rows, merges, width)` | 只读模式的 `ReadOnlyWorksheet` 没有 `merged_cells`，改为先从 XML 读区间再传入 |
| 隐藏表只测 `hidden` | `hidden` 与 `veryHidden` 都测 | `sheet_state` 有三种取值，只测一种会漏 |

## 已知边界

* 只解析 `.xlsx`；`.xls`、`.xlsm`、加密工作簿不在本单范围。
* 公式只用缓存值，不重算；没有缓存值时按空值处理并记 warning。
* 合并区域若跨过行数上限的截断边界，被截掉的行没有任何输出，锚点行保留锚点值。
* 空工作表只输出标题块，不输出空表格块。
* 不解析样式、批注、图表、数据透视表。
