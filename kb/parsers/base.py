"""知识库文档块模型与解析器注册表。

工单 S3-KB-007 的 XLSX 解析器实现后注册到本注册表，
知识库入库流水线按文件后缀挑选解析器。
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

__all__ = [
    "Block",
    "Parser",
    "register",
    "get_parser",
    "supported_suffixes",
    "parse_file",
]

CellValue = Any
Row = tuple[CellValue, ...]

# 解析器统一签名：接收路径，可选接收 warnings 列表，返回文档块列表。
Parser = Callable[..., "list[Block]"]


@dataclass(frozen=True)
class Block:
    """解析器输出的一块内容。

    kind:
        - ``heading``：标题块，``level`` 是层级，工作表名固定为 level 1。
        - ``table``：表格块，``rows[0]`` 恒为表头行。
        - ``paragraph``：普通段落块，例如行数截断说明。
    """

    kind: str
    level: int = 0
    text: str = ""
    rows: tuple[Row, ...] = ()
    sheet: str = ""

    def to_dict(self) -> dict[str, Any]:
        """转成可 JSON 序列化的字典（用于期望输出快照比对）。"""
        return {
            "kind": self.kind,
            "level": self.level,
            "text": self.text,
            "rows": [list(row) for row in self.rows],
            "sheet": self.sheet,
        }


_REGISTRY: dict[str, Parser] = {}


def register(suffix: str, parser: Parser, *, overwrite: bool = False) -> None:
    """注册一个后缀解析器，例如 ``register(".xlsx", parse_xlsx)``。"""
    key = suffix.lower()
    if not key.startswith("."):
        key = "." + key
    if key in _REGISTRY and not overwrite:
        raise ValueError("解析器已注册: " + key)
    _REGISTRY[key] = parser


def get_parser(suffix: str) -> Parser:
    """按后缀取解析器；未注册时抛 KeyError。"""
    key = suffix.lower()
    if not key.startswith("."):
        key = "." + key
    try:
        return _REGISTRY[key]
    except KeyError:
        raise KeyError("未注册的解析器: " + key) from None


def supported_suffixes() -> list[str]:
    """已注册的后缀列表（升序）。"""
    return sorted(_REGISTRY)


def parse_file(path: str | Path, *, warnings: list[str] | None = None) -> list[Block]:
    """按文件后缀分发到对应解析器。"""
    target = Path(path)
    parser = get_parser(target.suffix)
    return parser(target, warnings=warnings)
