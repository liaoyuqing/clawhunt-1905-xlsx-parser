"""知识库文档解析器集合。

导入本包即完成 ``.xlsx`` 解析器的注册。
"""

from .base import (
    Block,
    get_parser,
    parse_file,
    register,
    supported_suffixes,
)
from .xlsx import parse_xlsx

__all__ = [
    "Block",
    "get_parser",
    "parse_file",
    "parse_xlsx",
    "register",
    "supported_suffixes",
]
