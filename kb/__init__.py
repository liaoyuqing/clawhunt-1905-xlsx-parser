"""会员空间平台 · 知识库解析器包。

本包只包含纯 Python 的解析逻辑，不依赖 FastAPI / SQLAlchemy / Redis 等后端框架，
因此可以被入库流水线、命令行工具和测试独立引用。
"""

__all__ = ["parsers"]
