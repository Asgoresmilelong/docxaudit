"""docxaudit —— Word 文档格式体检。

    >>> from docxaudit import check
    >>> findings = check("论文.docx")
    >>> [f.message for f in findings if f.severity == "error"]
"""

from __future__ import annotations

from .loader import load_docx
from .model import Document, Paragraph, Run, Style
from .rules import RULES, Finding, audit

__version__ = "0.1.0"
__all__ = ["check", "audit", "load_docx", "RULES", "Finding", "Document", "Paragraph", "Run", "Style", "__version__"]


def check(path, only=None, ignore=None) -> list[Finding]:
    """加载并检查一个 .docx，返回按严重程度排序的问题列表。"""
    doc = load_docx(path)
    return audit(doc, only=only, ignore=ignore)
