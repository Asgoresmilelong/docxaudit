"""数据模型：把 Word 文档还原成「段落 / 样式 / 字体」的纯 Python 结构。

刻意不依赖 python-docx 的高层 API —— 它对 pPr / rPr 的暴露太少，
而格式体检恰恰要看这些底层属性。
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class Run:
    """一个 w:r，携带解析后的有效字体信息。"""

    text: str
    east_asia: str | None = None
    latin: str | None = None
    size_pt: float | None = None
    bold: bool = False
    italic: bool = False

    @property
    def has_cjk(self) -> bool:
        return any("\u4e00" <= ch <= "\u9fff" for ch in self.text)


@dataclass
class Paragraph:
    no: int  # 文档顺序中的序号，1 开始
    text: str
    style_id: str | None = None
    style_name: str | None = None
    outline_level: int | None = None  # 0 表示一级标题，与 OOXML 一致
    in_table: bool = False
    alignment: str | None = None
    first_line_indent: int | None = None  # twips
    first_line_chars: int | None = None  # 1/100 字符
    line_spacing: int | None = None  # twips 或 240 分之一行
    line_rule: str | None = None
    field_codes: list[str] = field(default_factory=list)
    runs: list[Run] = field(default_factory=list)

    @property
    def stripped(self) -> str:
        return self.text.strip()

    @property
    def heading_level(self) -> int | None:
        """标题层级：优先 outlineLvl，其次按样式名推断。返回 0 基。"""
        if self.outline_level is not None and 0 <= self.outline_level <= 8:
            return self.outline_level
        name = (self.style_name or "").strip()
        if not name:
            return None
        import re

        m = re.match(r"^(?:heading|标题)\s*([1-9])", name, re.I)
        if m:
            return int(m.group(1)) - 1
        m = re.match(r"^([1-9])\s*级标题", name)
        if m:
            return int(m.group(1)) - 1
        return None


@dataclass
class Style:
    style_id: str
    name: str | None = None
    based_on: str | None = None
    outline_level: int | None = None
    east_asia: str | None = None
    latin: str | None = None
    size_pt: float | None = None
    type: str | None = None


@dataclass
class Document:
    path: str
    paragraphs: list[Paragraph] = field(default_factory=list)
    styles: dict[str, Style] = field(default_factory=dict)
    default_east_asia: str | None = None
    default_latin: str | None = None
    default_size_pt: float | None = None

    def style_chain(self, style_id: str | None, limit: int = 12) -> list[Style]:
        """按 basedOn 向上回溯的样式链（不含 None）。"""
        out: list[Style] = []
        seen: set[str] = set()
        cur = style_id
        while cur and cur in self.styles and cur not in seen and len(out) < limit:
            seen.add(cur)
            st = self.styles[cur]
            out.append(st)
            cur = st.based_on
        return out
