"""OOXML 读取层：把 .docx 解包，恢复段落顺序、有效样式与有效字体。

关键点（也是 python-docx 高层 API 覆盖不到的部分）：
- 段落属性可能写在段落上、样式上、样式基类上，需要逐层继承；
- 字体可能来自 run、来自样式、来自 docDefaults 的 rPrDefault；
- 域代码（w:instrText）里藏着 SEQ / REF，题注编号靠它自动生成。
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from lxml import etree

from .model import Document, Paragraph, Run, Style

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def qn(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _attr(el, name: str) -> str | None:
    if el is None:
        return None
    return el.get(qn(name))


def _toggle(el) -> bool | None:
    """w:b / w:i 这类开关属性：存在即 True，val=0/false 视为 False。"""
    if el is None:
        return None
    val = _attr(el, "val")
    if val is None:
        return True
    return val.lower() not in {"0", "false", "off"}


def _half_points(raw: str | None) -> float | None:
    if not raw:
        return None
    try:
        return int(float(raw)) / 2
    except ValueError:
        return None


def _theme_font(raw: str | None) -> str | None:
    """theme 字体（如 theme:minorEastAsia）在这里无法解析出实际字体名。"""
    if not raw:
        return None
    return None if raw.startswith("theme:") else raw


def _parse_styles(xml: bytes | None) -> tuple[dict[str, Style], dict[str, str | None]]:
    styles: dict[str, Style] = {}
    defaults: dict[str, str | None] = {}
    if not xml:
        return styles, defaults

    root = etree.fromstring(xml)

    docdefaults = root.find(qn("docDefaults"))
    if docdefaults is not None:
        rpr = docdefaults.find(f"{qn('rPrDefault')}/{qn('rPr')}")
        fonts = rpr.find(qn("rFonts")) if rpr is not None else None
        defaults["east_asia"] = _theme_font(_attr(fonts, "eastAsia"))
        defaults["latin"] = _theme_font(_attr(fonts, "ascii"))
        defaults["size_pt"] = _half_points(_attr(rpr.find(qn("sz")) if rpr is not None else None, "val"))

    for st in root.findall(qn("style")):
        sid = _attr(st, "styleId")
        if not sid:
            continue
        name_el = st.find(qn("name"))
        based_el = st.find(qn("basedOn"))
        ppr = st.find(qn("pPr"))
        rpr = st.find(qn("rPr"))
        fonts = rpr.find(qn("rFonts")) if rpr is not None else None
        sz = rpr.find(qn("sz")) if rpr is not None else None
        styles[sid] = Style(
            style_id=sid,
            name=_attr(name_el, "val"),
            based_on=_attr(based_el, "val"),
            outline_level=int(_attr(ppr.find(qn("outlineLvl")), "val"))
            if ppr is not None and ppr.find(qn("outlineLvl")) is not None
            else None,
            east_asia=_theme_font(_attr(fonts, "eastAsia")),
            latin=_theme_font(_attr(fonts, "ascii")),
            size_pt=_half_points(_attr(sz, "val")),
            type=_attr(st, "type"),
        )
    return styles, defaults


def _iter_paragraphs(container, in_table: bool = False):
    """按文档顺序产出 (w:p, in_table)，会下钻表格与内容控件（sdt）。"""
    for el in container:
        tag = el.tag
        if tag == qn("p"):
            yield el, in_table
        elif tag == qn("tbl"):
            for row in el.findall(qn("tr")):
                for cell in row.findall(qn("tc")):
                    yield from _iter_paragraphs(cell, True)
        elif tag == qn("sdt"):
            content = el.find(qn("sdtContent"))
            if content is not None:
                yield from _iter_paragraphs(content, in_table)
        elif tag in (qn("sdtContent"),):
            yield from _iter_paragraphs(el, in_table)


def _paragraph_text(p) -> str:
    parts: list[str] = []
    for node in p.iter():
        if node.tag == qn("t"):
            # 跳过修订记录里被删除的文字
            anc = node.getparent()
            deleted = False
            while anc is not None:
                if anc.tag == qn("del"):
                    deleted = True
                    break
                anc = anc.getparent()
            if not deleted and node.text:
                parts.append(node.text)
        elif node.tag == qn("tab"):
            parts.append("\t")
        elif node.tag in (qn("br"), qn("cr")):
            parts.append("\n")
    return "".join(parts)


def _paragraph_fields(p) -> list[str]:
    return [el.text.strip() for el in p.iter(qn("instrText")) if el.text and el.text.strip()]


def _resolve_font(
    run_rpr,
    style_chain,
    doc: Document,
) -> tuple[str | None, str | None, float | None]:
    east: str | None = None
    latin: str | None = None
    size: float | None = None

    fonts = run_rpr.find(qn("rFonts")) if run_rpr is not None else None
    east = _theme_font(_attr(fonts, "eastAsia")) or _theme_font(_attr(fonts, "hAnsi"))
    latin = _theme_font(_attr(fonts, "ascii"))
    if run_rpr is not None:
        size = _half_points(_attr(run_rpr.find(qn("sz")), "val"))

    for st in style_chain:
        if east is None and st.east_asia:
            east = st.east_asia
        if latin is None and st.latin:
            latin = st.latin
        if size is None and st.size_pt is not None:
            size = st.size_pt

    east = east or doc.default_east_asia
    latin = latin or doc.default_latin
    size = size if size is not None else doc.default_size_pt
    return east, latin, size


def load_docx(path: str | Path) -> Document:
    path = Path(path)
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        if "word/document.xml" not in names:
            raise ValueError(f"{path} 不是有效的 .docx（缺少 word/document.xml）")
        document_xml = zf.read("word/document.xml")
        styles_xml = zf.read("word/styles.xml") if "word/styles.xml" in names else None

    styles, defaults = _parse_styles(styles_xml)
    doc = Document(
        path=str(path),
        styles=styles,
        default_east_asia=defaults.get("east_asia"),
        default_latin=defaults.get("latin"),
        default_size_pt=defaults.get("size_pt"),
    )

    root = etree.fromstring(document_xml)
    body = root.find(qn("body"))
    if body is None:
        raise ValueError(f"{path} 的 document.xml 中没有 body")

    for idx, (p, in_table) in enumerate(_iter_paragraphs(body), start=1):
        ppr = p.find(qn("pPr"))
        style_id = _attr(ppr.find(qn("pStyle")), "val") if ppr is not None else None
        chain = doc.style_chain(style_id)
        style_name = chain[0].name if chain else None

        outline = None
        for st in chain:
            if st.outline_level is not None:
                outline = st.outline_level
                break
        if ppr is not None and ppr.find(qn("outlineLvl")) is not None:
            try:
                outline = int(_attr(ppr.find(qn("outlineLvl")), "val"))
            except (TypeError, ValueError):
                pass

        ind = ppr.find(qn("ind")) if ppr is not None else None
        spacing = ppr.find(qn("spacing")) if ppr is not None else None
        jc = ppr.find(qn("jc")) if ppr is not None else None

        runs: list[Run] = []
        for r in p.iter(qn("r")):
            rpr = r.find(qn("rPr"))
            text = "".join(t.text or "" for t in r.findall(qn("t")))
            if not text:
                continue
            east, latin, size = _resolve_font(rpr, chain, doc)
            runs.append(
                Run(
                    text=text,
                    east_asia=east,
                    latin=latin,
                    size_pt=size,
                    bold=bool(_toggle(rpr.find(qn("b"))) if rpr is not None else False),
                    italic=bool(_toggle(rpr.find(qn("i"))) if rpr is not None else False),
                )
            )

        doc.paragraphs.append(
            Paragraph(
                no=idx,
                text=_paragraph_text(p),
                style_id=style_id,
                style_name=style_name,
                outline_level=outline,
                in_table=in_table,
                alignment=_attr(jc, "val"),
                first_line_indent=_int_or_none(_attr(ind, "firstLine")),
                first_line_chars=_int_or_none(_attr(ind, "firstLineChars")),
                line_spacing=_int_or_none(_attr(spacing, "line")),
                line_rule=_attr(spacing, "lineRule"),
                field_codes=_paragraph_fields(p),
                runs=runs,
            )
        )

    return doc


def _int_or_none(raw: str | None) -> int | None:
    if raw is None:
        return None
    try:
        return int(float(raw))
    except ValueError:
        return None
