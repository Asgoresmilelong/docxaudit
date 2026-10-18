"""规则引擎：每条规则只看 Document，产出 Finding，不做任何 I/O。

命名约定：R1..R6 稳定不变，便于在 CI 里 --ignore R5 这种写法长期有效。
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Callable, Iterable

from .model import Document, Paragraph

SEVERITY_ORDER = {"error": 0, "warn": 1, "info": 2}


@dataclass
class Finding:
    rule: str
    severity: str  # error | warn | info
    message: str
    para_no: int | None = None
    snippet: str = ""
    hint: str = ""

    @property
    def line(self) -> str:
        loc = f"第 {self.para_no} 段" if self.para_no else "全文"
        return f"{loc}：{self.message}"


@dataclass
class Rule:
    id: str
    name: str
    description: str
    run: Callable[[Document], list[Finding]]


def snippet(text: str, width: int = 36) -> str:
    t = re.sub(r"\s+", " ", text).strip()
    return t if len(t) <= width else t[: width - 1] + "…"


# ---------------------------------------------------------------- 公用解析


CAPTION_RE = re.compile(
    r"^\s*(?P<kind>图|表|Figure|Table|Fig\.)\s*(?P<num>[0-9]+(?:\s*[-–—.．]\s*[0-9]+)*)",
    re.I,
)
CITE_RE = re.compile(r"\[([0-9][0-9,，\-\–—\s]*)\]")
ENTRY_NO_RE = re.compile(r"^\s*[\[\(（【]?\s*([0-9]{1,3})\s*[\]\)）】]?\s*[.、，,]?")

CAPTION_STYLE_HINTS = ("caption", "题注")
TOC_STYLE_HINTS = ("toc", "目录")


def norm_num(raw: str) -> str:
    raw = re.sub(r"\s+", "", raw)
    return re.sub(r"[-–—.．]", "-", raw)


def is_caption(p: Paragraph) -> bool:
    if p.heading_level is not None:
        return False
    m = CAPTION_RE.match(p.stripped)
    if not m:
        return False
    # 题注通常较短；长段落里出现“图 3 显示了……”不算题注
    style = (p.style_name or "").lower()
    if any(h in style for h in CAPTION_STYLE_HINTS):
        return True
    return len(p.stripped) <= 60 and not p.stripped.startswith(("图表", "表中", "图中"))


def captions(doc: Document) -> list[tuple[str, str, Paragraph]]:
    out = []
    for p in doc.paragraphs:
        if not is_caption(p):
            continue
        m = CAPTION_RE.match(p.stripped)
        if not m:
            continue
        kind = "图" if m.group("kind").lower().startswith(("图", "fig")) else "表"
        out.append((kind, norm_num(m.group("num")), p))
    return out


def is_toc_like(p: Paragraph) -> bool:
    style = (p.style_name or "").lower()
    return any(h in style for h in TOC_STYLE_HINTS)


def body_paragraphs(doc: Document, min_chars: int = 20) -> Iterable[Paragraph]:
    """正文段落：排除表格、标题、题注、目录，以及参考文献区。

    参考文献区有自己的排版规范（悬挂缩进、字号更小），把它算进“正文多数派”
    会把真正的正文偏差稀释掉。
    """
    excluded = _reference_para_nos(doc)
    for p in doc.paragraphs:
        if p.no in excluded:
            continue
        if p.in_table or p.heading_level is not None or is_caption(p) or is_toc_like(p):
            continue
        if len(p.stripped) < min_chars:
            continue
        yield p


# ---------------------------------------------------------------- R1 标题层级


def _style_heading_level(p: Paragraph) -> int | None:
    return p.heading_level


def rule_heading_sequence(doc: Document) -> list[Finding]:
    findings: list[Finding] = []
    headings = [(p, lv) for p in doc.paragraphs if (lv := _style_heading_level(p)) is not None]
    if not headings:
        return findings

    first_p, first_lv = headings[0]
    if first_lv != 0:
        findings.append(
            Finding(
                "R1",
                "warn",
                f"全文第一个标题是 {first_lv + 1} 级，通常应从 1 级标题开始",
                first_p.no,
                snippet(first_p.stripped),
                "检查标题样式是否被误用（例如正文误套用“标题 2”）。",
            )
        )

    prev_lv = first_lv
    for p, lv in headings[1:]:
        if lv > prev_lv + 1:
            findings.append(
                Finding(
                    "R1",
                    "error",
                    f"标题层级跳跃：从 {prev_lv + 1} 级直接跳到 {lv + 1} 级",
                    p.no,
                    snippet(p.stripped),
                    "中间缺一级标题会导致导航窗格和自动目录结构错乱。",
                )
            )
        prev_lv = lv
    return findings


# ---------------------------------------------------------------- R2 题注编号


def rule_caption_numbering(doc: Document) -> list[Finding]:
    findings: list[Finding] = []
    caps = captions(doc)
    if not caps:
        return findings

    seen: dict[tuple[str, str], Paragraph] = {}
    for kind, num, p in caps:
        key = (kind, num)
        if key in seen:
            findings.append(
                Finding(
                    "R2",
                    "error",
                    f"{kind} {num} 重复编号（与第 {seen[key].no} 段重复）",
                    p.no,
                    snippet(p.stripped),
                    "重复编号通常来自复制粘贴后没有更新 SEQ 域（F9 更新域）。",
                )
            )
        else:
            seen[key] = p

    groups: dict[tuple[str, str], list[tuple[str, Paragraph]]] = {}
    for kind, num, p in caps:
        parts = num.split("-")
        prefix = "-".join(parts[:-1])
        if parts[-1].isdigit():
            groups.setdefault((kind, prefix), []).append((parts[-1], p))

    for (kind, prefix), items in groups.items():
        expected = None
        for raw, p in items:
            value = int(raw)
            if expected is not None and value != expected:
                label = f"{prefix}-{raw}" if prefix else raw
                want = f"{prefix}-{expected}" if prefix else str(expected)
                findings.append(
                    Finding(
                        "R2",
                        "warn",
                        f"{kind} 编号不连续：出现 {kind} {label}，按顺序应为 {kind} {want}",
                        p.no,
                        snippet(p.stripped),
                        "手动改过编号，或删图后没有更新域。",
                    )
                )
            expected = value + 1
    return findings


# ---------------------------------------------------------------- R3 交叉引用


def _mentioned_numbers(doc: Document) -> list[tuple[str, str, Paragraph]]:
    out = []
    for p in doc.paragraphs:
        if is_caption(p) or is_toc_like(p):
            continue
        for m in re.finditer(r"(图|表)\s*([0-9]+(?:\s*[-–—.．]\s*[0-9]+)*)", p.stripped):
            kind = m.group(1)
            out.append((kind, norm_num(m.group(2)), p))
    return out


def rule_cross_references(doc: Document) -> list[Finding]:
    findings: list[Finding] = []
    caps = captions(doc)
    if not caps:
        return findings

    full: dict[str, set[str]] = {"图": set(), "表": set()}
    tails: dict[str, set[str]] = {"图": set(), "表": set()}
    for kind, num, _ in caps:
        full[kind].add(num)
        tails[kind].add(num.split("-")[-1])

    missing: dict[tuple[str, str], Paragraph] = {}
    for kind, num, p in _mentioned_numbers(doc):
        if num in full[kind] or num.split("-")[-1] in tails[kind]:
            continue
        missing.setdefault((kind, num), p)

    for (kind, num), p in sorted(missing.items()):
        findings.append(
            Finding(
                "R3",
                "error",
                f"正文引用了「{kind} {num}」，但全文没有这个编号的题注",
                p.no,
                snippet(p.stripped),
                "图号写错，或题注被删除后正文未同步修改。",
            )
        )
    return findings


# ---------------------------------------------------------------- R4 参考文献


def _reference_para_nos(doc: Document) -> set[int]:
    """参考文献区覆盖的段落号（含标题），用于把这片区域从“正文”里排除。"""
    head, entries = _reference_section(doc)
    if head is None:
        return set()
    return {head.no} | {p.no for p in entries}


def _reference_section(doc: Document) -> tuple[Paragraph | None, list[Paragraph]]:
    ref_head = None
    for p in doc.paragraphs:
        lv = p.heading_level
        if lv is None:
            continue
        if re.search(r"参考文献|references", p.stripped, re.I):
            ref_head = p
            break
    if ref_head is None:
        return None, []

    entries: list[Paragraph] = []
    for p in doc.paragraphs:
        if p.no <= ref_head.no:
            continue
        lv = p.heading_level
        if lv is not None and lv <= ref_head.heading_level:
            break
        if p.stripped:
            entries.append(p)
    return ref_head, entries


def _expand_citation(raw: str) -> list[int]:
    nums: list[int] = []
    for part in re.split(r"[,，]", raw):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d+)\s*[-–—]\s*(\d+)$", part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= b and b - a <= 200:
                nums.extend(range(a, b + 1))
            continue
        if part.isdigit():
            nums.append(int(part))
    return nums


def rule_reference_list(doc: Document) -> list[Finding]:
    findings: list[Finding] = []
    head, entries = _reference_section(doc)
    if head is None or not entries:
        return findings

    listed: dict[int, Paragraph] = {}
    for p in entries:
        m = ENTRY_NO_RE.match(p.stripped)
        if m:
            listed.setdefault(int(m.group(1)), p)

    cited: dict[int, Paragraph] = {}
    for p in doc.paragraphs:
        if p.no >= head.no or is_toc_like(p):
            continue
        for m in CITE_RE.finditer(p.stripped):
            for n in _expand_citation(m.group(1)):
                cited.setdefault(n, p)

    for n in sorted(set(cited) - set(listed)):
        p = cited[n]
        findings.append(
            Finding(
                "R4",
                "error",
                f"正文引用了文献 [{n}]，但参考文献列表里没有这一条",
                p.no,
                snippet(p.stripped),
                "补条目，或发现引用编号写错。",
            )
        )

    uncited = sorted(set(listed) - set(cited))
    if len(listed) >= 2 and uncited:
        shown = "、".join(f"[{n}]" for n in uncited[:12])
        more = f" 等 {len(uncited)} 条" if len(uncited) > 12 else ""
        findings.append(
            Finding(
                "R4",
                "info",
                f"有 {len(uncited)} 条参考文献在正文中从未被引用：{shown}{more}",
                listed[uncited[0]].no,
                snippet(listed[uncited[0]].stripped),
                "多数学校要求文献表与正文引用一一对应；这些条目要么补引用，要么删掉。",
            )
        )
    return findings


# ---------------------------------------------------------------- R5 字体一致性


def _cjk_fonts(p: Paragraph) -> list[str]:
    return [r.east_asia or "（未设置）" for r in p.runs if r.has_cjk]


def _majority(values: list[str]) -> tuple[str | None, int]:
    if not values:
        return None, 0
    counts: dict[str, int] = {}
    for v in values:
        counts[v] = counts.get(v, 0) + 1
    value, count = max(counts.items(), key=lambda kv: kv[1])
    return value, count


def rule_font_consistency(doc: Document) -> list[Finding]:
    findings: list[Finding] = []
    paras = list(body_paragraphs(doc))

    mixed = [p for p in paras if len(set(_cjk_fonts(p))) > 1]
    for p in mixed[:6]:
        fonts = "、".join(sorted(set(_cjk_fonts(p))))
        findings.append(
            Finding(
                "R5",
                "warn",
                f"同一段内混用了多种中文字体：{fonts}",
                p.no,
                snippet(p.stripped),
                "常见于从网页或 PDF 粘贴文字，建议统一后清除格式。",
            )
        )

    votes = [(_cjk_fonts(p)[0] if _cjk_fonts(p) else None) for p in paras]
    votes = [v for v in votes if v]
    value, count = _majority(votes)
    if not value or count < 3:
        return findings

    ratio = count / len(votes)
    if ratio < 0.6:
        return findings

    deviants = [p for p, v in zip([p for p in paras if _cjk_fonts(p)], votes) if v != value]
    for p in deviants[:8]:
        actual = _cjk_fonts(p)[0]
        findings.append(
            Finding(
                "R5",
                "warn",
                f"正文中文字体应为「{value}」，此段为「{actual}」",
                p.no,
                snippet(p.stripped),
                "全文正文应统一字体；标题和题注可以不同。",
            )
        )
    if len(deviants) > 8:
        findings.append(
            Finding("R5", "warn", f"另有 {len(deviants) - 8} 段字体不一致，未逐条列出")
        )
    return findings


# ---------------------------------------------------------------- R6 行距与缩进


def _spacing_key(p: Paragraph) -> tuple:
    return (p.line_spacing, p.line_rule, p.first_line_indent, p.first_line_chars)


def _describe_key(key: tuple) -> str:
    line, rule, indent, indent_chars = key
    parts = []
    if line is None:
        parts.append("行距未显式设置")
    elif rule == "exact":
        parts.append(f"固定行距 {line / 20:.1f} 磅")
    elif rule == "auto":
        parts.append(f"多倍行距 {line / 240:.2f} 倍")
    else:
        parts.append(f"行距 {line / 20:.1f} 磅")
    if indent_chars:
        parts.append(f"首行缩进 {indent_chars / 100:.2f} 字符")
    elif indent:
        parts.append(f"首行缩进 {indent / 20:.1f} 磅")
    else:
        parts.append("无首行缩进")
    return "，".join(parts)


def rule_spacing_consistency(doc: Document) -> list[Finding]:
    findings: list[Finding] = []
    paras = list(body_paragraphs(doc))
    if len(paras) < 5:
        return findings

    keys = [(_describe_key(_spacing_key(p)), _spacing_key(p), p) for p in paras]
    labels = [k[0] for k in keys]
    value, count = _majority(labels)
    if not value or count / len(labels) < 0.6:
        return findings

    deviants = [(label, p) for label, _, p in keys if label != value]
    for label, p in deviants[:8]:
        findings.append(
            Finding(
                "R6",
                "warn",
                f"段落间距为「{label}」，多数正文为「{value}」",
                p.no,
                snippet(p.stripped),
                "行距/缩进不一致在打印稿上非常明显，且会影响页数统计。",
            )
        )
    if len(deviants) > 8:
        findings.append(
            Finding("R6", "warn", f"另有 {len(deviants) - 8} 段行距或缩进不一致，未逐条列出")
        )
    return findings


# ---------------------------------------------------------------- 注册表


RULES: list[Rule] = [
    Rule("R1", "标题层级", "标题是否从 1 级开始、层级是否跳跃", rule_heading_sequence),
    Rule("R2", "题注编号", "图/表编号是否连续、是否重复", rule_caption_numbering),
    Rule("R3", "交叉引用", "正文提到的图号/表号是否存在对应题注", rule_cross_references),
    Rule("R4", "参考文献", "正文引用与文献列表是否一一对应", rule_reference_list),
    Rule("R5", "中文字体", "正文中文字体是否统一、段内是否混用", rule_font_consistency),
    Rule("R6", "行距缩进", "正文行距与首行缩进是否统一", rule_spacing_consistency),
]

RULES_BY_ID = {r.id: r for r in RULES}


def audit(
    doc: Document,
    only: Iterable[str] | None = None,
    ignore: Iterable[str] | None = None,
) -> list[Finding]:
    only_set = {s.upper() for s in only} if only else None
    ignore_set = {s.upper() for s in ignore} if ignore else set()

    findings: list[Finding] = []
    for rule in RULES:
        if only_set is not None and rule.id not in only_set:
            continue
        if rule.id in ignore_set:
            continue
        findings.extend(rule.run(doc))

    findings.sort(key=lambda f: (SEVERITY_ORDER.get(f.severity, 9), f.para_no or 0, f.rule))
    return findings
