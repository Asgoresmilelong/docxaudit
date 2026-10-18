"""输出层：终端报告与 JSON。"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import asdict

from .rules import Finding

_ANSI = {
    "error": "\033[31m",
    "warn": "\033[33m",
    "info": "\033[36m",
    "dim": "\033[2m",
    "bold": "\033[1m",
    "reset": "\033[0m",
}

_LABEL = {"error": "错误", "warn": "警告", "info": "提示"}


def _supports_color(stream) -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    return bool(getattr(stream, "isatty", lambda: False)())


def format_text(findings: list[Finding], source: str = "", color: bool = False) -> str:
    def c(key: str, text: str) -> str:
        return f"{_ANSI[key]}{text}{_ANSI['reset']}" if color else text

    counts = {k: sum(1 for f in findings if f.severity == k) for k in _LABEL}
    lines: list[str] = []
    if source:
        lines.append(c("bold", f"docxaudit · {source}"))
    for f in findings:
        head = c(f.severity, f"[{f.rule} {_LABEL[f.severity]}]")
        loc = f"第 {f.para_no} 段" if f.para_no else "全文"
        lines.append(f"{head} {loc}：{f.message}")
        if f.snippet:
            lines.append(c("dim", f"    … {f.snippet}"))
        if f.hint:
            lines.append(c("dim", f"    → {f.hint}"))

    lines.append("")
    summary = (
        f"共 {len(findings)} 条：错误 {counts['error']} · 警告 {counts['warn']} · 提示 {counts['info']}"
        if findings
        else "未发现问题。"
    )
    lines.append(c("bold", summary))
    return "\n".join(lines)


def format_json(findings: list[Finding], source: str = "") -> str:
    payload = {
        "tool": "docxaudit",
        "source": source,
        "summary": {
            "total": len(findings),
            "error": sum(1 for f in findings if f.severity == "error"),
            "warn": sum(1 for f in findings if f.severity == "warn"),
            "info": sum(1 for f in findings if f.severity == "info"),
        },
        "findings": [asdict(f) for f in findings],
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def print_report(findings: list[Finding], source: str = "", as_json: bool = False, stream=None) -> None:
    stream = stream or sys.stdout
    if as_json:
        print(format_json(findings, source), file=stream)
    else:
        print(format_text(findings, source, color=_supports_color(stream)), file=stream)
