"""命令行入口：docxaudit check report.docx"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .loader import load_docx
from .report import print_report
from .rules import RULES, audit


def _split_ids(raw: str | None) -> list[str] | None:
    if not raw:
        return None
    ids = [s.strip().upper() for s in raw.replace("，", ",").split(",") if s.strip()]
    return ids or None


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="docxaudit",
        description="Word 文档格式体检：标题层级、题注编号、交叉引用、参考文献、字体与行距。",
    )
    parser.add_argument("--version", action="version", version=f"docxaudit {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    check = sub.add_parser("check", help="检查一个 .docx")
    check.add_argument("path", help="要检查的 .docx 文件")
    check.add_argument("--json", action="store_true", help="输出 JSON（便于接 CI / 二次处理）")
    check.add_argument("--only", metavar="R1,R2", help="只跑这些规则")
    check.add_argument("--ignore", metavar="R5,R6", help="跳过这些规则")
    check.add_argument("--max-errors", type=int, default=0, help="错误数超过该值才返回非零退出码，默认 0")
    check.add_argument("--no-fail", action="store_true", help="无论结果如何都返回 0")

    sub.add_parser("rules", help="列出全部规则")
    return parser


def cmd_rules() -> int:
    print("可用规则：")
    for rule in RULES:
        print(f"  {rule.id}  {rule.name:<6}  {rule.description}")
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    path = Path(args.path)
    if not path.exists():
        print(f"找不到文件：{path}", file=sys.stderr)
        return 2
    if path.suffix.lower() not in {".docx", ".docm"}:
        print(f"只支持 .docx：{path}（.doc 请先用 Word 另存为 .docx）", file=sys.stderr)
        return 2

    try:
        doc = load_docx(path)
    except Exception as exc:  # 损坏的包 / 加密文档
        print(f"无法读取 {path}：{exc}", file=sys.stderr)
        return 2

    findings = audit(doc, only=_split_ids(args.only), ignore=_split_ids(args.ignore))
    print_report(findings, source=path.name, as_json=args.json)

    if args.no_fail:
        return 0
    errors = sum(1 for f in findings if f.severity == "error")
    return 1 if errors > args.max_errors else 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "rules":
        return cmd_rules()
    return cmd_check(args)


if __name__ == "__main__":
    raise SystemExit(main())
