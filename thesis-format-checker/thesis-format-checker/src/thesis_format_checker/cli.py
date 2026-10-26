"""命令行入口：``compare`` 检测文档、``rules`` 查看规则清单。

设计约定
========
* 参数校验与错误码保持既有习惯：参数错误由 argparse 返回 ``2``；
  运行/文档/配置错误返回 ``3``；检出格式问题仍然返回 ``0``。
* 终端输出集中在 :mod:`thesis_format_checker.engine.reporting`，
  本模块只负责解析参数与调用流程。
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

import yaml

from thesis_format_checker.engine.pipeline import compare_documents
from thesis_format_checker.engine.reporting import print_rules, print_summary, write_json_report
from thesis_format_checker.core.word_session import WordAutomationError


def load_mapping(path: Path | None) -> dict[str, Any] | None:
    """读取可选映射配置（YAML 或 JSON）。"""

    if path is None:
        return None
    if not path.is_file():
        raise WordAutomationError(f"映射配置不存在：{path}")
    try:
        text = path.read_text(encoding="utf-8")
        data = json.loads(text) if path.suffix.casefold() == ".json" else yaml.safe_load(text)
    except Exception as exc:
        raise WordAutomationError(f"无法解析映射配置：{path}") from exc
    if data is None:
        return {}
    if not isinstance(data, dict):
        raise WordAutomationError("映射配置的顶层必须是对象。")
    return data


def build_parser() -> argparse.ArgumentParser:
    """构造命令行解析器。"""

    parser = argparse.ArgumentParser(prog="thesis-checker", description="论文 Word 格式对照检测器")
    subparsers = parser.add_subparsers(dest="command", required=True)

    compare = subparsers.add_parser("compare", help="以模板 A 对照检测文档 B，并写入 Word 批注")
    compare.add_argument("--template", type=Path, required=True, help="模板 A（.doc/.docx）")
    compare.add_argument("--document", type=Path, required=True, help="待检测文档 B（.doc/.docx）")
    compare.add_argument("--output", type=Path, required=True, help="输出 .docx")
    compare.add_argument("--mapping", type=Path, help="可选 YAML/JSON 正文及节映射")
    compare.add_argument("--overwrite", action="store_true", help="显式允许覆盖已存在的输出文件")
    compare.add_argument("--margin-tolerance-mm", type=float, default=0.2)
    compare.add_argument("--font-size-tolerance-pt", type=float, default=0.1)
    compare.add_argument("--indent-tolerance-pt", type=float, default=0.5)
    compare.add_argument("--merge-repeats", type=int, default=0,
                         help="同一问题跨段落重复达到该次数时合并为一条汇总批注；0 表示不合并")
    compare.add_argument("--report-json", type=Path, help="可选：把检测结果另存为 JSON 报告")
    compare.add_argument("--show-comments", action="store_true",
                         help="在输出副本中打开“显示批注”开关（会修改 word/settings.xml，默认关闭）")
    compare.add_argument("--no-cache", action="store_true",
                         help="关闭结果缓存；默认复用上一次检测的 Word 段落属性与排版快照")
    compare.add_argument("--cache-dir", type=Path,
                         help="缓存目录（默认 %%LOCALAPPDATA%%\\thesis-checker\\cache）")

    rules = subparsers.add_parser("rules", help="列出 15 个大类、165 条规则及其实现状态")
    rules.add_argument("--category", help="只看某个大类，可写 c07 或 7")
    rules.add_argument("--pending", action="store_true", help="只列出本阶段尚未实现的规则")
    return parser


def main(argv: list[str] | None = None) -> int:
    """命令行主函数。"""

    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "compare":
            for name in ("margin_tolerance_mm", "font_size_tolerance_pt", "indent_tolerance_pt"):
                if getattr(args, name) < 0:
                    raise WordAutomationError("所有容差参数必须大于或等于 0。")
            if args.merge_repeats < 0:
                raise WordAutomationError("--merge-repeats 必须大于或等于 0。")
            summary = compare_documents(
                args.template,
                args.document,
                args.output,
                mapping=load_mapping(args.mapping),
                overwrite=args.overwrite,
                margin_tolerance_mm=args.margin_tolerance_mm,
                font_size_tolerance_pt=args.font_size_tolerance_pt,
                indent_tolerance_pt=args.indent_tolerance_pt,
                merge_repeats=args.merge_repeats,
                show_comments=args.show_comments,
                use_cache=not args.no_cache,
                cache_dir=args.cache_dir,
            )
            print_summary(summary)
            if args.report_json:
                print(f"JSON 报告：{write_json_report(summary, args.report_json)}")
            return 0
        if args.command == "rules":
            print_rules(args.category, args.pending)
            return 0
    except (WordAutomationError, OSError, ValueError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 3
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
