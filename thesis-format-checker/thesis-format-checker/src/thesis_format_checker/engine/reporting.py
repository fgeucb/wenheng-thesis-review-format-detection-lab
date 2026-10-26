"""报告输出：命令行摘要、规则清单、JSON 报告。

所有面向用户的输出都集中在这里，便于将来增加 HTML/Excel 报告而不影响检测逻辑。

* :func:`print_summary` —— 检测完成后的中文摘要（含错误/警告/提醒三档统计）；
* :func:`print_rules`   —— ``thesis-checker rules`` 打印规则族与细粒度规则；
* :func:`write_json_report` —— 把结论导出为机器可读的 JSON（可选）。
"""

from __future__ import annotations

import json
from pathlib import Path

from thesis_format_checker.checks.catalog import (
    CATEGORIES,
    STATUS_LABELS,
    fine_rules_of_category,
    rules_of_category,
)
from thesis_format_checker.core.models import (
    SEVERITY_LABELS,
    Finding,
    RunSummary,
    Severity,
    Status,
)
from thesis_format_checker.core.word_session import WordAutomationError

#: 命令行摘要里逐条列出的规则（各阶段主要结论，保持既有输出习惯）。
SUMMARY_RULES: tuple[tuple[str, str], ...] = (
    ("023", "纸张大小 A4"),
    ("025", "页面边距"),
    ("035", "页眉文字"),
    ("045", "一级标题"),
    ("046", "二级标题"),
    ("047", "三级及以下标题"),
    ("067", "正文中文字体"),
    ("068", "正文英文及数字字体"),
    ("069", "正文字号"),
    ("072", "正文首行缩进"),
    ("074", "正文行距"),
    ("075", "段前段后间距及空行"),
    ("122", "中英文标点混用"),
    ("123", "标点重复或不成对"),
)


def print_summary(summary: RunSummary) -> None:
    """打印检测摘要：节数、正文段数、各规则问题数、等级统计、无法判定与跳过项、输出路径。"""

    print("检测完成")
    print(f"已检查节数：{summary.sections_checked}")
    print(f"已检查正文段落：{summary.body_paragraphs_checked}")
    counts = summary.counts_by_rule()
    for rule_id, title in SUMMARY_RULES:
        print(f"【{rule_id} {title}】发现 {counts.get(rule_id, 0)} 处")
    severity_counts = summary.counts_by_severity()
    total = sum(severity_counts.values())
    print(f"问题总数：{total}（错误 {severity_counts.get(Severity.ERROR, 0)} 处、"
          f"警告 {severity_counts.get(Severity.WARNING, 0)} 处、"
          f"提醒 {severity_counts.get(Severity.HINT, 0)} 处）")
    print(f"无法判定：{len(summary.unknowns)} 项")
    for reason in summary.unknowns:
        print(f"  - {reason}")
    print(f"模板未涉及（已跳过）：{len(summary.not_applicable)} 项")
    print(f"仅报告不写批注（未确定或无法安全定位）：{len(summary.report_only)} 项")
    print(f"未实现/已删除：{len(summary.skipped)} 项")
    print(f"去重：同一段落重复合并 {summary.duplicates_removed} 条；"
          f"同类问题跨段落合并 {summary.merged_repeats} 条")
    cache = summary.cache_stats or {}
    if cache.get("enabled") and (cache.get("hits") or cache.get("misses")):
        hits = "、".join(cache.get("hits") or []) or "无"
        misses = "、".join(cache.get("misses") or []) or "无"
        print(f"缓存：命中 {hits}；未命中 {misses}")
        for note in cache.get("notes") or []:
            print(f"  - {note}")
    timings = summary.timings or {}
    if timings:
        # 分段耗时：让使用者一眼看出这次慢在 Word 阶段还是 docx 静态检测。
        order = ("Word段落属性与分页", "docx静态检测", "批注锚定与写入", "输出校验与落盘", "总耗时")
        parts = [f"{name} {timings[name]} 秒" for name in order if name in timings]
        parts += [f"{name} {value} 秒" for name, value in timings.items() if name not in order]
        print("耗时：" + "；".join(parts))
    print(f"输出文件：{summary.output_path}")


def print_rules(category: str | None = None, pending: bool = False) -> None:
    """打印规则目录：大类 → 规则族 → 细粒度规则（等级 / 状态 / 定位）。

    Args:
        category: 只看某个大类，可写 ``c07`` 或 ``7``。
        pending: 只列出尚未实现的细粒度规则。
    """

    wanted: str | None = None
    if category:
        key = category.strip().casefold()
        if key.isdigit():
            key = f"c{int(key):02d}"
        wanted = key
        if wanted not in {item.key for item in CATEGORIES}:
            raise WordAutomationError(f"未知大类：{category}；可用取值如 c01…c15")

    shown_total = 0
    for item in CATEGORIES:
        if wanted and item.key != wanted:
            continue
        fine = fine_rules_of_category(item.key)
        print(f"{item.key} 第{item.index}类 {item.name}：{len(fine)} 条细粒度规则")
        print(f"    {item.summary}")
        for family in rules_of_category(item.key):
            if family.status.value == "removed":
                print(f"  【{family.rule_id}】{family.title} —— {STATUS_LABELS[family.status]}")
                continue
            print(f"  【{family.rule_id}】{family.title} —— {STATUS_LABELS[family.status]}"
                  f"｜等级：{SEVERITY_LABELS[family.severity]}")
            for rule in family.items:
                if pending and rule.status.value in {"implemented", "aggregated", "legacy"}:
                    continue
                shown_total += 1
                print(f"      {rule.code} {rule.title}"
                      f"｜{SEVERITY_LABELS[rule.severity]}｜{STATUS_LABELS[rule.status]}"
                      f"｜定位：{rule.location}")
        print()
    if pending:
        print(f"以上为尚未实现的细粒度规则，共 {shown_total} 条。")


def _finding_payload(finding: Finding) -> dict:
    """把一条结论序列化为 JSON 友好的字典。"""

    severity = finding.severity or Severity.ERROR
    return {
        "rule_id": finding.rule_id,
        "title": finding.title,
        "severity": severity.value,
        "severity_label": SEVERITY_LABELS[severity],
        "status": finding.status.value,
        "location": finding.location,
        "expected": finding.expected,
        "actual": finding.actual,
        "message": finding.message,
        "paragraph_index": finding.paragraph_index,
        "section_index": finding.section_index,
        "start": finding.start,
        "end": finding.end,
        "excerpt": finding.excerpt,
        "page_story": finding.page_story,
        "header_kind": finding.header_kind,
        "page_variant": finding.page_variant,
        "page_scope": finding.page_scope,
        "page_config": finding.page_config,
        "page_issue": finding.page_issue,
        "page_skip_first": finding.page_skip_first,
        "affected_pages": list(finding.affected_pages),
        "merged_count": finding.merged_count,
        "merged_locations": list(finding.merged_locations),
    }


def _rule_payload() -> list[dict]:
    """规则目录的 JSON 形式（规则族 + 细粒度规则）。"""

    payload: list[dict] = []
    for category in CATEGORIES:
        for family in rules_of_category(category.key):
            payload.append({
                "family": family.rule_id,
                "category": category.key,
                "title": family.title,
                "status": family.status.value,
                "severity": family.severity.value,
                "aggregated": family.aggregated,
                "note": family.note,
                "items": [{
                    "code": item.code,
                    "title": item.title,
                    "severity": item.severity.value,
                    "status": item.status.value,
                    "location": item.location,
                } for item in family.items],
            })
    return payload


def write_json_report(summary: RunSummary, path: Path) -> Path:
    """导出 JSON 报告（检测摘要 + 逐条结论 + 等级 + 规则目录）。"""

    severity_counts = summary.counts_by_severity()
    payload = {
        "output_path": summary.output_path,
        "sections_checked": summary.sections_checked,
        "body_paragraphs_checked": summary.body_paragraphs_checked,
        "counts_by_rule": summary.counts_by_rule(),
        "counts_by_severity": {SEVERITY_LABELS[key]: value for key, value in severity_counts.items()},
        "unknowns": list(summary.unknowns),
        "not_applicable": list(summary.not_applicable),
        "report_only": list(summary.report_only),
        "skipped": list(summary.skipped),
        "duplicates_removed": summary.duplicates_removed,
        "merged_repeats": summary.merged_repeats,
        "cache": dict(summary.cache_stats or {}),
        "timings": dict(summary.timings or {}),
        "page_comments": list(summary.page_comments),
        "findings": [_finding_payload(finding) for finding in summary.findings],
        "rules": _rule_payload(),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


__all__ = ["SUMMARY_RULES", "print_rules", "print_summary", "write_json_report"]
