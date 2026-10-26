from __future__ import annotations

from thesis_format_checker.core.effective_format import SpanFormat
from thesis_format_checker.core.models import Finding, Status


def compare_fonts(
    spans: list[SpanFormat],
    expected_chinese: str | None,
    expected_latin: str | None,
    paragraph_index: int,
    text: str,
) -> list[Finding]:
    findings: list[Finding] = []
    for span in spans:
        expected = expected_chinese if span.script == "chinese" else expected_latin
        if expected is None or span.font_name is None:
            continue
        if span.font_name.casefold() == expected.casefold():
            continue
        rule_id = "067" if span.script == "chinese" else "068"
        title = "正文中文字体" if span.script == "chinese" else "正文英文及数字字体"
        excerpt = text[span.start : span.end]
        findings.append(
            Finding(
                rule_id=rule_id,
                title=title,
                status=Status.FAIL,
                message=(
                    f"【{rule_id} {title}】模板 A：{expected}；当前文字：{span.font_name}。"
                    "请核对本处正文字体。"
                ),
                expected=expected,
                actual=span.font_name,
                paragraph_index=paragraph_index,
                start=span.start,
                end=span.end,
                excerpt=excerpt,
            )
        )
    return findings

