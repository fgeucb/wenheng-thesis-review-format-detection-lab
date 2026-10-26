from __future__ import annotations

from thesis_format_checker.core.effective_format import SpanFormat
from thesis_format_checker.core.models import Finding, Status


def compare_font_sizes(
    spans: list[SpanFormat],
    expected_pt: float | None,
    paragraph_index: int,
    text: str,
    tolerance_pt: float = 0.1,
) -> list[Finding]:
    if expected_pt is None:
        return []
    findings: list[Finding] = []
    for span in spans:
        if span.size_pt is None or abs(span.size_pt - expected_pt) <= tolerance_pt:
            continue
        findings.append(
            Finding(
                rule_id="069",
                title="正文字号",
                status=Status.FAIL,
                message=(
                    f"【069 正文字号】模板 A：{expected_pt:g} 磅；当前文字：{span.size_pt:g} 磅。"
                    "请核对本处正文字号。"
                ),
                expected=expected_pt,
                actual=span.size_pt,
                paragraph_index=paragraph_index,
                start=span.start,
                end=span.end,
                excerpt=text[span.start : span.end],
            )
        )
    return findings

