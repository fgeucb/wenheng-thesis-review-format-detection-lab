"""页边距比较（规则 025）：**一个方向一条结论**。

旧版 Word 有效格式流程会把四个方向的差异合并成一条批注，这里按细粒度规则拆开：
025.1 上、025.2 下、025.3 左、025.4 右。
"""

from __future__ import annotations

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.models import Finding, MarginSpec, Severity, Status, build_comment

LABELS = {"top": "上", "bottom": "下", "left": "左", "right": "右"}

#: 边距方向 → 细粒度规则编号
MARGIN_CODES = {"top": "025.1", "bottom": "025.2", "left": "025.3", "right": "025.4"}


def compare_margin(
    expected: MarginSpec, actual: MarginSpec, section_index: int, tolerance_mm: float = 0.2
) -> list[Finding]:
    """比较第 ``section_index`` 节的四向页边距，返回每个不符方向的一条结论。"""

    findings: list[Finding] = []
    for field, label in LABELS.items():
        expected_value = getattr(expected, field)
        actual_value = getattr(actual, field)
        if abs(expected_value - actual_value) <= tolerance_mm:
            continue
        code = MARGIN_CODES[field]
        title = FINE_RULES[code].title
        expected_text = f"{label}边距 {expected_value:.2f} mm"
        actual_text = f"{label}边距 {actual_value:.2f} mm"
        finding = Finding(code, title, Status.FAIL, "", severity=Severity.ERROR,
                          expected=expected_text, actual=actual_text, section_index=section_index)
        finding.message = build_comment(Severity.ERROR, code, title, expected_text, actual_text,
                                        location=f"第 {section_index} 节")
        findings.append(finding)
    return findings
