from __future__ import annotations

from thesis_format_checker.core.effective_format import normalized_number
from thesis_format_checker.core.models import Finding, Severity, SpacingSpec, Status, build_comment


LINE_RULE_LABELS = {
    0: "单倍行距",
    1: "1.5 倍行距",
    2: "双倍行距",
    3: "最小值行距",
    4: "固定值行距",
    5: "多倍行距",
}


def describe_line_spacing(spec: SpacingSpec, *, show_calculated_points: bool = False) -> str | None:
    label = LINE_RULE_LABELS.get(spec.line_rule)
    if label is None:
        return None
    if spec.line_rule in (3, 4):
        return f"{label} {spec.line_points:g} 磅"
    if spec.line_rule == 5:
        return f"{label}（Word 计算值约 {spec.line_points:g} 磅）"
    if show_calculated_points:
        return f"{label}（Word 计算行高约 {spec.line_points:g} 磅）"
    return label


def read_spacing(paragraph) -> SpacingSpec | None:
    fmt = paragraph.Format
    try:
        rule = int(fmt.LineSpacingRule)
        line = normalized_number(fmt.LineSpacing)
        before = normalized_number(fmt.SpaceBefore)
        after = normalized_number(fmt.SpaceAfter)
    except Exception:
        return None
    if None in (line, before, after):
        return None
    return SpacingSpec(rule, line, before, after)


def compare_spacing(expected: SpacingSpec, actual: SpacingSpec, index: int, tolerance_pt: float = 0.5) -> list[Finding]:
    result = []
    same_rule_different_height = expected.line_rule == actual.line_rule and abs(expected.line_points - actual.line_points) > tolerance_pt
    expected_line = describe_line_spacing(expected, show_calculated_points=same_rule_different_height)
    actual_line = describe_line_spacing(actual, show_calculated_points=same_rule_different_height)
    if expected_line is not None and actual_line is not None and (expected.line_rule != actual.line_rule or abs(expected.line_points - actual.line_points) > tolerance_pt):
        result.append(Finding(
            "074", "正文行距", Status.FAIL,
            f"【074 正文行距】模板 A：{expected_line}；当前文档：{actual_line}。",
            expected=expected, actual=actual, paragraph_index=index, start=0, end=1,
        ))
    if abs(expected.before - actual.before) > tolerance_pt or abs(expected.after - actual.after) > tolerance_pt:
        # 段前与段后分别成条（075.1 / 075.2）
        for code, title, want, current in (
            ("075.1", "段前间距不符合要求", expected.before, actual.before),
            ("075.2", "段后间距不符合要求", expected.after, actual.after),
        ):
            if abs(want - current) <= tolerance_pt:
                continue
            want_text, current_text = f"{want:g} 磅", f"{current:g} 磅"
            finding = Finding(code, title, Status.FAIL, "",
                              expected=want_text, actual=current_text,
                              paragraph_index=index, start=0, end=1)
            finding.message = build_comment(Severity.WARNING, code, title, want_text, current_text)
            result.append(finding)
    return result


def blank_paragraph_finding(index: int) -> Finding:
    title = "用空行代替段落间距"
    expected, actual = "正文段落之间不使用空行", "正文段落之间存在空段落"
    finding = Finding("075.3", title, Status.FAIL, "", expected=expected, actual=actual,
                      paragraph_index=index, start=0, end=1)
    finding.message = build_comment(Severity.WARNING, "075.3", title, expected, actual)
    return finding
