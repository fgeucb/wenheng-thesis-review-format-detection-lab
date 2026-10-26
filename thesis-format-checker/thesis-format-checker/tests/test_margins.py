from thesis_format_checker.core.models import MarginSpec
from thesis_format_checker.detectors.word_effective.margins import compare_margin


def test_equal_margins_pass():
    standard = MarginSpec(25.4, 25.4, 31.7, 31.7)
    assert compare_margin(standard, MarginSpec(25.41, 25.4, 31.7, 31.7), 1) == []


def test_reports_each_changed_direction_separately():
    """一个方向一条结论：上边距与右边距分别成条（025.1 / 025.4）。"""
    standard = MarginSpec(25, 25, 30, 30)
    findings = compare_margin(standard, MarginSpec(20, 25, 30, 35), 2)
    assert {item.rule_id for item in findings} == {"025.1", "025.4"}
    for item in findings:
        assert "第 2 节" in item.message
        assert "边距" in item.message


def test_multiple_sections_are_compared_independently():
    standard = MarginSpec(25, 25, 30, 30)
    results = [
        compare_margin(standard, MarginSpec(25, 25, 30, 30), 1),
        compare_margin(standard, MarginSpec(25, 25, 20, 30), 2),
    ]
    assert results[0] == []
    assert len(results[1]) == 1
    assert results[1][0].rule_id == "025.3" and results[1][0].section_index == 2
