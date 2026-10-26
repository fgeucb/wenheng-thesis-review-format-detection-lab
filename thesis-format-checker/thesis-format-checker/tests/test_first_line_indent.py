from thesis_format_checker.core.models import IndentSpec
from thesis_format_checker.detectors.word_effective.first_line_indent import compare_first_line_indent

from conftest import FakeParagraph


def family_of(rule_id):
    """细粒度规则编号 → 所属规则族（036.1 → 036）。"""

    return rule_id.split(".")[0]



def test_character_indent_passes():
    paragraph = FakeParagraph(points=21, characters=2)
    assert compare_first_line_indent(paragraph, "正文内容", IndentSpec(21, 2), 1) == []


def test_indent_mismatch_and_manual_spaces_are_both_reported():
    paragraph = FakeParagraph(points=0, characters=0)
    findings = compare_first_line_indent(paragraph, "\u3000\u3000正文", IndentSpec(21, 2), 4)
    assert len(findings) == 2
    assert all(family_of(item.rule_id) == "072" for item in findings)
    assert "2 个全角空格" in findings[0].message


def test_tab_is_reported():
    findings = compare_first_line_indent(FakeParagraph(21, 2), "\t正文", IndentSpec(21, 2), 1)
    assert len(findings) == 1 and "Tab" in findings[0].message
