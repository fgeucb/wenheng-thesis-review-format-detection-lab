"""批注语言：不写规则编号、不暴露字段名/True/False/U+ 码位，空白值写“无”。"""

from __future__ import annotations

from docx import Document

from thesis_format_checker.core.models import Severity, build_comment
from thesis_format_checker.detectors.formats import (
    _display_key, _display_profile_value, _formula_label,
)
from thesis_format_checker.detectors.headers import _even_pages_text
from thesis_format_checker.detectors.pages import _story_text
from thesis_format_checker.detectors.word_effective.headers import compare_headers


def test_comment_head_has_no_rule_number():
    text = build_comment(Severity.ERROR, "037.2", "偶数页页眉内容或位置不符合模板", "无", "有")
    assert text.splitlines()[0] == "【错误】偶数页页眉内容或位置不符合模板"
    assert "037" not in text


def test_even_pages_text_is_plain_chinese():
    assert "分开设置" in _even_pages_text(True)
    assert "相同" in _even_pages_text(False)
    for text in (_even_pages_text(True), _even_pages_text(False)):
        assert "True" not in text and "False" not in text


def test_story_text_hides_raw_tuples():
    assert _story_text(("第一章绪论", False)) == "文字「第一章绪论」"
    assert _story_text(("", False)) == "无内容"
    assert _story_text(("", True)) == "图片或其他对象"
    assert _story_text(None) == "含动态域或无法读取"
    assert "False" not in _story_text(("第一章绪论", False))


def test_header_comparison_writes_none_instead_of_blank():
    findings = compare_headers({1: "", 2: "", 3: ""},
                               {1: "", 2: "", 3: "南京农业大学学士学位论文"}, 2)
    assert len(findings) == 1
    finding = findings[0]
    assert "无" in finding.expected and "南京农业大学学士学位论文" in finding.actual


def test_internal_format_labels_are_chinese():
    assert _display_key("段落/alignment") == "对齐方式"
    assert _display_key("布局/paragraph_alignment") == "公式段落对齐"
    assert _display_profile_value("段落/alignment", 0) == "左对齐"
    assert _display_profile_value("段落/alignment", 3) == "两端对齐"
    assert _display_profile_value("段落/line_spacing_rule", 4) == "固定值"
    assert _display_profile_value("段落/keep_with_next", True) == "是"


def test_formula_label_uses_stable_object_locator():
    """Formula labels identify the object without quoting surrounding prose."""

    assert _formula_label({"value": "3-1"}, 12, "display") == (
        "\u72ec\u7acb\u516c\u5f0f\uff08\u7f16\u53f7 3-1\uff0c\u7b2c 12 \u6bb5\uff09"
    )
    assert _formula_label(None, 12, "display") == "\u72ec\u7acb\u516c\u5f0f\uff08\u7b2c 12 \u6bb5\uff09"
    assert _formula_label({"value": "3-1"}, 12, "inline") == (
        "\u884c\u5185\u516c\u5f0f\uff08\u7f16\u53f7 3-1\uff0c\u7b2c 12 \u6bb5\uff09"
    )
