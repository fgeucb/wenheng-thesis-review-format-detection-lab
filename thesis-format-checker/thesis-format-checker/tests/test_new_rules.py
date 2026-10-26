from types import SimpleNamespace

from thesis_format_checker.core.models import HeadingSpec, IndentSpec, SpacingSpec
from thesis_format_checker.detectors.word_effective.paper_size import compare_paper_size
from thesis_format_checker.detectors.word_effective.headers import compare_headers, read_headers
from thesis_format_checker.detectors.word_effective.headings import compare_heading, heading_level
from thesis_format_checker.detectors.word_effective.paragraph_spacing import compare_spacing
from thesis_format_checker.detectors.word_effective.punctuation import inspect_punctuation
from thesis_format_checker.detectors.word_effective.first_line_indent import compare_first_line_indent


def test_a4_both_orientations_and_a3():
    def section(w, h):
        return SimpleNamespace(PageSetup=SimpleNamespace(PageWidth=w * 72 / 25.4, PageHeight=h * 72 / 25.4))
    assert compare_paper_size(section(210, 297), 1) is None
    assert compare_paper_size(section(297, 210), 2) is None
    assert compare_paper_size(section(297, 420), 3).rule_id == "023"


def test_header_content_mismatch():
    result = compare_headers({1: "南京理工大学", 2: "", 3: ""}, {1: "其他学校", 2: "", 3: ""}, 1)
    assert len(result) == 1 and result[0].header_kind == 1


def test_disabled_even_header_uses_primary_text_as_template_basis():
    def section(primary, even, enabled):
        return SimpleNamespace(
            PageSetup=SimpleNamespace(
                DifferentFirstPageHeaderFooter=False,
                OddAndEvenPagesHeaderFooter=enabled),
            Headers=lambda kind: SimpleNamespace(Range=SimpleNamespace(
                Text=primary if kind == 1 else even)),
        )
    expected = read_headers(section("学校论文", "未启用文字", False))
    actual = read_headers(section("其他标题", "学校论文", True))
    findings = compare_headers(expected, actual, 1)
    assert [finding.header_kind for finding in findings] == [1]


def test_heading_levels_and_properties():
    para = SimpleNamespace(Range=SimpleNamespace(Style=SimpleNamespace(NameLocal="Heading 2")), OutlineLevel=2)
    assert heading_level(para) == 2
    standard = HeadingSpec("黑体", "Arial", 16, True, 1, IndentSpec(0, 0), 12, 6)
    wrong = HeadingSpec("宋体", "Arial", 14, False, 0, IndentSpec(12, 1), 0, 0)
    findings, unknowns = compare_heading(standard, wrong, 2, 5)
    # 一个属性一条结论：中文字体（046.1）等分别成条
    by_code = {item.rule_id: item for item in findings}
    assert {code.split(".")[0] for code in by_code} == {"046"}
    assert by_code["046.1"].expected == "黑体" and by_code["046.1"].actual == "宋体"
    assert by_code["046.3"].expected == "16 磅" and by_code["046.3"].actual == "14 磅"
    assert by_code["046.4"].expected == "加粗" and by_code["046.4"].actual == "未加粗"
    assert by_code["046.5"].expected == "居中" and by_code["046.5"].actual == "左对齐"
    assert by_code["046.6"].expected == "0 字符" and by_code["046.6"].actual == "1 字符"
    assert by_code["046.7"].expected == "12 磅" and by_code["046.7"].actual == "0 磅"
    messages = "；".join(item.message for item in findings)
    assert all(token not in messages for token in ("True", "False", "模板 1，当前 0"))
    assert not unknowns
    third, _ = compare_heading(standard, wrong, 3, 6)
    assert {item.rule_id.split(".")[0] for item in third} == {"047"}
    first, _ = compare_heading(standard, wrong, 1, 4)
    first_by_code = {item.rule_id: item for item in first}
    assert first_by_code["045.4"].expected == "加粗" and first_by_code["045.4"].actual == "未加粗"
    assert first_by_code["045.5"].expected == "居中" and first_by_code["045.5"].actual == "左对齐"


def test_heading_fonts_only_check_scripts_present_in_text():
    standard = HeadingSpec("黑体", "Arial", 16, True, 1, IndentSpec(0, 0), 0, 0)
    wrong = HeadingSpec("宋体", "Times New Roman", 16, True, 1, IndentSpec(0, 0), 0, 0)
    chinese, _ = compare_heading(standard, wrong, 1, 4, text="第一章 绪论")
    assert {finding.rule_id for finding in chinese} == {"045.1"}
    assert (chinese[0].start, chinese[0].end) == (0, 3)
    english, _ = compare_heading(standard, wrong, 1, 5, text="ABSTRACT")
    assert {finding.rule_id for finding in english} == {"045.2"}
    assert (english[0].start, english[0].end) == (0, 8)


def test_word_heading_classification_excludes_english_caption():
    facts = SimpleNamespace(
        style_name=lambda index: "Heading 2",
        outline_level=lambda index: 2,
        clean_text=lambda index: "Fig.2-2  The overall model flow diagram.",
    )
    assert heading_level(object(), facts=facts, index=1) is None


def test_unknown_heading_alignment_does_not_show_word_code():
    standard = HeadingSpec("黑体", "Arial", 16, True, 9999999, IndentSpec(0, 0), 0, 0)
    actual = HeadingSpec("黑体", "Arial", 16, True, 0, IndentSpec(0, 0), 0, 0)
    findings, unknowns = compare_heading(standard, actual, 1, 3)
    assert findings == []
    assert len(unknowns) == 1 and "对齐方式无法识别" in unknowns[0]
    assert "9999999" not in unknowns[0]


def test_line_and_paragraph_spacing():
    expected = SpacingSpec(1, 18, 0, 6)
    actual = SpacingSpec(2, 20, 3, 0)
    # 行距（074）+ 段前（075.1）/段后（075.2）分别成条
    assert [f.rule_id for f in compare_spacing(expected, actual, 7)] == ["074", "075.1", "075.2"]
    assert "模板 A：1.5 倍行距；当前文档：双倍行距" in compare_spacing(expected, actual, 7)[0].message
    assert "规则 1" not in compare_spacing(expected, actual, 7)[0].message


def test_fixed_line_spacing_and_same_rule_different_height():
    fixed = compare_spacing(SpacingSpec(4, 20, 0, 0), SpacingSpec(4, 24, 0, 0), 1)[0]
    assert "固定值行距 20 磅" in fixed.message and "固定值行距 24 磅" in fixed.message
    same_rule = compare_spacing(SpacingSpec(1, 18, 0, 0), SpacingSpec(1, 21, 0, 0), 1)[0]
    assert "计算行高约 18 磅" in same_rule.message and "计算行高约 21 磅" in same_rule.message


def test_punctuation_detects_mixed_unpaired_repeated_and_suspected_missing():
    findings = inspect_punctuation("中文,文本（内容！！", 4)
    # 细粒度拆分：中英文标点混用 122.1，括号不成对 123.2，重复标点 123.3
    assert {f.rule_id for f in findings} == {"122.1", "123.2", "123.3"}
    assert any("疑似" in f.message for f in inspect_punctuation("这一段完整的中文正文没有句末标点", 5))
    assert inspect_punctuation("中文标点正确，英文 word, next。", 6) == []
    assert any("段首" in f.message for f in inspect_punctuation("，段首多余标点。", 7))
    assert any("段末" in f.message for f in inspect_punctuation("正文段落尚未结束，", 8))


def test_two_character_indent_even_if_template_spec_differs():
    para = SimpleNamespace(Format=SimpleNamespace(FirstLineIndent=0, CharacterUnitFirstLineIndent=0))
    result = compare_first_line_indent(para, "正文内容", IndentSpec(0, 0), 1)
    assert any("2 字符" in f.message for f in result)
