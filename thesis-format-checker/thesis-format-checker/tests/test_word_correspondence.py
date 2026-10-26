"""模板 ←→ 论文“同一层级 + 同一位置”对应关系的单元测试（不需要 Word）。"""

from thesis_format_checker.core.models import BodySample, HeadingSample, HeadingSpec, IndentSpec, SpacingSpec
from thesis_format_checker.core.pairing import dominant_value, partner_index, style_key
from thesis_format_checker.core.scope import ParagraphInfo
from thesis_format_checker.detectors.word_effective.correspondence import (
    build_body_basis,
    build_heading_basis,
)


def body_sample(index, style_name="正文", size_pt=12.0, chinese_font="宋体"):
    return BodySample(index=index, style_name=style_name, chinese_font=chinese_font, size_pt=size_pt)


def body_sample_with_spacing(index, style_name, spacing):
    return BodySample(index=index, style_name=style_name, size_pt=12.0, spacing=spacing)


def body_paragraph(index, style_name="正文"):
    return ParagraphInfo(index=index, paragraph=None, text="示例正文", style_name=style_name)


def heading_spec(size_pt):
    return HeadingSpec("宋体", "Times New Roman", size_pt, True, 1, IndentSpec(0, 0), 0, 0)


def test_style_key_ignores_spaces_and_case():
    assert style_key(" 正 文 ") == style_key("正文")
    assert style_key("Body Text") == style_key("bodytext")
    assert style_key(None) == ""


def test_partner_index_clamps_and_maps_by_ratio():
    assert partner_index(0, 3, 3) == 0
    assert partner_index(2, 3, 3) == 2
    assert partner_index(0, 1, 9) == 0
    assert partner_index(8, 1, 9) == 0
    assert partner_index(0, 2, 9) == 0
    assert partner_index(8, 2, 9) == 1
    assert partner_index(4, 0, 9) == -1


def test_body_basis_uses_position_when_counts_match():
    template = [body_sample(1, size_pt=12.0), body_sample(2, size_pt=10.5), body_sample(3, size_pt=12.0)]
    document = [body_paragraph(10), body_paragraph(11), body_paragraph(12)]
    basis = build_body_basis(template, document)
    assert [basis[index].index for index in (10, 11, 12)] == [1, 2, 3]


def test_body_basis_uses_level_mode_when_counts_differ():
    """模板只放了两段样张时取同层级最常见样例，不做比例硬对齐。"""

    template = [body_sample(1, size_pt=12.0), body_sample(2, size_pt=10.5)]
    document = [body_paragraph(index) for index in range(10, 15)]
    basis = build_body_basis(template, document)
    assert {basis[index].size_pt for index in range(10, 15)} == {12.0}
    assert {basis[index].index for index in range(10, 15)} == {1}


def test_body_basis_groups_different_styles_apart():
    template = [body_sample(1, "正文"), body_sample(2, "参考文献", size_pt=10.5)]
    document = [body_paragraph(10, "正文"), body_paragraph(11, "参考文献")]
    basis = build_body_basis(template, document)
    assert basis[10].index == 1
    assert basis[11].index == 2


def test_body_basis_falls_back_to_whole_template_for_new_style():
    template = [body_sample(1), body_sample(2)]
    document = [body_paragraph(10, "引用"), body_paragraph(11, "引用")]
    basis = build_body_basis(template, document)
    assert {basis[index].style_name for index in (10, 11)} == {"正文"}


def test_body_basis_without_template_paragraphs_returns_nothing():
    assert build_body_basis([], [body_paragraph(10)]) == {}


def test_body_basis_ignores_tiny_style_group():
    """模板里只有零星几段同名字段落时，改用整篇正文的典型样例（避免行距基准取错）。"""

    single = SpacingSpec(0, 12.0, 0.0, 0.0)
    fixed_20 = SpacingSpec(4, 20.0, 0.0, 0.0)
    template = [body_sample_with_spacing(1, "正文", single)] + [
        body_sample_with_spacing(index, "论文正文", fixed_20) for index in (2, 3, 4)
    ]
    document = [body_paragraph(index) for index in range(10, 15)]
    basis = build_body_basis(template, document)
    assert {sample.spacing.line_points for sample in basis.values()} == {20.0}


def test_heading_basis_uses_position_when_counts_match():
    template = [
        HeadingSample(3, 1, heading_spec(16)),
        HeadingSample(9, 1, heading_spec(18)),
    ]
    document = [(100, 1, None), (200, 1, None)]
    basis = build_heading_basis(template, document)
    assert basis[100].index == 3 and basis[200].index == 9


def test_heading_basis_uses_level_mode_when_counts_differ():
    template = [HeadingSample(3, 1, heading_spec(16)), HeadingSample(9, 1, heading_spec(16))]
    document = [(index, 1, None) for index in range(100, 108)]
    basis = build_heading_basis(template, document)
    assert {basis[index].spec.size_pt for index, _, _ in document} == {16.0}


def test_heading_basis_leaves_missing_level_unmapped():
    template = [HeadingSample(3, 1, heading_spec(16))]
    assert build_heading_basis(template, [(100, 3, None)]) == {}


def test_dominant_value_prefers_weight_and_keeps_first_on_tie():
    assert dominant_value([("宋体", 5), ("黑体", 1)]) == "宋体"
    assert dominant_value([("宋体", 3), ("黑体", 3)]) == "宋体"
    assert dominant_value([(None, 9), ("黑体", 2)]) == "黑体"
    assert dominant_value([]) is None
