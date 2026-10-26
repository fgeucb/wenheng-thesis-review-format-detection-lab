from copy import deepcopy

from docx import Document
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm

from thesis_format_checker.core.docx_context import DocxContext, UNREADABLE
from thesis_format_checker.detectors.pages import (
    auto_header_section_map, auto_section_map, check_page_errors, extract_page_baseline,
)
from thesis_format_checker.core.models import Status


def run_check(template, document, mapping=None):
    before = (template.element.xml, document.element.xml,
              tuple(part.blob for part in document.part.package.parts))
    result = check_page_errors(extract_page_baseline(DocxContext(template)), DocxContext(document), mapping)
    after = (template.element.xml, document.element.xml,
             tuple(part.blob for part in document.part.package.parts))
    assert before == after, "read-only page detection changed the document package"
    return result


def family_of(rule_id):
    """细粒度规则编号 → 所属规则族（026.1 → 026）。"""

    return rule_id.split(".")[0]


def failed(result, rule):
    """按规则族查询失败项：族编号或其细粒度规则都算命中。"""

    return [f for f in result if family_of(f.rule_id) == rule and f.status == Status.FAIL]


def add_inline_image(paragraph, width):
    """在段落里放一个只有尺寸的嵌入图（不需要真实图片字节），用于触发 033 的宽度判断。"""

    drawing, inline, extent = (OxmlElement(tag) for tag in ("w:drawing", "wp:inline", "wp:extent"))
    extent.set("cx", str(width))
    extent.set("cy", str(width))
    inline.append(extent)
    drawing.append(inline)
    paragraph.add_run()._r.append(drawing)
    return paragraph


def test_page_dimensions_dynamic_and_existing_checks_not_duplicated():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("正文测试文字")
        doc.sections[0].gutter = Mm(7)
        doc.sections[0].header_distance = Mm(13)
        doc.sections[0].footer_distance = Mm(17)
    template.sections[0].orientation = WD_ORIENT.LANDSCAPE
    result = run_check(template, document)
    assert len(failed(result, "024")) == 1
    assert not any(family_of(f.rule_id) in {"023", "025"} for f in result)
    document.sections[0].orientation = WD_ORIENT.LANDSCAPE
    document.sections[0].gutter = Mm(3)
    document.sections[0].header_distance = Mm(22)
    document.sections[0].different_first_page_header_footer = True
    result = run_check(template, document)
    assert failed(result, "026") and failed(result, "027") and failed(result, "031")
    assert not failed(result, "024")
    assert f"{template.sections[0].gutter.mm:.3f} 毫米" in failed(result, "026")[0].message


def test_multisection_baseline_without_roles_is_skipped_not_asked_for_mapping():
    """模板多节且无法按角色自动对应时：本条不写批注，也不再要求人工配置 section_map。"""

    template, document = Document(), Document()
    template.add_paragraph("一般说明")
    template.add_section(WD_SECTION_START.NEW_PAGE).header_distance = Mm(31)
    template.add_paragraph("其他说明")
    document.add_paragraph("正文")
    document.sections[0].header_distance = Mm(24)
    result = run_check(template, document)
    assert not failed(result, "027")
    assert any(family_of(f.rule_id) == "027" and f.status == Status.NOT_APPLICABLE for f in result)
    # 显式配置 section_map 时仍按配置比对。
    result = run_check(template, document, {"section_map": {"1": 2}})
    assert failed(result, "027")
    assert f"{template.sections[1].header_distance.mm:.3f} 毫米" in failed(result, "027")[0].message


def test_auto_section_map_matches_cover_and_body_by_role():
    """论文把封面+声明、摘要+目录合并成更少的节时，按节角色自动对应模板节。"""

    template = Document()
    template.add_paragraph("论文题目：基于××的研究")     # 节 1：封面
    template.add_section(WD_SECTION_START.NEW_PAGE)
    template.add_paragraph("摘要")                       # 节 2：摘要
    template.add_section(WD_SECTION_START.NEW_PAGE)
    template.add_paragraph("第一章 绪论", "Heading 1")   # 节 3：正文

    document = Document()
    document.add_paragraph("论文题目：基于××的研究")     # 节 1：封面
    document.add_section(WD_SECTION_START.NEW_PAGE)
    document.add_paragraph("摘要")                       # 节 2：摘要+目录合并
    document.add_paragraph("目录")
    document.add_section(WD_SECTION_START.NEW_PAGE)
    document.add_paragraph("第一章 绪论", "Heading 1")   # 节 3：正文

    assert auto_section_map(DocxContext(template), DocxContext(document)) == {"1": 1, "2": 2, "3": 3}


def test_front_matter_static_story_and_dynamic_field_limitation():
    template, document = Document(), Document()
    template.add_paragraph("摘要")
    document.add_paragraph("摘要")
    document.sections[0].header.paragraphs[0].text = "正文页眉"
    result = run_check(template, document)
    # 逐变体比对：普通页/首页/偶数页各出一条（首页、偶数页变体继承普通页内容）。
    front = failed(result, "022")
    assert len(front) == 3
    assert all("前置部分" in item.expected and "页眉" in item.expected for item in front)
    assert not any(item.status == Status.UNKNOWN for item in result if family_of(item.rule_id) == "022")
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    document.sections[0].footer.paragraphs[0]._p.append(field)
    result = run_check(template, document)
    # 含 PAGE 域的页脚不再打包成“无法判定”：逐变体报出“模板无页脚、论文出现动态页码页脚”。
    footer = [f for f in failed(result, "022") if "页脚" in f.expected]
    assert len(footer) == 1     # 页脚按区域级比对：模板没有页码域、论文出现动态页码域
    assert "动态" in footer[0].actual
    assert not any(f.status == Status.UNKNOWN for f in result if family_of(f.rule_id) == "022")


def test_disabled_first_header_does_not_count_as_visible():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("摘要")
    document.sections[0].first_page_header.paragraphs[0].text = "不启用的首页页眉"
    document.sections[0].different_first_page_header_footer = False
    assert not failed(run_check(template, document), "022")


def test_inherited_heading_page_break_and_natural_pagination_unknown():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("上一章内容")
        doc.add_heading("第二章 方法", level=1)
    template.styles["Heading 1"].paragraph_format.page_break_before = True
    result = run_check(template, document)
    assert not failed(result, "030")
    assert any(family_of(f.rule_id) == "030" and "自然分页" in f.message for f in result)
    document.styles["Heading 1"].paragraph_format.page_break_before = True
    assert not any(family_of(f.rule_id) == "030" for f in run_check(template, document))


def test_section_boundary_change_is_reported_only_for_matched_transition():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("摘要")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_heading("第一章 绪论", level=1)
    document.sections[1].gutter = Mm(8)
    result = run_check(template, document)
    assert failed(result, "029")
    assert failed(result, "026")


def test_transition_ignores_disabled_first_header_and_tracks_even_pages():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("论文题目：测试")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_paragraph("原创性声明")
    section = document.sections[1]
    section.first_page_header.is_linked_to_previous = False
    section.first_page_header.paragraphs[0].text = "不启用的首页页眉"
    section.even_page_header.is_linked_to_previous = False
    section.even_page_header.paragraphs[0].text = "异常偶数页文字"
    document.settings.odd_and_even_pages_header_footer = True
    result = [f for f in run_check(template, document) if f.rule_id == "029" and f.status == Status.FAIL]
    assert len(result) == 1
    f = result[0]
    assert f.page_story == "header" and f.page_variant == "even"
    assert f.section_index == 2 and f.anchor is None
    assert "首页" not in f.actual


def test_fixed_table_width_and_autofit_limit():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("正文")
        doc.sections[0].page_width = Mm(210)
        doc.sections[0].left_margin = Mm(30)
        doc.sections[0].right_margin = Mm(30)
    table = document.add_table(rows=1, cols=2)
    table.autofit = False
    for column in table.columns:
        column.width = Mm(90)
    for cell in table.rows[0].cells:
        cell.width = Mm(90)
    table.cell(0, 0).text = "过宽表格"
    result = run_check(template, document)
    assert failed(result, "033")
    usable = template.sections[0].page_width - template.sections[0].left_margin - template.sections[0].right_margin
    assert f"{usable / 36000:.3f} 毫米" in failed(result, "033")[0].message
    table.autofit = True
    result = run_check(template, document)
    assert not failed(result, "033")
    assert any("自动调整表格" in f.message for f in result)


def test_unreadable_section_property_reports_exact_limitation():
    template, document = Document(), Document()
    template.add_paragraph("正文")
    document.add_paragraph("正文")
    margins = document.sections[0]._sectPr.find(qn("w:pgMar"))
    margins.set(qn("w:header"), "invalid-width")
    result = run_check(template, document)
    assert any(family_of(f.rule_id) == "027" and f.status == Status.UNKNOWN and UNREADABLE in f.message
               for f in result)


def test_matching_template_does_not_raise_format_failures():
    template = Document()
    template.add_paragraph("摘要")
    template.add_paragraph("摘要内容")
    template.add_section(WD_SECTION_START.NEW_PAGE)
    template.add_heading("第一章 绪论", level=1)
    template.add_paragraph("内容")
    document = Document()
    document.element.body.clear()
    for child in template.element.body:
        document.element.body.append(deepcopy(child))
    assert not any(f.status == Status.FAIL for f in run_check(template, document))


def test_template_overwide_table_is_an_unknown_exception_not_a_new_rule():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("正文")
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        for column in table.columns:
            column.width = Mm(150)
        table.cell(0, 0).text = "模板中的特殊宽表格"
    result = run_check(template, document)
    assert not failed(result, "033")
    assert any("模板含超出内容区域的表格示例" in f.message for f in result)


def test_direct_outline_heading_and_page_anchor_skip_leading_blank():
    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph()
        paragraph = doc.add_paragraph("第一章 绪论")
        outline = OxmlElement("w:outlineLvl")
        outline.set(qn("w:val"), "0")
        paragraph._p.get_or_add_pPr().append(outline)
    template.paragraphs[1].paragraph_format.page_break_before = True
    document.sections[0].header_distance = Mm(30)
    result = run_check(template, document)
    assert any(family_of(f.rule_id) == "030" and f.status == Status.UNKNOWN for f in result)
    assert failed(result, "027")[0].anchor.text == "第一章 绪论"


def test_roles_match_moved_sections_without_assuming_same_section_number():
    template, document = Document(), Document()
    template.add_paragraph("摘要")
    template.sections[0].header_distance = Mm(11)
    template.add_section(WD_SECTION_START.NEW_PAGE)
    template.add_heading("第一章 绪论", level=1)
    template.sections[1].header_distance = Mm(19)
    document.add_heading("第一章 绪论", level=1)
    document.sections[0].header_distance = Mm(19)
    document.add_section(WD_SECTION_START.NEW_PAGE)
    document.add_paragraph("摘要")
    document.sections[1].header_distance = Mm(11)
    result = run_check(template, document)
    assert not failed(result, "027")


def test_header_section_map_uses_consensus_across_same_role_sections():
    template, document = Document(), Document()
    template.add_heading("第一章 绪论", level=1)
    template.sections[0].header.paragraphs[0].text = "学校论文页眉"
    template.add_section(WD_SECTION_START.NEW_PAGE)
    template.add_heading("第二章 方法", level=1)
    document.add_heading("第一章 绪论", level=1)
    mapping = auto_header_section_map(
        DocxContext(template), DocxContext(document))
    assert mapping == {"1": 1}


def test_character_unit_indent_is_converted_before_width_judgement():
    """字符单位缩进按字号换算成磅值参与宽度判定（不再默认 0，也不再一律无法判定）。"""

    template, document = Document(), Document()
    template.add_paragraph("正文")
    paragraph = document.add_paragraph("图说明")
    ind = OxmlElement("w:ind")
    ind.set(qn("w:leftChars"), "200")
    paragraph._p.get_or_add_pPr().append(ind)
    drawing, inline, extent = (OxmlElement(tag) for tag in ("w:drawing", "wp:inline", "wp:extent"))
    extent.set("cx", str(Mm(120)))
    extent.set("cy", str(Mm(20)))
    inline.append(extent)
    drawing.append(inline)
    paragraph.add_run()._r.append(drawing)
    result = run_check(template, document)
    # 120 毫米图片 + 2 字符左缩进仍在可打印宽度内 → 不应报裁切风险。
    assert not failed(result, "033")


def test_conflicting_table_cell_and_grid_width_are_unknown():
    template, document = Document(), Document()
    template.add_paragraph("正文")
    document.add_paragraph("正文")
    table = document.add_table(rows=1, cols=2)
    table.autofit = False
    # Deliberately leave tcW unchanged while changing only tblGrid.
    for column in table.columns:
        column.width = Mm(200)
    result = run_check(template, document)
    assert not failed(result, "033")
    assert any("单元格宽度与网格列宽冲突" in f.message for f in result)


def test_nested_fixed_table_reports_layout_limit():
    template, document = Document(), Document()
    template.add_paragraph("正文")
    document.add_paragraph("正文")
    outer = document.add_table(rows=1, cols=1)
    nested = outer.cell(0, 0).add_table(rows=1, cols=1)
    nested.autofit = False
    nested.cell(0, 0).text = "嵌套内容"
    result = run_check(template, document)
    assert any("嵌套表格需要外围单元格" in f.message for f in result)


def test_cover_table_image_is_left_to_cover_media_rule():
    """封面表格里的校徽、校名图片由 017 规则比对，033 不再输出“无法判定”。"""

    template, document = Document(), Document()
    for doc in (template, document):
        doc.add_paragraph("正文")
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "论文题目"
        table.cell(1, 0).text = "学号"
        add_inline_image(table.cell(1, 1).paragraphs[0], Mm(20))
    result = run_check(template, document)
    assert not any("表格单元格中的图片" in f.message for f in result)
    assert not failed(result, "033")


def test_body_table_image_keeps_page_width_limitation():
    """正文表格图片没有封面基线可比对，仍保留原有的保守结论。"""

    template, document = Document(), Document()
    template.add_paragraph("正文")
    document.add_paragraph("正文")
    table = document.add_table(rows=1, cols=1)
    add_inline_image(table.cell(0, 0).paragraphs[0], Mm(20))
    result = run_check(template, document)
    assert any("表格单元格中的图片" in f.message for f in result)
