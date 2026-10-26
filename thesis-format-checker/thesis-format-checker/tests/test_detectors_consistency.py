"""Internal consistency does not use the template's example thesis values."""
from copy import deepcopy

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from thesis_format_checker.core.docx_context import DocxContext, UNREADABLE
from thesis_format_checker.detectors.consistency import extract_consistency_baseline, check_consistency_errors
from thesis_format_checker.core.models import Status


def family_of(rule_id):
    """细粒度规则编号 → 所属规则族（036.1 → 036）。"""

    return rule_id.split(".")[0]



def detect(doc, template=None):
    template = template if template is not None else Document()
    before = doc.element.xml
    header_before = [r.target_part.blob for r in doc.part.rels.values() if r.reltype.endswith('/header')]
    template_before = template.element.xml
    result = check_consistency_errors(extract_consistency_baseline(DocxContext(template)), DocxContext(doc))
    assert doc.element.xml == before
    assert template.element.xml == template_before
    assert [r.target_part.blob for r in doc.part.rels.values() if r.reltype.endswith('/header')] == header_before
    assert all(f.expected.startswith('内部一致性要求（非模板固定值）') for f in result)
    # 统一批注格式：抬头 + 应该 + 实际（定位由调度层补充）
    assert all(len(f.message.splitlines()) == 3 for f in result)
    assert all(family_of(f.rule_id) in {'134', '136', '137'} for f in result)
    return result


def failed(results, rule):
    return [f for f in results if family_of(f.rule_id) == rule and f.status == Status.FAIL]


def unknown(results, rule):
    found = [f for f in results if family_of(f.rule_id) == rule and f.status == Status.UNKNOWN]
    assert all(UNREADABLE in f.message for f in found)
    return found


def bookmark(p, name, ident=1, start_run=0, end_run=None):
    begin = OxmlElement('w:bookmarkStart')
    begin.set(qn('w:name'), name)
    begin.set(qn('w:id'), str(ident))
    finish = OxmlElement('w:bookmarkEnd')
    finish.set(qn('w:id'), str(ident))
    runs = p.runs
    if runs:
        runs[start_run]._r.addprevious(begin)
        runs[-1 if end_run is None else end_run]._r.addnext(finish)
    else:
        p._p.append(begin)
        p._p.append(finish)


def field(p, instruction, value):
    node = OxmlElement('w:fldSimple')
    node.set(qn('w:instr'), instruction)
    run = OxmlElement('w:r')
    text = OxmlElement('w:t')
    text.text = value
    run.append(text)
    node.append(run)
    p._p.append(node)
    return node


def caption(doc, number='1', kind='figure'):
    if kind == 'figure':
        # Same structural inline picture evidence used by the caption reader;
        # this unit fixture does not require a Word drawing renderer.
        object_p = doc.add_paragraph()
        drawing = OxmlElement('w:drawing')
        inline = OxmlElement('wp:inline')
        drawing.append(inline)
        object_p.add_run()._r.append(drawing)
        return doc.add_paragraph('图' + number + ' 实验结果', style='Caption')
    p = doc.add_paragraph('表' + number + ' 结果', style='Caption')
    doc.add_table(rows=1, cols=1)
    return p


def equation(doc, number='1'):
    p = doc.add_paragraph()
    mathp = OxmlElement('m:oMathPara')
    math = OxmlElement('m:oMath')
    run = OxmlElement('m:r')
    text = OxmlElement('m:t')
    text.text = 'x=y'
    run.append(text)
    math.append(run)
    mathp.append(math)
    p._p.append(mathp)
    p.add_run('(' + number + ')')
    return p


def test_134_document_title_values_are_internal_not_template_values():
    template = Document()
    template.add_paragraph('论文题目：模板示例论文')
    doc = Document()
    doc.add_paragraph('论文题目：实际论文')
    doc.add_heading('摘要', 1)
    doc.add_paragraph('论文题目：实际论文')
    doc.add_paragraph('英文题目：A Different Language')
    assert not failed(detect(doc, template), '134')


def test_134_explicit_mismatch_and_subtitle_anchor():
    doc = Document()
    doc.add_paragraph('论文题目：研究：副标题')
    doc.add_heading('摘要', 1)
    target = doc.add_paragraph('论文题目：研究：另一副标题')
    results = failed(detect(doc), '134')
    assert len(results) == 1
    assert results[0].anchor._p is target._p
    assert '研究：副标题' in results[0].expected


def test_134_label_in_table_and_generic_later_question_is_not_thesis_title():
    doc = Document()
    row = doc.add_table(rows=1, cols=2).rows[0]
    row.cells[0].text = '论文题目：'
    row.cells[1].text = '论文甲'
    doc.add_heading('摘要', 1)
    doc.add_paragraph('题目：不是论文标题')
    doc.add_paragraph('论文题目：论文甲')
    assert not failed(detect(doc), '134')


def test_134_static_labeled_header_compares_document_title_only():
    doc = Document()
    doc.add_paragraph('论文题目：论文甲')
    doc.sections[0].header.paragraphs[0].text = '论文题目：论文乙'
    f = failed(detect(doc), '134')
    assert f and 'header' in f[0].actual
    assert f[0].anchor is None  # 页眉题目不能回退到封面题目
    assert f[0].section_index == 1
    assert f[0].page_story == "header" and f[0].page_variant == "default"


def test_134_template_establishes_unlabeled_header_role_but_not_fullness():
    template = Document()
    template.add_paragraph('论文题目：模板完整题目')
    template.sections[0].header.paragraphs[0].text = '模板完整题目'
    doc = Document()
    doc.add_paragraph('论文题目：文档完整题目')
    doc.sections[0].header.paragraphs[0].text = '文档简称'
    results = detect(doc, template)
    assert unknown(results, '134')
    assert not failed(results, '134')


def test_134_dynamic_header_and_revised_title_are_unknown():
    doc = Document()
    doc.add_paragraph('论文题目：论文甲')
    p = doc.sections[0].header.paragraphs[0]
    p.text = '论文题目：'
    field(p, 'STYLEREF Title', '论文乙')
    assert unknown(detect(doc), '134')
    assert not failed(detect(doc), '134')
    p = doc.add_paragraph('论文题目：论文乙')
    insert = OxmlElement('w:ins')
    insert.append(p.runs[0]._r)
    p._p.append(insert)
    assert not failed(detect(doc), '134')


def test_134_disabled_first_even_and_orphan_header_parts_are_not_compared():
    doc = Document()
    doc.add_paragraph('论文题目：实际论文')
    section = doc.sections[0]
    section.first_page_header.paragraphs[0].text = '论文题目：旧首页题目'
    section.even_page_header.paragraphs[0].text = '论文题目：旧偶数页题目'
    section.different_first_page_header_footer = False
    doc.settings.odd_and_even_pages_header_footer = False
    assert not failed(detect(doc), '134')
    section.header.paragraphs[0].text = '论文题目：孤立部件题目'
    ref = next(r for r in section._sectPr.findall(qn('w:headerReference')) if r.get(qn('w:type')) == 'default')
    section._sectPr.remove(ref)
    assert not failed(detect(doc), '134')


def test_134_enabled_first_even_headers_and_inherited_first_are_compared():
    from docx.enum.section import WD_SECTION
    for variant in ('first', 'even'):
        doc = Document()
        doc.add_paragraph('论文题目：实际论文')
        if variant == 'first':
            doc.sections[0].first_page_header.paragraphs[0].text = '论文题目：错误题目'
            doc.sections[0].different_first_page_header_footer = True
        else:
            doc.sections[0].even_page_header.paragraphs[0].text = '论文题目：错误题目'
            doc.settings.odd_and_even_pages_header_footer = True
        assert failed(detect(doc), '134')
    doc = Document()
    doc.add_paragraph('论文题目：实际论文')
    doc.sections[0].first_page_header.paragraphs[0].text = '论文题目：继承错误题目'
    doc.sections[0].different_first_page_header_footer = False
    next_section = doc.add_section(WD_SECTION.NEW_PAGE)
    next_section.different_first_page_header_footer = True
    assert next_section.first_page_header.is_linked_to_previous
    result = failed(detect(doc), '134')
    assert len(result) == 1 and '第2节' in result[0].actual


def test_134_disabled_template_header_does_not_create_title_role():
    template = Document()
    template.add_paragraph('论文题目：模板题目')
    template.sections[0].first_page_header.paragraphs[0].text = '模板题目'
    template.sections[0].different_first_page_header_footer = False
    doc = Document()
    doc.add_paragraph('论文题目：实际题目')
    doc.sections[0].header.paragraphs[0].text = '章节名称'
    assert not [f for f in detect(doc, template) if f.rule_id == '134']


def test_134_hidden_header_title_is_unknown_not_visible_mismatch():
    doc = Document()
    doc.add_paragraph('论文题目：实际题目')
    p = doc.sections[0].header.paragraphs[0]
    p.text = '论文题目：隐藏旧题目'
    p.runs[0].font.hidden = True
    results = detect(doc)
    assert not failed(results, '134')
    assert any('隐藏文字' in f.actual for f in unknown(results, '134'))


def test_134_explicit_bookmark_title_is_compared_and_case_spaces_preserved():
    doc = Document()
    doc.add_paragraph('英文题目：A Title')
    p = doc.add_paragraph('A  Title')
    bookmark(p, 'thesis_title_en')
    assert failed(detect(doc), '134')


def test_136_reuses_reliable_toc_link_difference_without_template_toc():
    from test_detectors_toc import fixture
    doc, _, entries, headings = fixture()
    headings[0].runs[0].text = '更改后的绪论'
    results = detect(doc)
    assert failed(results, '136')
    assert failed(results, '136')[0].anchor._p is entries[0]._p
    assert unknown(results, '136')


def test_136_same_toc_and_saved_page_never_create_false_page_fail():
    from test_detectors_toc import fixture
    doc, _, entries, _ = fixture()
    entries[0].runs[-2].text = '999'
    results = detect(doc)
    assert not failed(results, '136')
    assert any('重新分页' in f.actual for f in unknown(results, '136'))


def test_136_missing_bookmark_and_manual_directory_are_unknown():
    from test_detectors_toc import fixture
    doc, _, _, headings = fixture()
    node = headings[0]._p.find(qn('w:bookmarkStart'))
    headings[0]._p.remove(node)
    results = detect(doc)
    assert unknown(results, '136')
    assert not failed(results, '136')
    doc = Document()
    doc.add_paragraph('目录')
    doc.add_paragraph('绪论\t1')
    assert unknown(detect(doc), '136')


def test_137_static_objects_match_and_four_types_missing_fail():
    doc = Document()
    doc.add_heading('第1章 绪论', 1)
    caption(doc)
    caption(doc, kind='table')
    equation(doc)
    doc.add_paragraph('见本文图1，本文表1，本文公式(1)，本文第1章。')
    assert not failed(detect(doc), '137')
    target = doc.add_paragraph('见本文图9，本文表9，本文公式(9)，本文第9章。')
    results = failed(detect(doc), '137')
    assert len(results) == 4
    assert all(f.anchor._p is target._p for f in results)


def test_137_ordinary_external_and_bibliographic_references_not_checked():
    doc = Document()
    doc.add_paragraph('见图9，表9，公式(9)，第9章。')
    doc.add_paragraph('引自文献[9]：本文图9。')
    doc.add_heading('参考文献', 1)
    doc.add_paragraph('本文图9 出版社。')
    assert not failed(detect(doc), '137')


def test_137_duplicate_number_range_and_local_scope_unknown():
    doc = Document()
    caption(doc)
    caption(doc)
    doc.add_paragraph('见本文图1。')
    doc.add_paragraph('见本文图1至图9。')
    doc.add_paragraph('见本章图1。')
    results = detect(doc)
    assert not failed(results, '137')
    assert len(unknown(results, '137')) == 3


def test_137_unattached_picture_prevents_absence_claim():
    doc = Document()
    p = doc.add_paragraph()
    drawing = OxmlElement('w:drawing')
    drawing.append(OxmlElement('wp:inline'))
    p.add_run()._r.append(drawing)
    doc.add_paragraph('见本文图9。')
    results = detect(doc)
    assert unknown(results, '137') and not failed(results, '137')


def test_137_seq_cache_not_static_object_number():
    doc = Document()
    p = caption(doc)
    p.text = '图'
    field(p, 'SEQ Figure', '9')
    p.add_run(' 结果')
    doc.add_paragraph('见本文图8。')
    results = detect(doc)
    assert unknown(results, '137') and not failed(results, '137')


def test_137_plain_ref_static_caption_bookmark_detects_cache_difference():
    doc = Document()
    p = caption(doc)
    bookmark(p, 'fig_result')
    target = doc.add_paragraph('见')
    field(target, 'REF fig_result \\h', '图9 实验结果')
    results = failed(detect(doc), '137')
    assert len(results) == 1
    assert results[0].anchor._p is target._p
    assert '图1 实验结果' in results[0].expected


def test_137_numbering_ref_switch_and_missing_bookmark():
    doc = Document()
    p = caption(doc)
    bookmark(p, 'fig_result')
    target = doc.add_paragraph('见本文图')
    field(target, 'REF fig_result \\n', '9')
    results = detect(doc)
    assert unknown(results, '137') and not failed(results, '137')
    target = doc.add_paragraph('见本文图')
    field(target, 'REF absent_figure \\h', '9')
    assert len(failed(detect(doc), '137')) == 1


def test_137_unrelated_missing_date_author_ref_not_classified_by_paragraph_label():
    doc = Document()
    caption(doc)
    p = doc.add_paragraph('见本文图1；报告日期：')
    field(p, 'REF missing_date', '2026年')
    p = doc.add_paragraph('见本文图1；作者：')
    field(p, 'REF missing_author', '张三')
    p = doc.add_paragraph('见本文图1；报告序号：')
    field(p, 'REF missing_report', '9')
    assert not failed(detect(doc), '137')


def test_137_field_cache_own_label_confirms_missing_object_ref():
    doc = Document()
    p = doc.add_paragraph('参见')
    field(p, 'REF missing_figure', '图9 实验结果')
    assert len(failed(detect(doc), '137')) == 1


def test_137_figure_label_pointing_to_table_number_reports_type_mismatch():
    doc = Document()
    p = caption(doc, kind='table')
    p.clear()
    p.add_run('表')
    p.add_run('1')
    p.add_run(' 数据')
    bookmark(p, 'table_num', start_run=1, end_run=1)
    ref = doc.add_paragraph('本文图')
    field(ref, 'REF table_num', '1')
    results = failed(detect(doc), '137')
    assert len(results) == 1 and '标签类型' in results[0].actual


def test_137_dirty_field_type_mismatch_cannot_assert_current_type():
    doc = Document()
    p = caption(doc, kind='table')
    bookmark(p, 'table_label')
    ref = doc.add_paragraph('见')
    f = field(ref, 'REF table_label', '图1')
    f.set(qn('w:dirty'), 'true')
    results = detect(doc)
    assert unknown(results, '137') and not failed(results, '137')


def test_137_revision_reference_not_definite_failure():
    doc = Document()
    p = doc.add_paragraph('本文图9。')
    change = OxmlElement('w:pPrChange')
    p._p.get_or_add_pPr().append(change)
    results = detect(doc)
    assert unknown(results, '137') and not failed(results, '137')


def test_137_deleted_object_version_blocks_absence_claim():
    doc = Document()
    p = caption(doc)
    deletion = OxmlElement('w:del')
    p._p.addprevious(deletion)
    deletion.append(p._p)
    doc.add_paragraph('本文图9。')
    results = detect(doc)
    assert unknown(results, '137') and not failed(results, '137')


def test_137_multilevel_section_and_ambiguous_short_section():
    doc = Document()
    doc.add_heading('第1章 绪论', 1)
    doc.add_heading('1.1 方法', 2)
    doc.add_paragraph('见本文第1.1节。')
    assert not failed(detect(doc), '137')
    doc.add_paragraph('见本文第1节。')
    assert unknown(detect(doc), '137')
    target = doc.add_paragraph('见本文第9.9节。')
    result = failed(detect(doc), '137')
    assert len(result) == 1 and result[0].anchor._p is target._p


def test_137_parenthesized_subfigure_needs_explicit_mapping():
    doc = Document()
    caption(doc)
    doc.add_paragraph('见本文图1(a)。')
    results = detect(doc)
    assert unknown(results, '137') and not failed(results, '137')


def test_134_multiline_title_retains_four_line_comment_contract():
    doc = Document()
    doc.add_paragraph('论文题目：标题甲\n副标题')
    doc.add_heading('摘要', 1)
    doc.add_paragraph('论文题目：标题乙\n副标题')
    assert failed(detect(doc), '134')


def test_failure_in_one_internal_reader_does_not_hide_other_check(monkeypatch):
    import thesis_format_checker.detectors.consistency as module
    doc = Document()
    doc.add_paragraph('论文题目：甲')
    doc.add_heading('摘要', 1)
    doc.add_paragraph('论文题目：乙')
    monkeypatch.setattr(module, '_snapshot', lambda ctx: 1 / 0)
    results = detect(doc)
    assert failed(results, '134') and unknown(results, '136')
