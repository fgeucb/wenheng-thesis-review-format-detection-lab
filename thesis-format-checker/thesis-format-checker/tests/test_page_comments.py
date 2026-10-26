"""页眉页脚批注：按页锚定（版心顶部/底部）、无损拆分 run、保存校验。

方案来自同门实现：Word 不允许把批注写进页眉页脚，因此把页眉问题锚在该页版心
顶部的正文文字上、页脚问题锚在版心底部，批注正文写明「第 N 页｜页眉/页脚」。
"""

from __future__ import annotations

from docx import Document

from thesis_format_checker.core.docx_context import DocxContext
from thesis_format_checker.core.models import Finding, RunSummary, Status
from thesis_format_checker.core.point_comments import (
    add_point_comment, content_tree, marker_offset,
)
from thesis_format_checker.core.review_settings import validate_visible_comments, visible_comments_settings
from thesis_format_checker.core.word_layout import (
    BodyPoint, Page, Pagination, _pagination_from_layout, paragraph_matches,
    select_body_points, select_safe_page_boundaries, utf16_offsets,
)
from thesis_format_checker.engine.pipeline import save_comments_only, write_page_comments


def test_paragraph_matches_only_unique_text():
    xml = ["摘要", "正文一", "正文一", "第一章"]
    word = ["摘要", "正文一", "正文一", "第一章", "文本框内容"]
    matches = paragraph_matches(xml, word)
    # 重复文字只有两侧唯一段落和整段序列完全一致时才能映射。
    assert matches == {0: 0, 1: 1, 2: 2, 3: 3}
    assert paragraph_matches(["摘要", "重复", "重复", "结尾"],
                             ["摘要", "重复", "结尾"]) == {0: 0, 2: 3}
    assert paragraph_matches(["甲", "乙"], ["甲", "乙"]) == {0: 0, 1: 1}


def test_utf16_offsets_count_surrogate_pairs():
    text = "中A😀B"
    offsets = utf16_offsets(text)
    assert offsets[-1] == len(text.encode("utf-16-le")) // 2
    assert offsets[3] - offsets[2] == 2      # emoji 占两个 UTF-16 码元


def test_select_body_points_short_page_never_reuses_top():
    page_height = 800.0
    top = BodyPoint(0, 0, 0, 100.0)
    bottom = BodyPoint(0, 5, 6, 120.0)
    chosen_top, chosen_bottom = select_body_points([(top, bottom)], page_height)
    assert chosen_top is top and chosen_bottom is bottom
    far_bottom = BodyPoint(1, 9, 40, 700.0)
    chosen_top, chosen_bottom = select_body_points([(top, bottom), (BodyPoint(1, 0, 20, 200.0), far_bottom)],
                                                   page_height)
    assert chosen_top is top and chosen_bottom is far_bottom


def test_page_boundaries_skip_only_complex_edges_on_the_same_page():
    records = [(0, 5, "复杂"), (5, 10, "顶部"), (10, 15, "底部"), (15, 20, "复杂")]
    points = {
        (1, False): BodyPoint(1, 0, 5),
        (2, False): BodyPoint(2, 0, 10),
        (1, True): BodyPoint(1, 2, 9),
        (2, True): BodyPoint(2, 2, 14),
    }

    top, bottom = select_safe_page_boundaries(
        [0, 1, 2, 3], lambda index, is_bottom: points.get((index, is_bottom)),
        records, 0, 20)

    assert top.paragraph == 1 and bottom.paragraph == 2


def test_add_point_comment_splits_run_without_changing_text(tmp_path):
    document = Document()
    paragraph = document.add_paragraph("前文问题文字后文")
    comment = add_point_comment(document, paragraph, 2, "【警告】测试批注")
    path = tmp_path / "point.docx"
    document.save(path)

    reloaded = Document(path)
    assert reloaded.paragraphs[0].text == "前文问题文字后文"      # 文字不变
    assert [item.text for item in reloaded.comments] == ["【警告】测试批注"]
    markers = reloaded.element.body.findall(".//w:commentRangeStart") if False else None
    del markers
    starts = reloaded.element.body.xpath(".//w:commentRangeStart")
    ends = reloaded.element.body.xpath(".//w:commentRangeEnd")
    assert len(starts) == len(ends) == 1
    assert marker_offset(starts[0].getparent(), starts[0]) == 2
    assert marker_offset(ends[0].getparent(), ends[0]) == 3
    assert comment.comment_id is not None


def test_content_tree_tolerates_run_splits_only():
    original = Document()
    paragraph = original.add_paragraph("前文问题文字后文")
    before = content_tree(paragraph._p)

    document = Document()
    paragraph = document.add_paragraph("前文问题文字后文")
    add_point_comment(document, paragraph, 2, "测试")
    assert content_tree(paragraph._p) == before

    other = Document()
    changed = other.add_paragraph("前文问题文字后文").runs[0]
    changed.text = "前文别字文字后文"
    assert content_tree(other.paragraphs[0]._p) != before


def test_write_page_comments_anchors_on_the_affected_page(tmp_path, monkeypatch):
    document = Document()
    document.add_paragraph("第一页正文")
    document.add_paragraph("第二页正文")
    context = DocxContext(document)
    summary = RunSummary()
    finding = Finding("036.1", "页眉字体不符合要求", Status.FAIL,
                      "【错误】页眉字体不符合要求\n应该：黑体\n实际：宋体",
                      section_index=2, page_story="header")
    pagination = Pagination(pages=[
        Page(1, 0, top=BodyPoint(0, 0, 0)),
        Page(2, 1, top=BodyPoint(1, 0, 100)),
    ])
    monkeypatch.setattr("thesis_format_checker.core.word_layout.read_pagination",
                        lambda source, ctx: pagination)

    context.render_facts = {"total_pages": 2}      # 模拟“跑过 Word 流程”
    written, unplaced = write_page_comments(summary, context, tmp_path / "source.docx", document, [finding])

    assert len(written) == 1
    assert unplaced == []
    assert written[0]["page"] == 2 and written[0]["story"] == "header"
    assert written[0]["paragraph"] == 2 and written[0]["offset"] == 0
    text = [comment.text for comment in document.comments][0]
    assert text.startswith("【警告】页眉字体不符合要求")
    assert "文件页序" not in text and "｜页眉" not in text
    starts = document.element.body.xpath(".//w:commentRangeStart")
    ends = document.element.body.xpath(".//w:commentRangeEnd")
    assert marker_offset(starts[0].getparent(), starts[0]) == marker_offset(ends[0].getparent(), ends[0])


def test_page_comment_is_unplaced_when_no_word_facts(tmp_path):
    """没跑过 Word 时不启动 Word：由调用方保留在报告中。"""

    document = Document()
    document.add_paragraph("正文")
    context = DocxContext(document)
    finding = Finding("037.2", "偶数页页眉不符合模板", Status.FAIL, "【警告】偶数页页眉",
                      section_index=1, page_story="header")
    written, unplaced = write_page_comments(RunSummary(), context, tmp_path / "s.docx",
                                            document, [finding])
    assert written == [] and unplaced == [finding]


def test_save_comments_only_allows_run_splits_and_keeps_other_parts(tmp_path):
    source = tmp_path / "source.docx"
    document = Document()
    document.add_paragraph("正文内容一")
    document.save(source)

    working = Document(source)
    add_point_comment(working, working.paragraphs[0], 0, "测试批注")
    output = tmp_path / "out.docx"
    save_comments_only(source, working, output, allow_run_splits=True)

    loaded = Document(output)
    assert loaded.paragraphs[0].text == "正文内容一"
    assert [comment.text for comment in loaded.comments] == ["测试批注"]


def test_visible_comments_settings_enables_markup():
    document = Document()
    blob = document.settings.element.xml.encode("utf-8")
    updated = visible_comments_settings(blob)
    assert validate_visible_comments(updated)


def test_pagination_from_layout_builds_top_and_bottom_points():
    """排版几何顺带产出的按页锚点：顶部取最靠上、底部取最靠下的正文段落。"""

    document = Document()
    document.add_paragraph("第一页顶部文字")
    document.add_paragraph("第一页底部文字")
    document.add_paragraph("第二页文字")
    context = DocxContext(document)
    layout = {
        "page_points": {
            1: [BodyPoint(0, 0, 0, 100.0), BodyPoint(1, 0, 0, 700.0)],
            2: [BodyPoint(2, 0, 0, 100.0)],
        },
        "page_info": {1: {"section": 0, "printed": 1}, 2: {"section": 0, "printed": 2}},
    }
    pagination = _pagination_from_layout(context, layout)
    assert [page.number for page in pagination.pages] == [1, 2]
    first = pagination.pages[0]
    assert first.top.paragraph == 0 and first.bottom.paragraph == 1
    assert first.adjusted_number == 1
    assert pagination.pages[1].top.paragraph == 2
