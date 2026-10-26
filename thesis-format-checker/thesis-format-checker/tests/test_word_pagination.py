"""需要 Word 排版结果的规则：054 标题同页、030 章节另起页、040 页码。"""

from __future__ import annotations

from types import SimpleNamespace

from thesis_format_checker.core.word_pages import read_page_facts
from thesis_format_checker.detectors.word_effective.pagination import (
    check_chapter_page_breaks, check_heading_pagination, check_toc_layout, verify_page_numbering,
)


class FakeRange:
    def __init__(self, text, page, section):
        self.Text = text
        self._page = page
        self._section = section
        self._fields = 0
        self.Start = 0
        self.End = max(len(text), 1)
        self.Fields = SimpleNamespace(Count=self._fields)

    def Information(self, kind):
        return self._page if kind == 3 else self._section


class FakeParagraph:
    def __init__(self, text, page, section=1):
        self.Range = FakeRange(text, page, section)
        self.Style = SimpleNamespace(NameLocal="正文")


class FakeStory:
    def __init__(self, text, fields):
        self.Range = SimpleNamespace(Text=text, Fields=SimpleNamespace(Count=fields))


class FakeSection:
    def __init__(self, footer_text="", fields=0):
        self._footer = FakeStory(footer_text, fields)
        self._header = FakeStory("", 0)

    def Footers(self, kind):
        return self._footer

    def Headers(self, kind):
        return self._header


class FakeCollection(list):
    @property
    def Count(self):
        return len(self)

    def __call__(self, index):
        return self[index - 1]


class FakeDocument:
    def __init__(self, paragraphs, sections=None):
        self.Paragraphs = FakeCollection(paragraphs)
        self.Sections = FakeCollection(sections or [FakeSection()])

    def Repaginate(self):
        return None

    def ComputeStatistics(self, kind):
        pages = [p.Range.Information(3) for p in self.Paragraphs if p.Range.Text.strip()]
        return max(pages) if pages else 0


def _headings(*items):
    return [(index, level, document.Paragraphs(index)) for index, level, document in items]


def test_read_page_facts_collects_pages_and_section_ranges():
    document = FakeDocument([FakeParagraph("第一章", 1), FakeParagraph("正文", 1),
                             FakeParagraph("第二章", 2)], sections=[FakeSection()])
    facts = read_page_facts(document)
    assert facts["total_pages"] == 2
    assert facts["pages"][1] == 1 and facts["pages"][3] == 2
    assert facts["sections"] == [{"index": 1, "first_page": 1, "last_page": 2}]


def test_heading_separated_from_body_is_reported():
    document = FakeDocument([FakeParagraph("第一章 绪论", 1), FakeParagraph("", 1),
                             FakeParagraph("正文内容。", 2)])
    findings = check_heading_pagination(document, _headings((1, 1, document)), read_page_facts(document))
    assert len(findings) == 1
    assert findings[0].rule_id == "054.1"
    assert "不在同一页" in findings[0].title or "同页" in findings[0].expected
    assert findings[0].paragraph_index == 1


def test_heading_on_same_page_as_body_is_not_reported():
    document = FakeDocument([FakeParagraph("第一章 绪论", 1), FakeParagraph("正文内容。", 1)])
    assert not check_heading_pagination(document, _headings((1, 1, document)), read_page_facts(document))


def test_toc_heading_is_not_treated_as_body_heading_pagination_error():
    document = FakeDocument([
        FakeParagraph("目    录", 9),
        FakeParagraph("", 10),
        FakeParagraph("第一章 绪论\t1", 11),
    ])
    findings = check_heading_pagination(
        document, _headings((1, 1, document)), read_page_facts(document)
    )
    assert findings == []


def test_heading_uses_following_paragraph_start_page_not_its_end_page():
    document = FakeDocument([
        FakeParagraph("2.1.1 第一阶段", 13),
        FakeParagraph("正文从第13页开始并延续到第14页。", 14),
    ])
    facts = {"pages": {1: 13, 2: 14}, "start_pages": {1: 13, 2: 13}}
    assert not check_heading_pagination(
        document, _headings((1, 3, document)), facts)


def test_chapter_not_starting_new_page_is_reported_when_template_breaks():
    document = FakeDocument([FakeParagraph("第1节", 1), FakeParagraph("第一节正文", 2),
                             FakeParagraph("第二章 方法", 2)])
    facts = read_page_facts(document)
    headings = _headings((3, 1, document))
    assert check_chapter_page_breaks(document, headings, facts, template_breaks=False) == []
    findings = check_chapter_page_breaks(document, headings, facts, template_breaks=True)
    assert len(findings) == 1 and findings[0].rule_id == "030.1"


def test_static_page_number_in_footer_is_reported():
    document = FakeDocument([FakeParagraph("正文", 1)],
                            sections=[FakeSection(footer_text="1", fields=0)])
    findings = verify_page_numbering(document, read_page_facts(document))
    assert len(findings) == 1 and findings[0].rule_id == "040.2"
    assert "固定数字" in findings[0].actual
    assert findings[0].page_story == "footer"
    assert findings[0].page_variant == "default"


def test_automatic_page_number_is_not_reported():
    document = FakeDocument([FakeParagraph("正文", 1)],
                            sections=[FakeSection(footer_text="1", fields=1)])
    assert verify_page_numbering(document, read_page_facts(document)) == []


def _toc_document(items):
    paragraphs = []
    for text, page, level, lines in items:
        paragraph = FakeParagraph(text, page)
        paragraph.Style = SimpleNamespace(NameLocal=f"TOC {level}")
        paragraph.Range.ComputeStatistics = lambda kind, count=lines: count
        paragraphs.append(paragraph)
    return FakeDocument(paragraphs)


def test_blank_toc_entry_is_reported():
    document = _toc_document([("第一章 绪论\t1", 1, 1, 1), ("", 1, 2, 1)])
    findings = check_toc_layout(document, read_page_facts(document))
    assert any(finding.rule_id == "065.2" for finding in findings)


def test_empty_toc_field_boundary_is_not_an_entry():
    document = _toc_document([("", 1, 1, 1), ("第一章 绪论\t1", 1, 1, 1)])
    document.Paragraphs(1).Range.Fields.Count = 3
    findings = check_toc_layout(document, read_page_facts(document))
    assert not any(finding.rule_id == "065.2" for finding in findings)


def test_multi_line_toc_entry_is_reported():
    document = _toc_document([("第一章 绪论\t1", 1, 1, 2)])
    findings = check_toc_layout(document, read_page_facts(document))
    assert any(finding.rule_id == "065.4" for finding in findings)
