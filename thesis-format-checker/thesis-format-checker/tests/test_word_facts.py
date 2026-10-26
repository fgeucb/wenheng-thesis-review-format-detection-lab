"""Word 段落事实表与分页事实的回归测试（性能优化 Tier 0）。

两件事必须锁死：

1. **值一致**：事实表读到的文本/样式/大纲级别/是否在表格/页码/节号，
   与逐段直接访问 COM 完全一致；
2. **只读一次**：同一属性重复查询命中缓存，COM 访问次数不随消费者增加。
"""

from __future__ import annotations

from types import SimpleNamespace

from thesis_format_checker.core.scope import collect_body_paragraphs
from thesis_format_checker.core.word_facts import WordFacts
from thesis_format_checker.core.word_pages import PageFacts, paragraph_page, read_page_facts
from thesis_format_checker.detectors.word_effective.headings import collect_headings


# ---------------------------------------------------------------------------
# 最小可用的 Word COM 替身（按页/段组织，足够覆盖事实表用到的接口）
# ---------------------------------------------------------------------------

class FakeRange:
    def __init__(self, document, start: int, end: int, *, text="", page=None, section=1,
                 in_table=False) -> None:
        self._document = document
        self.Start = start
        self.End = end
        self.Text = text
        self._page = page
        self._section = section
        self._in_table = in_table
        self.Style = SimpleNamespace(NameLocal="正文")
        self.InlineShapes = SimpleNamespace(Count=0)
        self.Shapes = SimpleNamespace(Count=0)

    def Information(self, kind):
        if kind == 3:
            return self._page
        if kind == 2:
            return self._section
        if kind == 12:
            return self._in_table
        return 0


class FakeParagraph:
    def __init__(self, document, index: int, text: str, *, page: int, section: int = 1,
                 style: str = "正文", outline: int = 10, in_table: bool = False) -> None:
        self._document = document
        self._index = index
        self._text = text
        self._page = page
        self._section = section
        self._style = style
        self.outline = outline
        self._in_table = in_table
        self._range = None

    @property
    def Range(self):
        if self._range is None:
            start = self._document.paragraph_start(self._index)
            self._range = FakeRange(self._document, start, start + len(self._text) + 1,
                                    text=self._text + "\r", page=self._page,
                                    section=self._section, in_table=self._in_table)
            self._range.Style = SimpleNamespace(NameLocal=self._style)
        return self._range

    @property
    def OutlineLevel(self):
        return self.outline

    @property
    def Style(self):
        """模拟 ``Paragraph.Style``：目录域空段的 ``Range.Style`` 为 None，这里仍有样式。"""

        return SimpleNamespace(NameLocal=self._style)


class FakeCollection(list):
    @property
    def Count(self):
        return len(self)

    def __call__(self, index):
        return self[index - 1]


class FakeContent:
    def __init__(self, document) -> None:
        self._document = document

    @property
    def Start(self):
        return 0

    @property
    def End(self):
        return self._document.end

    @property
    def Text(self):
        return "".join(paragraph._text + "\r" for paragraph in self._document.Paragraphs)


class FakeDocument:
    """由 (文本, 页码, 样式, 大纲级别, 是否在表格) 列表构成的假文档。"""

    def __init__(self, rows) -> None:
        self._rows = rows
        self._starts = []
        position = 0
        for text, *_ in rows:
            self._starts.append(position)
            position += len(text) + 1
        self.end = position
        self.Paragraphs = FakeCollection([
            FakeParagraph(self, index, row[0], page=row[1],
                          style=(row[2] if len(row) > 2 else "正文"),
                          outline=(row[3] if len(row) > 3 else 10),
                          in_table=(row[4] if len(row) > 4 else False))
            for index, row in enumerate(rows, start=1)])
        self.Content = FakeContent(self)
        self.Sections = FakeCollection([SimpleNamespace(PageSetup=SimpleNamespace())])

    def paragraph_start(self, index: int) -> int:
        return self._starts[index - 1]

    def Repaginate(self) -> None:
        return None

    def ComputeStatistics(self, kind: int):
        if kind == 2:
            return max(row[1] for row in self._rows)
        return 0

    def GoTo(self, what, which, count):
        target = None
        for index, row in enumerate(self._rows, start=1):
            if row[1] == count:
                target = index
                break
        if target is None:
            target = len(self._rows)
        return SimpleNamespace(Start=self._starts[target - 1])

    def Range(self, start, end):
        texts = []
        page = None
        for index, row in enumerate(self._rows, start=1):
            position = self._starts[index - 1]
            if start <= position < end:
                texts.append(row[0] + "\r")
            if position <= start < position + len(row[0]) + 1 and page is None:
                page = row[1]
        return FakeRange(self, start, end, text="".join(texts), page=page, section=1)


def _document():
    return FakeDocument([
        ("封面题目", 1, "标题", 1),
        ("摘要内容", 1, "正文1", 10),
        ("第一章 绪论", 2, "标题 1", 1),
        ("这是正文第一段。", 2, "正文1", 10),
        ("", 2, "正文1", 10),
        ("表 2-1 数据", 3, "题注1", 10, True),
        ("第三章 方法", 3, "标题 1", 1),
    ])


# ---------------------------------------------------------------------------
# WordFacts：值与缓存
# ---------------------------------------------------------------------------

def test_word_facts_match_direct_com_reads():
    document = _document()
    facts = WordFacts(document, label="test")
    assert facts.count == 7
    for index in range(1, facts.count + 1):
        paragraph = document.Paragraphs(index)
        assert facts.text(index) == paragraph.Range.Text
        assert facts.start(index) == paragraph.Range.Start
        assert facts.style_name(index) == paragraph.Range.Style.NameLocal
        assert facts.outline_level(index) == paragraph.OutlineLevel
        assert facts.in_table(index) == bool(paragraph.Range.Information(12))


def test_word_facts_reads_each_attribute_once():
    document = _document()
    facts = WordFacts(document)
    for index in facts.indices():
        facts.text(index)
        facts.style_name(index)
        facts.outline_level(index)
        facts.in_table(index)
    first_round = facts.calls
    # 再消费两轮（正文/标题/分页/空段…）：不应产生任何新的 COM 访问。
    for _ in range(2):
        for index in facts.indices():
            facts.text(index)
            facts.style_name(index)
            facts.outline_level(index)
            facts.in_table(index)
    assert facts.calls == first_round
    # 每段最多 7 次：段落对象 + Range + 4 个属性（段落数只查一次）
    assert first_round <= 1 + 7 * 7


def test_style_falls_back_to_paragraph_style_when_range_style_missing():
    """目录域里的空段 ``Range.Style`` 为 None：退回 ``Paragraph.Style``（历史口径）。

    样例论文的第 61/151 段就是这种情况；不回退会漏报"没有内容的目录条目"（065.2）。
    """

    document = _document()
    paragraph = document.Paragraphs(2)
    paragraph.Range.Style = None
    facts = WordFacts(document)
    assert facts.style_name(2) == "正文1"
    # 其余段落仍按 Range.Style 读，值不变
    assert facts.style_name(3) == "标题 1"


def test_body_and_heading_collection_identical_with_and_without_facts():
    document = _document()
    facts = WordFacts(document)
    plain = collect_body_paragraphs(document, None, "document")
    cached = collect_body_paragraphs(document, None, "document", facts=facts)
    assert [(item.index, item.text, item.style_name) for item in plain] == \
           [(item.index, item.text, item.style_name) for item in cached]
    plain_headings = collect_headings(document)
    cached_headings = collect_headings(document, facts=facts)
    assert [(index, level) for index, level, _ in plain_headings] == \
           [(index, level) for index, level, _ in cached_headings]


# ---------------------------------------------------------------------------
# PageFacts：页码/节号/空白页
# ---------------------------------------------------------------------------

def test_page_facts_matches_reference_implementation():
    document = _document()
    facts = WordFacts(document)
    pagination = PageFacts(document, facts)
    reference = read_page_facts(document)
    assert pagination.total_pages == reference["total_pages"]
    for index in range(1, facts.count + 1):
        assert pagination.page_of(index) == reference["pages"][index]
        assert pagination.section_of(index) == reference["section_of"][index]
    assert paragraph_page(pagination, 3) == paragraph_page(reference, 3)


def test_page_facts_page_at_uses_page_boundaries():
    document = _document()
    pagination = PageFacts(document, WordFacts(document))
    assert pagination.page_at(0) == 1
    assert pagination.page_at(document.paragraph_start(4)) == 2
    assert pagination.page_at(document.paragraph_start(7)) == 3


def test_page_facts_blank_pages_come_from_page_ranges():
    document = _document()
    pagination = PageFacts(document, WordFacts(document))
    # 三页都有内容（第 1 页封面/摘要，第 2 页正文，第 3 页题注与章节）
    assert pagination.blank_pages() == []
    assert pagination.content_pages() == [1, 2, 3]
    payload = pagination.as_dict()
    assert payload["total_pages"] == 3 and payload["blank_pages"] == []
    assert payload["sections"]


def test_page_facts_blank_page_is_reported_when_page_has_no_content():
    document = FakeDocument([
        ("第一页内容", 1),
        ("", 2),          # 第 2 页只有一个空段落 → 空白页
        ("第三页内容", 3),
    ])
    pagination = PageFacts(document, WordFacts(document))
    assert pagination.blank_pages() == [2]
