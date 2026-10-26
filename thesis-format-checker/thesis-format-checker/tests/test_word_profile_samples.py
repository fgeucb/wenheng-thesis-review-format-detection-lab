"""模板剖面逐段基准测试：用伪 Word 对象验证“同层级 + 同位置”的取样，无需打开 Word。"""

from types import SimpleNamespace

from thesis_format_checker.core.scope import ParagraphInfo, collect_body_paragraphs
from thesis_format_checker.detectors.word_effective.headings import collect_headings
from thesis_format_checker.detectors.word_effective.profile import build_template_profile


class FakeCollection:
    def __init__(self, items):
        self._items = list(items)
        self.Count = len(self._items)

    def __call__(self, index):
        return self._items[index - 1]


class FakeFont:
    def __init__(self, chinese="宋体", latin="Times New Roman", size=12.0, bold=False):
        self.NameFarEast = chinese
        self.NameAscii = latin
        self.Size = size
        self.Bold = bold
        self.ThemeFont = 0


class FakeRange:
    def __init__(self, text, start, font, style_name):
        self.Text = text
        self.Start = start
        self.Font = font
        self.Style = SimpleNamespace(NameLocal=style_name, Name=style_name)

    def Information(self, kind):
        return False


class FakeFormat:
    def __init__(self, first_line=24.0, characters=2.0, line_rule=0, line=18.0, before=0.0, after=0.0):
        self.FirstLineIndent = first_line
        self.CharacterUnitFirstLineIndent = characters
        self.LineSpacingRule = line_rule
        self.LineSpacing = line
        self.SpaceBefore = before
        self.SpaceAfter = after
        self.Alignment = 3


class FakeParagraph:
    def __init__(self, index, text, font, style_name="正文", outline_level=10, fmt=None):
        self.Range = FakeRange(text, index * 100, font, style_name)
        self.Format = fmt if fmt is not None else FakeFormat()
        self.OutlineLevel = outline_level


class FakePageSetup:
    def __init__(self):
        self.TopMargin = 72.0
        self.BottomMargin = 72.0
        self.LeftMargin = 72.0
        self.RightMargin = 72.0
        self.DifferentFirstPageHeaderFooter = False
        self.OddAndEvenPagesHeaderFooter = False


class FakeSection:
    def __init__(self):
        self.PageSetup = FakePageSetup()
        self.Headers = lambda kind: SimpleNamespace(Range=SimpleNamespace(Text=""))


class FakeDocument:
    """只实现 build_template_profile 会用到的最小接口。"""

    def __init__(self, paragraphs):
        self.Paragraphs = FakeCollection(paragraphs)
        self.Sections = FakeCollection([FakeSection()])
        self._spans = [
            (item.Range.Start, item.Range.Start + max(len(item.Range.Text), 1), item.Range.Font)
            for item in paragraphs
        ]

    def Range(self, start, end):
        font = next((item_font for first, last, item_font in self._spans if first <= start < last), None)
        return SimpleNamespace(Font=font)


def body_info(paragraph, style_name="正文"):
    return ParagraphInfo(index=1, paragraph=paragraph, text=paragraph.Range.Text, style_name=style_name)


def test_profile_keeps_per_paragraph_samples_when_template_has_several_styles():
    first_body = FakeParagraph(2, "正文一段。", FakeFont(size=12.0))
    second_body = FakeParagraph(3, "正文二段。", FakeFont(size=10.5))
    title = FakeParagraph(1, "第一章 绪论", FakeFont(size=16.0, bold=True), style_name="标题 1", outline_level=1)
    closing = FakeParagraph(4, "第三章 结论", FakeFont(size=18.0, bold=True), style_name="标题 1", outline_level=1)
    document = FakeDocument([title, first_body, second_body, closing])

    profile = build_template_profile(document, [body_info(first_body), body_info(second_body)])

    # 逐段基准保留每个段落自己的格式，而不是在全篇强行选一个标准。
    assert [sample.size_pt for sample in profile.body_samples] == [12.0, 10.5]
    assert [sample.style_name for sample in profile.body_samples] == ["正文", "正文"]
    assert [sample.first_line_indent.characters for sample in profile.body_samples] == [2.0, 2.0]
    # 12 磅与 10.5 磅各占一半：全篇主导值确实无法确定，但这不再阻塞比对。
    assert profile.font_size_pt is None
    assert any("逐段比对" in reason for reason in profile.unknown_reasons)


def test_profile_records_heading_samples_in_order_without_forcing_one_standard():
    first = FakeParagraph(1, "第一章 绪论", FakeFont(size=16.0, bold=True), style_name="标题 1", outline_level=1)
    second = FakeParagraph(2, "第一章 方法", FakeFont(size=18.0, bold=True), style_name="标题 1", outline_level=1)
    document = FakeDocument([first, second])

    profile = build_template_profile(document, [body_info(first, "标题 1")])

    assert [(sample.index, sample.level) for sample in profile.heading_samples] == [(1, 1), (2, 1)]
    assert [sample.spec.size_pt for sample in profile.heading_samples] == [16.0, 18.0]
    # 不再输出“模板格式存在冲突，未确定标准”这类阻塞性说明。
    assert not any("未确定标准" in reason for reason in profile.unknown_reasons)
    assert 1 not in profile.headings


def test_word_path_headings_skip_cover_and_declaration_before_abstract():
    """封面/声明（摘要之前）不作为标题样本：它们只有封面规范，没有章节标题规范。"""

    declaration = FakeParagraph(1, "南京农业大学本科毕业论文（设计）原创性声明",
                                FakeFont(size=18.0, bold=True), style_name="标题 1", outline_level=1)
    cover = FakeParagraph(2, "基于深度学习的××研究", FakeFont(size=22.0, bold=True),
                          style_name="标题 1", outline_level=1)
    abstract = FakeParagraph(3, "摘  要", FakeFont(size=16.0, bold=True), style_name="标题 1", outline_level=1)
    chapter = FakeParagraph(4, "第一章 绪论", FakeFont(size=16.0, bold=True), style_name="标题 1", outline_level=1)
    document = FakeDocument([declaration, cover, abstract, chapter])

    assert [index for index, _, _ in collect_headings(document)] == [3, 4]


def test_word_path_body_samples_exclude_cover_and_declaration():
    """声明正文不能当正文格式基准，即使它的段落在样式上就是“正文”。"""

    declaration = FakeParagraph(1, "南京农业大学本科毕业论文（设计）原创性声明",
                                FakeFont(), style_name="标题 1", outline_level=1)
    statement = FakeParagraph(2, "本人郑重声明：所呈交的论文是独立完成的研究成果。",
                              FakeFont(), style_name="正文")
    abstract = FakeParagraph(3, "摘  要", FakeFont(), style_name="标题 1", outline_level=1)
    chapter = FakeParagraph(4, "第一章 绪论", FakeFont(), style_name="标题 1", outline_level=1)
    body = FakeParagraph(5, "研究背景正文内容。", FakeFont(), style_name="正文")
    document = FakeDocument([declaration, statement, abstract, chapter, body])

    assert [info.index for info in collect_body_paragraphs(document, None, "document")] == [5]


def test_word_path_body_samples_exclude_captions_and_numbered_headings():
    """图表题注与编号标题都不是正文：题注五号字、标题三号字，混进正文就会报“正文字号”。"""

    caption_cn = FakeParagraph(1, "图2-2  模型整体流程图", FakeFont(size=10.5), style_name="正文")
    caption_en = FakeParagraph(2, "Fig.2-2  The overall model flow diagram",
                               FakeFont(size=10.5), style_name="正文")
    heading = FakeParagraph(3, "2.2.1 跨行业注意力权重的市值分层特征", FakeFont(size=12.0), style_name="正文")
    chapter = FakeParagraph(4, "第一章 绪论", FakeFont(size=16.0), style_name="标题 1", outline_level=1)
    body = FakeParagraph(5, "研究背景正文内容，包含足够字符以便样式判定。", FakeFont(size=12.0), style_name="正文")
    document = FakeDocument([caption_cn, caption_en, heading, chapter, body])

    assert [info.index for info in collect_body_paragraphs(document, None, "document")] == [5]
