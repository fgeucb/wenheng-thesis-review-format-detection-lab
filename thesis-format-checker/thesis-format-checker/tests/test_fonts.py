from types import SimpleNamespace

from thesis_format_checker.core.effective_format import (
    SpanFormat,
    iter_script_spans,
    merge_span_formats,
    resolve_effective_font,
)
from thesis_format_checker.detectors.word_effective.fonts import compare_fonts


def test_mixed_chinese_latin_spans():
    assert list(iter_script_spans("中文 ABC123。")) == [
        (0, 2, "chinese"),
        (3, 9, "latin"),
    ]


def test_reports_chinese_and_latin_font_errors():
    spans = [
        SpanFormat(0, 2, "chinese", "黑体", 12),
        SpanFormat(2, 5, "latin", "Arial", 12),
    ]
    findings = compare_fonts(spans, "宋体", "Times New Roman", 3, "中文ABC")
    assert [item.rule_id for item in findings] == ["067", "068"]


def test_adjacent_equal_spans_are_merged():
    spans = merge_span_formats(
        [
            SpanFormat(0, 1, "chinese", "宋体", 12),
            SpanFormat(1, 2, "chinese", "宋体", 12),
        ]
    )
    assert len(spans) == 1 and spans[0].end == 2


class FakeThemeFonts:
    def __init__(self, latin: str, east_asian: str):
        self.names = {1: latin, 3: east_asian}

    def Item(self, index: int):
        return SimpleNamespace(Name=self.names[index])


def _document_theme():
    scheme = SimpleNamespace(
        MajorFont=FakeThemeFonts("Times New Roman", "黑体"),
        MinorFont=FakeThemeFonts("Arial", "宋体"),
    )
    return SimpleNamespace(DocumentTheme=SimpleNamespace(ThemeFontScheme=scheme))


def test_resolves_word_theme_placeholders_to_concrete_font_names():
    document = _document_theme()
    font = SimpleNamespace(ThemeFont=1)
    assert resolve_effective_font(document, font, "chinese", "+中文标题") == "黑体"
    assert resolve_effective_font(document, font, "latin", "+西文标题") == "Times New Roman"
    assert resolve_effective_font(document, font, "chinese", "+中文正文") == "宋体"
    assert resolve_effective_font(document, font, "latin", "+西文正文") == "Arial"


def test_unresolvable_theme_placeholder_is_not_reported_as_a_font_name():
    broken_document = SimpleNamespace()
    font = SimpleNamespace(ThemeFont=1)
    assert resolve_effective_font(broken_document, font, "chinese", "+中文标题") is None
    assert resolve_effective_font(broken_document, font, "latin", "Times New Roman") == "Times New Roman"
