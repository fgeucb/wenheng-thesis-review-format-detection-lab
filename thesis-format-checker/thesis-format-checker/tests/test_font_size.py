from thesis_format_checker.core.effective_format import SpanFormat
from thesis_format_checker.detectors.word_effective.font_size import compare_font_sizes


def test_size_tolerance_and_direct_override():
    spans = [
        SpanFormat(0, 3, "chinese", "宋体", 12.05),
        SpanFormat(3, 6, "latin", "Times New Roman", 10.5),
    ]
    findings = compare_font_sizes(spans, 12, 1, "正文ABC", tolerance_pt=0.1)
    assert len(findings) == 1
    assert findings[0].actual == 10.5

