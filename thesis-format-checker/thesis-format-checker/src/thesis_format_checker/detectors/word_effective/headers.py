from __future__ import annotations

from thesis_format_checker.core.effective_format import clean_word_text
from thesis_format_checker.core.models import Finding, Status

HEADER_TYPES = (1, 2, 3)  # primary, first page, even page
HEADER_NAMES = {1: "普通页", 2: "首页", 3: "偶数页"}


def read_headers(section) -> dict[int, str]:
    result: dict[int, str] = {}
    different_first = bool(section.PageSetup.DifferentFirstPageHeaderFooter)
    different_even = bool(section.PageSetup.OddAndEvenPagesHeaderFooter)
    for kind in HEADER_TYPES:
        if (kind == 2 and not different_first) or (kind == 3 and not different_even):
            result[kind] = result[1]
            continue
        result[kind] = clean_word_text(section.Headers(kind).Range.Text).strip()
    result["different_first"] = different_first
    result["different_even"] = different_even
    return result


def compare_headers(expected: dict[int, str], actual: dict[int, str], section_index: int) -> list[Finding]:
    findings = []
    for kind in HEADER_TYPES:
        if kind == 2 and actual.get("different_first") is False:
            continue
        if kind == 3 and actual.get("different_even") is False:
            continue
        a, b = expected.get(kind, ""), actual.get(kind, "")
        if a == b:
            continue
        findings.append(Finding(
            "035", "页眉文字", Status.FAIL,
            f"【035 页眉文字】第 {section_index} 节{HEADER_NAMES[kind]}页眉：模板 A 为“{a or '空'}”；文档 B 为“{b or '空'}”。",
            # 模板没有页眉时明确写“无”，不要留空让用户猜。
            expected=(a or "无（模板该节没有页眉）"), actual=(b or "无（该节没有页眉）"),
            section_index=section_index, header_kind=kind,
            page_story="header", page_scope="variant",
            page_variant={1: "default", 2: "first", 3: "even"}[kind],
        ))
    return findings
