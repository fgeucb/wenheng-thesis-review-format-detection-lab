"""需要 Word 排版结果才能判定的规则：标题同页（054）、章节另起页（030）、页码（040）。

这些规则在 OOXML 里只能给出“无法判定”，本模块用 :mod:`core.word_pages`
的分页事实（物理页码、节页范围、总页数）给出确定结论。
"""
from __future__ import annotations

import re
from collections import defaultdict

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.effective_format import clean_word_text
from thesis_format_checker.core.models import Finding, Status
from thesis_format_checker.core.word_pages import (
    paragraph_page, paragraph_start_page, section_pages,
)

TITLES = {
    "030.1": "章节未按模板另起页",
    "040.2": "页码重复或跳号",
    "054.1": "标题与后续内容不在同一页",
    "065.1": "目录存在孤立条目",
    "065.2": "目录存在空白条目",
    "065.4": "目录条目存在错误换行",
}

_TOC_STYLE = re.compile(r"^(?:toc|目录)\s*\d", re.I)
_TOC_HEADINGS = {
    "目录", "contents", "图目录", "表目录", "插图目录", "附表目录",
    "listoffigures", "listoftables",
}


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "").casefold()


def title_of(code: str) -> str:
    item = FINE_RULES.get(code)
    return item.title if item else TITLES[code]


def _paragraph_text(document, index: int, word_facts=None) -> str:
    """段落文本（去掉段落标记）；``word_facts`` 存在时直接命中事实表缓存。"""

    if word_facts is not None:
        return word_facts.clean_text(index).strip()
    try:
        return clean_word_text(document.Paragraphs(index).Range.Text).strip()
    except Exception:
        return ""


def check_heading_pagination(document, headings, facts, *, word_facts=None) -> list[Finding]:
    """054：标题不应与后续内容分离（标题留在页末、正文从下一页开始）。

    ``headings`` 来自 :func:`collect_headings`（Word 段号 + 级别 + 段落）。
    ``facts`` 可以是 :class:`~thesis_format_checker.core.word_pages.PageFacts`
    （按需查询）或旧的分页事实字典。
    """

    findings: list[Finding] = []
    total = int(word_facts.count) if word_facts is not None else int(document.Paragraphs.Count)
    for index, level, paragraph in headings:
        heading_text = _paragraph_text(document, index, word_facts)
        # “目录”是前置部分标题，其后可能有自动分页、空白偶数页或 TOC 域边界；
        # 不能套用正文标题“与下段同页”的规则。
        if _compact(heading_text) in _TOC_HEADINGS:
            continue
        page = paragraph_start_page(facts, index)
        if page is None:
            continue
        following = None
        for offset in range(index + 1, total + 1):
            text = _paragraph_text(document, offset, word_facts)
            if text:
                following = offset
                break
        if following is None:
            continue
        next_page = paragraph_start_page(facts, following)
        if next_page is None or next_page == page:
            continue
        findings.append(Finding(
            "054.1", title_of("054.1"), Status.FAIL,
            f"标题“{heading_text[:20]}”落在第 {page} 页末尾，后续内容从第 {next_page} 页开始，"
            "标题与正文被分页拆开；请让标题与后续内容保持在同一页。",
            expected="标题与其后的正文在同一页",
            actual=f"标题在第 {page} 页末，后续内容在第 {next_page} 页",
            paragraph_index=index, excerpt=heading_text[:30],
            section_index=int(getattr(paragraph.Range, "Information", lambda _x: 0)(2) or 0) or None,
        ))
    return findings


def check_chapter_page_breaks(document, headings, facts, template_breaks: bool,
                               *, word_facts=None) -> list[Finding]:
    """030：模板的章节设置“另起页”时，论文每章也必须另起页。"""

    if not template_breaks:
        return []
    findings: list[Finding] = []
    for index, level, paragraph in headings:
        if level != 1:
            continue
        page = paragraph_page(facts, index)
        if page is None or index <= 1:
            continue
        previous = index - 1
        while previous > 1 and not _paragraph_text(document, previous, word_facts):
            previous -= 1
        previous_page = paragraph_page(facts, previous)
        if previous_page is None or previous_page != page:
            continue
        text = _paragraph_text(document, index, word_facts)
        findings.append(Finding(
            "030.1", title_of("030.1"), Status.FAIL,
            f"章节“{text[:20]}”紧接上一页内容开始；模板要求此类章节另起一页。",
            expected="章节从新的一页开始",
            actual=f"章节标题与上一段同在第 {page} 页",
            paragraph_index=index, excerpt=text[:30],
        ))
    return findings


def verify_page_numbering(document, facts: dict) -> list[Finding]:
    """040：用排版后的节页范围核对页码是否自动、连续。

    自动页码（PAGE 域）由 Word 计算，同一页内不会重复或跳号；本函数据此给出
    “已核对”的结论（不写批注），只在发现**固定数字冒充页码**时报告——那种写法
    每页都会显示同一个数字，是最常见的“重复/跳号”来源。
    """

    findings: list[Finding] = []
    for index in range(1, int(document.Sections.Count) + 1):
        first_page, last_page = section_pages(facts, index)
        if first_page is None or last_page is None:
            continue
        for area, label in (("footer", "页脚"), ("header", "页眉")):
            story = _story(document, index, area)
            if story is None:
                continue
            text, fields = story
            digits = "".join(ch for ch in text if ch.isdigit())
            if fields or not digits:
                continue
            findings.append(Finding(
                "040.2", title_of("040.2"), Status.FAIL,
                f"第 {index} 节{label}里的“{digits[:6]}”是手工输入的数字，不是自动页码域；"
                "各页都会显示同一个数字，看起来就是重复页码。请改用“插入页码”生成自动页码。",
                expected="每一页显示当前页码（自动页码域）",
                actual=f"第 {index} 节{label}为固定数字“{digits[:6]}”，共 {last_page - first_page + 1} 页",
                section_index=index,
                page_story=area, page_scope="variant", page_variant="default", header_kind=1,
            ))
    return findings


def _story(document, index: int, area: str):
    """取某节页眉/页脚（普通页变体）的纯文字与域数量；读不到返回 None。"""

    try:
        container = (document.Sections(index).Footers(1) if area == "footer"
                     else document.Sections(index).Headers(1))
    except Exception:
        return None
    try:
        text = clean_word_text(container.Range.Text).strip()
        fields = int(container.Range.Fields.Count)
    except Exception:
        return None
    return text, fields


def check_toc_layout(document, facts, *, word_facts=None, page_facts=None) -> list[Finding]:
    """065：目录条目结构——孤立条目、空白条目、条目与页码跨页拆开/错误换行。

    这些都要排版结果才能判定：条目占几行、是否被分页拆开、某页是否只剩一条。
    """

    findings: list[Finding] = []
    entries: list[tuple[int, int | None, int | None, int, str]] = []
    total = int(word_facts.count) if word_facts is not None else int(document.Paragraphs.Count)
    for index in range(1, total + 1):
        if word_facts is not None:
            style = word_facts.style_name(index)
            paragraph = word_facts.paragraph(index)
        else:
            paragraph = document.Paragraphs(index)
            try:
                style = str(paragraph.Style.NameLocal or "")
            except Exception:
                continue
        if not _TOC_STYLE.match(style):
            continue
        text = _paragraph_text(document, index, word_facts)
        if not text:
            # Word exposes the paragraph that owns the TOC field as an empty
            # TOC-styled paragraph. Its Range.Fields collection contains the
            # complete field result, so it is a field boundary, not an entry.
            try:
                field_count = int(paragraph.Range.Fields.Count)
            except Exception:
                field_count = 0
            if field_count:
                continue
            findings.append(Finding(
                "065.2", title_of("065.2"), Status.FAIL,
                f"第 {index} 段是空的目录条目（只有条目样式没有内容），多出来的空行会让目录页码错位。",
                expected="目录条目应有标题与页码", actual="存在没有内容的目录条目",
                paragraph_index=index, excerpt=""))
            continue
        if word_facts is not None:
            position = word_facts.start(index)
            if position is None:
                continue
            # Range.End = Range.Start + len(Range.Text)，与逐段读取完全等价。
            end_position = position + max(len(word_facts.text(index)), 1) - 1
        else:
            position = int(paragraph.Range.Start)
            end_position = max(int(paragraph.Range.End) - 1, position)
        start = (_page_at(page_facts, position) if page_facts is not None
                 else _page_at(document, position))
        end = (_page_at(page_facts, end_position) if page_facts is not None
               else _page_at(document, end_position))
        lines = _line_count(paragraph)
        entries.append((index, start, end, lines, text))
    for index, start, end, lines, text in entries:
        if start is not None and end is not None and start != end:
            findings.append(Finding(
                "065.4", title_of("065.4"), Status.FAIL,
                f"目录条目“{text[:24]}”被分页拆开：前半在第 {start} 页、页码在后一页，"
                "打印后会被页码分隔线切断。",
                expected="每条目录条目与其页码在同一页",
                actual=f"条目跨第 {start}–{end} 页", paragraph_index=index, excerpt=text[:30]))
        elif lines > 1:
            findings.append(Finding(
                "065.4", title_of("065.4"), Status.FAIL,
                f"目录条目“{text[:24]}”占 {lines} 行，超出模板的单行条目排版。",
                expected="目录条目为单行（超出部分由标题本身控制）",
                actual=f"该条目占 {lines} 行", paragraph_index=index, excerpt=text[:30]))
    by_page: dict[int, list[int]] = defaultdict(list)
    for index, start, _, _, _ in entries:
        if start:
            by_page[start].append(index)
    for page, items in sorted(by_page.items()):
        if len(items) != 1 or len(by_page) < 2:
            continue
        index = items[0]
        text = next(text for item, _, _, _, text in entries if item == index)
        findings.append(Finding(
            "065.1", title_of("065.1"), Status.FAIL,
            f"第 {page} 页只剩一条目录条目“{text[:24]}”，与其余条目分离（孤立条目）。",
            expected="目录条目连续排列、不单独占一页",
            actual=f"第 {page} 页只有这一条目录条目", paragraph_index=index, excerpt=text[:30]))
    del facts
    return findings


def _page_at(source, position: int):
    """字符位置 → 页码；``source`` 可以是 Word 文档，也可以是分页事实对象。"""

    if hasattr(source, "page_at"):
        return source.page_at(position)
    try:
        return int(source.Range(position, position + 1).Information(3))
    except Exception:
        return None


def _line_count(paragraph) -> int:
    try:
        return int(paragraph.Range.ComputeStatistics(1))  # wdStatisticLines
    except Exception:
        return 0
