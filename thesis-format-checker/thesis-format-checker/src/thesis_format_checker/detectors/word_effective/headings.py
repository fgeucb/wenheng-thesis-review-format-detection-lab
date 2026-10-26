from __future__ import annotations

import re

from thesis_format_checker.core.effective_format import clean_word_text, normalized_number, resolve_effective_font
from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.models import Finding, HeadingSpec, Status, build_comment
from thesis_format_checker.core.scope import front_matter_end
# 与 docx 侧共用同一套“编号 / 像不像标题”判据（那里的实现更完整，这里只取纯文本部分）。
from thesis_format_checker.detectors.headings import (
    _CAPTION_LIKE as CAPTION_LIKE,
    _CHAPTER_PREFIX as CHAPTER_PREFIX,
    _looks_like_heading as looks_like_heading,
    _numbering_level as numbering_level,
)

#: 标题层级与属性 → 细粒度规则编号
_HEADING_CODES: dict[int, dict[str, str]] = {
    1: {"中文字体": "045.1", "英文字体": "045.2", "字号": "045.3",
        "加粗": "045.4", "对齐方式": "045.5"},
    2: {"中文字体": "046.1", "英文字体": "046.2", "字号": "046.3", "加粗": "046.4",
        "对齐方式": "046.5", "首行缩进": "046.6", "首行缩进磅值": "046.6",
        "段前间距": "046.7", "段后间距": "046.8"},
    3: {"中文字体": "047.1", "英文字体": "047.2", "字号": "047.3", "加粗": "047.4",
        "对齐方式": "047.5", "首行缩进": "047.6", "首行缩进磅值": "047.6",
        "段前间距": "047.7", "段后间距": "047.8"},
}


def _split_diff(text: str) -> tuple[str, str]:
    """把“模板 X，当前(文档) Y”拆成（应当, 实际）；无法解析时原样保留。"""

    match = re.search(r"模板\s*(.*?)[，,]\s*当前(?:文档)?\s*(.*)$", text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return "按模板要求", text
from thesis_format_checker.detectors.word_effective.first_line_indent import read_indent


ALIGNMENT_LABELS = {
    0: "左对齐",
    1: "居中",
    2: "右对齐",
    3: "两端对齐",
    4: "分散对齐",
    5: "中等程度两端对齐",
    7: "高度两端对齐",
    8: "低度两端对齐",
    9: "泰文两端对齐",
}

#: 目录域缓存段落（toc 1/2/3）：它们是文档结构的证据，本身不是正文标题。
_TOC_STYLE = re.compile(r"^(?:toc|目录)\s*\d*$", re.I)
#: 封面题目等文档标题样式（含中文 Word 的本地化名称“标题”）。
_DOCUMENT_TITLE_STYLES = {"title", "标题", "subtitle", "副标题", "题目"}


def _display(name: str, value) -> str:
    if name == "加粗":
        return "加粗" if value else "未加粗"
    if name == "对齐方式":
        return ALIGNMENT_LABELS[value]
    if name == "字号":
        return f"{value:g} 磅"
    if name == "首行缩进":
        return f"{value:g} 字符"
    if name in ("首行缩进磅值", "段前间距", "段后间距"):
        return f"{value:g} 磅"
    return str(value)


def _style_name(paragraph, facts=None, index: int | None = None) -> str:
    """段落样式名：优先走事实表缓存，否则逐段读取（同一 COM 接口，值一致）。"""

    if facts is not None and index is not None:
        return facts.style_name(index)
    from thesis_format_checker.core.scope import style_name_of

    return style_name_of(paragraph)


def _outline_level(paragraph, facts=None, index: int | None = None) -> int | None:
    """大纲级别：优先走事实表缓存。"""

    if facts is not None and index is not None:
        return facts.outline_level(index)
    try:
        return int(paragraph.OutlineLevel)
    except Exception:
        return None


def _style_level(paragraph, mapping_styles: dict[str, int] | None = None,
                 *, facts=None, index: int | None = None) -> int | None:
    """仅按样式与大纲级别的层级判断（旧逻辑）。"""

    name = (_style_name(paragraph, facts, index) or "").strip()
    if _TOC_STYLE.match(name) or name.casefold() in _DOCUMENT_TITLE_STYLES:
        return None
    for style, level in (mapping_styles or {}).items():
        if name.casefold() == style.casefold() and 1 <= int(level) <= 9:
            return int(level)
    match = re.search(r"(?:heading|标题)\s*([1-9])\b", name, re.I)
    if match:
        return int(match.group(1))
    level = _outline_level(paragraph, facts, index)
    return level if level is not None and 1 <= level <= 9 else None


def heading_level(paragraph, mapping_styles: dict[str, int] | None = None,
                  chapter_prefix: bool = False, *, facts=None, index: int | None = None) -> int | None:
    """标题级别：编号/“像标题”判据 + 样式证据，与 docx 侧同一口径。

    只按样式判断会把误用「标题 2」样式的整段正文当标题；这里先用正文编号
    （第X章 / N / N.N）和“短、无句末标点”确认它确实是标题。
    """

    name = (_style_name(paragraph, facts, index) or "").strip()
    if _TOC_STYLE.match(name) or name.casefold() in _DOCUMENT_TITLE_STYLES:
        return None  # 目录缓存与封面题目都不是章节标题（即使带编号）
    style_level = _style_level(paragraph, mapping_styles, facts=facts, index=index)
    try:
        text = (facts.clean_text(index) if facts is not None and index is not None
                else clean_word_text(paragraph.Range.Text)).strip()
    except Exception:
        return style_level  # 取不到文字时退回旧的样式判断
    if not text:
        return None
    if CAPTION_LIKE.match(text):
        return None
    if not looks_like_heading(text):
        return None
    numbered, strong = numbering_level(text, chapter_prefix)
    if numbered is not None and not strong and style_level is None:
        numbered = None
    return numbered or style_level


def collect_headings(document, mapping_styles: dict[str, int] | None = None, *, facts=None):
    """按层级树口径收集标题；``chapter_prefix`` 表示文档使用“第X章”编号体系。

    摘要之前的封面、题名页、声明页只由封面模块按模板同位置比对，不作为标题
    样本：那里的“原创性声明”一类标题是按封面规范排版的，不该套用一级标题格式。
    """

    paragraphs = []
    total = facts.count if facts is not None else int(document.Paragraphs.Count)
    for index in range(1, total + 1):
        if facts is not None:
            # 事实表已读过的属性直接命中缓存（值与逐段读取一致）。
            paragraph = facts.paragraph(index)
            text = facts.clean_text(index).strip()
        else:
            paragraph = document.Paragraphs(index)
            try:
                text = clean_word_text(paragraph.Range.Text).strip()
            except Exception:
                text = ""
        paragraphs.append((index, paragraph, text))
    front_end = front_matter_end([text for _, _, text in paragraphs])
    chapter_prefix = False
    for index, _, text in paragraphs:
        if 0 < front_end and index <= front_end:
            continue  # 前置区不参与标题体系判断
        if text and CHAPTER_PREFIX.match(text):
            chapter_prefix = True
            break
    result = []
    for index, para, text in paragraphs:
        if 0 < front_end and index <= front_end:
            continue  # 封面/声明等前置区不是章节标题
        if not text:
            continue
        if facts is not None:
            if facts.in_table(index):
                continue
        else:
            try:
                if bool(para.Range.Information(12)):  # wdWithInTable
                    continue
            except Exception:
                pass
        level = heading_level(para, mapping_styles, chapter_prefix, facts=facts, index=index)
        if level:
            result.append((index, level, para))
    return result


def read_heading(document, paragraph) -> HeadingSpec:
    font = paragraph.Range.Font
    bold_value = font.Bold
    bold = None if bold_value is None or int(bold_value) == 9999999 else bool(bold_value)
    indent = read_indent(paragraph)
    fmt = paragraph.Format
    return HeadingSpec(
        resolve_effective_font(document, font, "chinese", font.NameFarEast),
        resolve_effective_font(document, font, "latin", font.NameAscii),
        normalized_number(font.Size), bold, int(fmt.Alignment), indent,
        normalized_number(fmt.SpaceBefore), normalized_number(fmt.SpaceAfter),
    )


def _script_range(text: str, pattern: str) -> tuple[int, int] | None:
    match = re.search(pattern, text)
    if not match:
        return None
    return match.start(), match.end()


def compare_heading(expected: HeadingSpec, actual: HeadingSpec, level: int, index: int,
                    tolerance_pt: float = 0.5, text: str | None = None
                    ) -> tuple[list[Finding], list[str]]:
    """按属性逐个比较标题格式：一个属性一条结论（045.x / 046.x / 047.x）。"""

    rule = "045" if level == 1 else "046" if level == 2 else "047"
    level_name = "一二三四五六七八九"[level - 1] if 1 <= level <= 9 else str(level)
    label = f"{level_name}级标题"
    codes = _HEADING_CODES.get(level, _HEADING_CODES[3])
    script_ranges = {
        "中文字体": _script_range(text, r"[\u3400-\u9fff\U00020000-\U0003134f]+")
        if text is not None else (0, 2),
        "英文字体": _script_range(text, r"[A-Za-z0-9\u00c0-\u024f]+")
        if text is not None else (0, 2),
    }
    fields = [
        ("中文字体", expected.chinese_font, actual.chinese_font),
        ("英文字体", expected.latin_font, actual.latin_font),
        ("字号", expected.size_pt, actual.size_pt),
        ("加粗", expected.bold, actual.bold),
        ("对齐方式", expected.alignment, actual.alignment),
    ]
    if level >= 2:
        expected_indent = expected.indent
        actual_indent = actual.indent
        if expected_indent and actual_indent and expected_indent.characters is not None and actual_indent.characters is not None:
            indent_field = ("首行缩进", expected_indent.characters, actual_indent.characters)
        else:
            indent_field = ("首行缩进磅值", expected_indent.points if expected_indent else None, actual_indent.points if actual_indent else None)
        fields += [
            indent_field,
            ("段前间距", expected.before, actual.before),
            ("段后间距", expected.after, actual.after),
        ]
    mismatches, unknown = [], []
    for name, a, b in fields:
        if name in script_ranges and script_ranges[name] is None:
            continue
        display_name = "首行缩进" if name == "首行缩进磅值" else name
        if a is None or b is None:
            unknown.append(f"【{rule}】第 {index} 段{label}的{display_name}无法判定")
        elif name == "对齐方式" and (a not in ALIGNMENT_LABELS or b not in ALIGNMENT_LABELS):
            unknown.append(f"【{rule}】第 {index} 段{label}的对齐方式无法识别，请在 Word 中核对")
        elif isinstance(a, (float, int)) and not isinstance(a, bool) and name not in ("对齐方式",):
            if abs(a - b) > (0.05 if name == "首行缩进" else tolerance_pt):
                mismatches.append((name, f"{display_name}：模板 {_display(name, a)}，当前 {_display(name, b)}"))
        elif (a != b if name in ("加粗", "对齐方式") else str(a).casefold() != str(b).casefold()):
            if name in ("加粗", "对齐方式"):
                mismatches.append((name, f"模板{_display(name, a)}，当前文档{_display(name, b)}"))
            else:
                mismatches.append((name, f"{display_name}：模板 {_display(name, a)}，当前 {_display(name, b)}"))
    if not mismatches:
        return [], unknown
    findings = []
    for name, text in mismatches:
        code = codes.get(name, rule)
        title = FINE_RULES[code].title
        severity = FINE_RULES[code].severity
        want, current = _split_diff(text)
        finding = Finding(code, title, Status.FAIL, "", severity=severity,
                          expected=want, actual=current, paragraph_index=index)
        if name in script_ranges:
            finding.start, finding.end = script_ranges[name]
        elif text is not None:
            finding.start, finding.end = 0, len(text)
        else:
            finding.start, finding.end = 0, 2
        finding.message = build_comment(severity, code, title, want, current)
        findings.append(finding)
    return findings, unknown
