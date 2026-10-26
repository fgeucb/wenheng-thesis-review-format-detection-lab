from __future__ import annotations

import math
import unicodedata
from collections.abc import Iterator
from dataclasses import dataclass

WD_UNDEFINED = 9999999


def clean_word_text(text: str) -> str:
    return text.rstrip("\r\x07")


def script_of(char: str) -> str | None:
    code = ord(char)
    if 0x3400 <= code <= 0x4DBF or 0x4E00 <= code <= 0x9FFF or 0xF900 <= code <= 0xFAFF:
        return "chinese"
    if char.isascii() and (char.isalpha() or char.isdigit()):
        return "latin"
    return None


def iter_script_spans(text: str) -> Iterator[tuple[int, int, str]]:
    start: int | None = None
    current: str | None = None
    for index, char in enumerate(text):
        kind = script_of(char)
        if kind is None:
            if start is not None:
                yield start, index, current  # type: ignore[arg-type]
                start = None
                current = None
            continue
        if current != kind:
            if start is not None:
                yield start, index, current  # type: ignore[arg-type]
            start = index
            current = kind
    if start is not None:
        yield start, len(text), current  # type: ignore[arg-type]


def is_word_undefined(value: object) -> bool:
    try:
        return int(value) == WD_UNDEFINED
    except (TypeError, ValueError, OverflowError):
        return False


def normalized_font(value: object) -> str | None:
    if value is None or is_word_undefined(value):
        return None
    text = str(value).strip()
    return text or None


def _theme_font_group(raw_name: str, theme_font_value: object = None) -> str | None:
    """识别 Word 的标题/正文字体主题占位名称。"""
    if not raw_name.startswith("+"):
        return None
    folded = raw_name.casefold().replace(" ", "")
    if any(marker in folded for marker in ("标题", "heading", "headings", "+mj")):
        return "major"
    if any(marker in folded for marker in ("正文", "body", "+mn")):
        return "minor"
    try:
        value = int(theme_font_value)
    except (TypeError, ValueError, OverflowError):
        return None
    # WdThemeFont：MajorAscii/HAnsi/Bidi/EastAsia 为 1..4，Minor 为 5..8。
    if value in (1, 2, 3, 4):
        return "major"
    if value in (5, 6, 7, 8):
        return "minor"
    return None


def resolve_effective_font(document, font, script: str, raw_value: object) -> str | None:
    """将 Word 主题字体占位名称解析为文档实际使用的字体名。"""
    raw_name = normalized_font(raw_value)
    if raw_name is None:
        return None
    try:
        theme_font_value = font.ThemeFont
    except Exception:
        theme_font_value = None
    group = _theme_font_group(raw_name, theme_font_value)
    if group is None:
        # 未能识别的 +占位名称不能当作真实字体比较，以免产生误报。
        return raw_name if not raw_name.startswith("+") else None
    try:
        scheme = document.DocumentTheme.ThemeFontScheme
        collection = scheme.MajorFont if group == "major" else scheme.MinorFont
        # MsoFontLanguageIndex：Latin = 1，EastAsian = 3。
        language_index = 3 if script == "chinese" else 1
        try:
            theme_font = collection.Item(language_index)
        except Exception:
            theme_font = collection(language_index)
        return normalized_font(theme_font.Name)
    except Exception:
        # 主题不能可靠解析时交给上层报告“无法判定”，不显示内部占位名称。
        return None


def normalized_number(value: object) -> float | None:
    if value is None or is_word_undefined(value):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not math.isfinite(number):
        return None
    return number


@dataclass(frozen=True)
class SpanFormat:
    start: int
    end: int
    script: str
    font_name: str | None
    size_pt: float | None


def read_span_formats(document, paragraph) -> list[SpanFormat]:
    """让 Word 解析主题字体和样式继承，返回正文脚本片段的有效格式。

    性能：字体只会在 **格式一致的小段** 边界变化，因此当整段读取返回“未定义”
    （混合直接格式）时，先按可用信息把区间细化，而不是逐字符读：

    * python-docx 段落（有 ``w:p`` 元素）：按 **run 边界** 精确切开，每个 run 一次调用；
    * Word COM 段落（旧版有效格式流程传入的对象，没有 ``w:p``）：按 **二分收敛**
      对半读，混合的一半继续对半，直到读得出格式或缩到单字符。

    两种方式都保留“可定位到字符”的能力，同时把 COM 调用从“每字符一次”降到
    “每个格式段一次 + 少量探测”。注意：这里必须兼容 COM 段落——旧版 Word 流程
    传进来的是 ``win32com`` 对象，直接访问 ``paragraph._p`` 会抛 ``AttributeError``，
    一旦抛出就会让整条 Word 检测链路退化成“无法判定”。
    """
    text = clean_word_text(paragraph.Range.Text)
    base = paragraph.Range.Start
    spans: list[SpanFormat] = []

    def read(start: int, end: int, script: str):
        rng = document.Range(base + start, base + end)
        font = rng.Font
        name_value = font.NameFarEast if script == "chinese" else font.NameAscii
        name = resolve_effective_font(document, font, script, name_value)
        return name, normalized_number(font.Size)

    def bisect(start: int, end: int, script: str, depth: int = 0) -> None:
        """混合区间对半收敛：读得出格式就收下，读不出继续对半，最细到单字符。"""

        if end - start <= 1 or depth >= 8:
            piece_name, piece_size = read(start, end, script)
            spans.append(SpanFormat(start, end, script, piece_name, piece_size))
            return
        middle = (start + end) // 2
        for left, right in ((start, middle), (middle, end)):
            piece_name, piece_size = read(left, right, script)
            if piece_name is not None and piece_size is not None:
                spans.append(SpanFormat(left, right, script, piece_name, piece_size))
            else:
                bisect(left, right, script, depth + 1)

    for start, end, script in iter_script_spans(text):
        name, size = read(start, end, script)
        if name is not None and size is not None:
            spans.append(SpanFormat(start, end, script, name, size))
            continue
        # 混合直接格式会令 COM 返回 wdUndefined：先按 run 边界细化。
        boundaries = _run_boundaries(paragraph, text, start, end)
        if boundaries is None:
            bisect(start, end, script)          # Word COM 段落：二分收敛
            continue
        for run_start, run_end in boundaries:
            run_name, run_size = read(run_start, run_end, script)
            if run_name is not None and run_size is not None:
                spans.append(SpanFormat(run_start, run_end, script, run_name, run_size))
                continue
            # 单个 run 内部仍混合（少见）：只对这一段继续二分。
            bisect(run_start, run_end, script)
    return merge_span_formats(spans)


def _run_boundaries(paragraph, text: str, start: int, end: int) -> list[tuple[int, int]] | None:
    """把 ``[start, end)`` 按段落里 run 的文字边界切开（只算文字类节点）。

    Returns:
        ``[(start, end), ...]``；若 ``paragraph`` 不是 python-docx 段落（例如旧版
        Word 流程传入的 COM 段落），返回 ``None``，由调用方改用二分收敛。
        XML 文字长度与 COM 文本长度不一致（含域/对象）时退回“整段一次读”。
    """

    from thesis_format_checker.core.xmlutil import qn

    element = getattr(paragraph, "_p", None)
    if element is None:
        return None                      # COM 段落：没有 OOXML 可读
    boundaries = [0]
    length = 0
    for run in element.findall(qn("w:r")):
        for child in run:
            if child.tag == qn("w:t"):
                length += len(child.text or "")
            elif child.tag in (qn("w:tab"), qn("w:br"), qn("w:cr")):
                length += 1
        boundaries.append(length)
    if length != len(text):
        return [(start, end)]        # 域/对象等导致长度不一致：整段退回一次读
    pieces = []
    for left, right in zip(boundaries, boundaries[1:]):
        lower, upper = max(left, start), min(right, end)
        if lower < upper:
            pieces.append((lower, upper))
    return pieces or [(start, end)]


def merge_span_formats(spans: list[SpanFormat]) -> list[SpanFormat]:
    merged: list[SpanFormat] = []
    for item in spans:
        if (
            merged
            and merged[-1].end == item.start
            and merged[-1].script == item.script
            and merged[-1].font_name == item.font_name
            and merged[-1].size_pt == item.size_pt
        ):
            previous = merged[-1]
            merged[-1] = SpanFormat(
                previous.start, item.end, item.script, item.font_name, item.size_pt
            )
        else:
            merged.append(item)
    return merged
