from __future__ import annotations

from collections import defaultdict
from typing import Callable, Hashable, Iterable, TypeVar

from thesis_format_checker.core.effective_format import normalized_number, read_span_formats
from thesis_format_checker.core.models import (
    BodySample,
    HeadingSample,
    IndentSpec,
    MarginSpec,
    SpacingSpec,
    TemplateProfile,
)
from thesis_format_checker.detectors.word_effective.correspondence import dominant_value
from thesis_format_checker.detectors.word_effective.headers import read_headers
from thesis_format_checker.detectors.word_effective.headings import collect_headings, read_heading
from thesis_format_checker.detectors.word_effective.paragraph_spacing import describe_line_spacing, read_spacing
from thesis_format_checker.core.scope import ParagraphInfo

T = TypeVar("T", bound=Hashable)
POINTS_PER_MM = 72.0 / 25.4


def points_to_mm(value: object) -> float:
    number = normalized_number(value)
    if number is None:
        raise ValueError("Word 返回了无法解析的页面尺寸")
    return round(number / POINTS_PER_MM, 3)


def read_margin(section) -> MarginSpec:
    setup = section.PageSetup
    return MarginSpec(
        top=points_to_mm(setup.TopMargin),
        bottom=points_to_mm(setup.BottomMargin),
        left=points_to_mm(setup.LeftMargin),
        right=points_to_mm(setup.RightMargin),
    )


def resolve_dominant(
    weighted_values: Iterable[tuple[T, int]], *, threshold: float = 0.8, minimum_weight: int = 2,
    format_value: Callable[[T], str] = str,
) -> tuple[T | None, str | None]:
    totals: dict[T, int] = defaultdict(int)
    total = 0
    for value, weight in weighted_values:
        if weight <= 0:
            continue
        totals[value] += weight
        total += weight
    if total < minimum_weight or not totals:
        return None, "样本不足"
    ordered = sorted(totals.items(), key=lambda pair: pair[1], reverse=True)
    value, weight = ordered[0]
    share = weight / total
    if len(ordered) > 1 and share < threshold:
        details = "、".join(f"{format_value(key)}（{count / total:.0%}）" for key, count in ordered[:4])
        return None, f"模板正文格式存在冲突：{details}"
    return value, None


def _indent_of(paragraph) -> IndentSpec | None:
    fmt = paragraph.Format
    points = normalized_number(fmt.FirstLineIndent)
    if points is None:
        return None
    chars = None
    try:
        chars = normalized_number(fmt.CharacterUnitFirstLineIndent)
    except Exception:
        pass
    return IndentSpec(round(points, 3), round(chars, 3) if chars is not None else None)


def _indent_key(spec: IndentSpec) -> tuple[float, float | None]:
    return round(spec.points, 1), round(spec.characters, 1) if spec.characters is not None else None


def _describe_indent(key: tuple[float, float | None]) -> str:
    points, characters = key
    if characters is None:
        return f"{points:g} 磅（字符单位无法判定）"
    return f"{characters:g} 字符（约 {points:g} 磅）"


def _describe_spacing(key: tuple[int, float, float, float]) -> str:
    rule, line, before, after = key
    label = describe_line_spacing(SpacingSpec(rule, line, before, after), show_calculated_points=True)
    return f"{label or '行距类型无法识别'}；段前 {before:g} 磅、段后 {after:g} 磅"


def build_template_profile(document, paragraphs: list[ParagraphInfo], heading_styles: dict[str, int] | None = None) -> TemplateProfile:
    profile = TemplateProfile()
    profile.body_style_names = sorted({p.style_name for p in paragraphs if p.style_name})
    profile.margins = [read_margin(document.Sections(index)) for index in range(1, document.Sections.Count + 1)]
    profile.headers = [read_headers(document.Sections(index)) for index in range(1, document.Sections.Count + 1)]
    if profile.margins and all(item == profile.margins[0] for item in profile.margins[1:]):
        profile.uniform_margin = profile.margins[0]
    elif len(profile.margins) == 1:
        profile.uniform_margin = profile.margins[0]
    else:
        profile.unknown_reasons.append(
            "模板包含多种节页面边距；需要在映射文件 section_map 中显式指定对应关系"
        )

    chinese: list[tuple[str, int]] = []
    latin: list[tuple[str, int]] = []
    sizes: list[tuple[float, int]] = []
    indents: list[tuple[tuple[float, float | None], int]] = []
    spacings = []
    spacing_lookup = {}
    indent_lookup: dict[tuple[float, float | None], IndentSpec] = {}
    for info in paragraphs:
        # 同一段落只让 Word 解析一次：既做全篇兜底统计，也留下逐段基准（同层级+同位置比对用）。
        spans = read_span_formats(document, info.paragraph)
        for span in spans:
            weight = span.end - span.start
            if span.font_name:
                (chinese if span.script == "chinese" else latin).append((span.font_name, weight))
            if span.size_pt is not None:
                sizes.append((round(span.size_pt, 2), weight))
        indent = _indent_of(info.paragraph)
        if indent:
            key = _indent_key(indent)
            indent_lookup[key] = indent
            indents.append((key, max(len(info.text), 1)))
        spacing = read_spacing(info.paragraph)
        if spacing:
            key = (spacing.line_rule, round(spacing.line_points, 1), round(spacing.before, 1), round(spacing.after, 1))
            spacing_lookup[key] = spacing
            spacings.append((key, max(len(info.text), 1)))
        profile.body_samples.append(BodySample(
            index=info.index,
            style_name=info.style_name,
            chinese_font=dominant_value((span.font_name, span.end - span.start)
                                        for span in spans if span.script == "chinese"),
            latin_font=dominant_value((span.font_name, span.end - span.start)
                                      for span in spans if span.script != "chinese"),
            size_pt=dominant_value((round(span.size_pt, 2), span.end - span.start)
                                   for span in spans if span.size_pt is not None),
            first_line_indent=indent,
            spacing=spacing,
            weight=max(len(info.text), 1),
        ))

    profile.chinese_font, reason = resolve_dominant(chinese)
    if reason:
        profile.unknown_reasons.append(f"模板中文字体存在多种样例，已按同层级同位置逐段比对：{reason}")
    profile.latin_font, reason = resolve_dominant(latin)
    if reason:
        profile.unknown_reasons.append(f"模板英文/数字字体存在多种样例，已按同层级同位置逐段比对：{reason}")
    profile.font_size_pt, reason = resolve_dominant(sizes, format_value=lambda value: f"{value:g} 磅")
    if reason:
        profile.unknown_reasons.append(f"模板正文字号存在多种样例，已按同层级同位置逐段比对：{reason}")
    indent_key, reason = resolve_dominant(indents, format_value=_describe_indent)
    if indent_key is not None:
        profile.first_line_indent = indent_lookup[indent_key]
    if reason:
        profile.unknown_reasons.append(f"模板首行缩进存在多种样例，已按同层级同位置逐段比对：{reason}")
    spacing_key, reason = resolve_dominant(spacings, format_value=_describe_spacing)
    if spacing_key is not None:
        profile.body_spacing = spacing_lookup[spacing_key]
    if reason:
        profile.unknown_reasons.append(f"模板正文行距/段间距存在多种样例，已按同层级同位置逐段比对：{reason}")

    # 标题：逐个记录模板样例（同级标题按出现顺序与论文对应），不再要求全模板唯一。
    for index, level, paragraph in collect_headings(document, heading_styles):
        profile.heading_samples.append(HeadingSample(index, level, read_heading(document, paragraph)))
    by_level: dict[int, list[HeadingSpec]] = defaultdict(list)
    for sample in profile.heading_samples:
        by_level[sample.level].append(sample.spec)
    for level, samples in by_level.items():
        # 完全一致时留一份兜底基准；不一致也照常比对，逐条按同层级同位置取样例。
        if len(set(samples)) == 1:
            profile.headings[level] = samples[0]
    return profile
