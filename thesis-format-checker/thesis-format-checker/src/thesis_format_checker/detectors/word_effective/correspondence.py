"""模板 ←→ 论文的“同一层级 + 同一位置”对应关系。

为什么需要它
============

论文格式检测做的是**模板比对**：论文里的第 i 个同类对象（同一样式的正文段、
同一级别的标题）只需要和模板里同一层级、同一相对位置的那个对象比较即可。

旧实现不是这样做的：它先要求模板正文“全篇选出一个占比 ≥80% 的主导值”，
再要求每一级标题“全篇完全一致”，否则整条规则无法判定。真实学校的模板常常在
一份文件里放多份样张（例如正文字号 12 磅 75% + 10.5 磅 22%），于是模板明明有
可用的对照段落，系统却对全篇输出“无法判定”。

现在的判定顺序
==============

1. **同一层级**：按段落样式名（正文/图题/表题…）或标题级别分组；
2. **同一位置**：只有当模板与论文在该层级下的对象**数量相同**时才逐个对应
   （论文第 3 个正文段 ↔ 模板第 3 个正文段）；
3. **同一层级**：数量不同时取该层级最常见的样例。模板通常只放几段样张，
   按比例硬对齐会把封面“声明”段错配到模板“摘要说明”段，因此不做比例映射；
4. 只有模板完全没有正文/标题样例时，才输出“无法判定”。

通用的纯函数（层级键、相对位置、主导值）放在 ``core/pairing.py``，
本模块只补充 Word 段落/标题对象之间的对应关系，不读 Word、不产生结论，
因此可以直接单测。
"""

from __future__ import annotations

from thesis_format_checker.core.models import BodySample, HeadingSample, HeadingSpec
from thesis_format_checker.core.pairing import dominant_value, group_by, style_key

__all__ = ["build_body_basis", "build_heading_basis", "dominant_value", "style_key"]

#: 样式组样本少于这个数时不作为基准：模板里零星几段同名字段落不代表正文，
#: 改用整篇模板正文的典型样例（否则会把封面/声明等杂散段落的行距、字号当成标准）。
MINIMUM_STYLE_SAMPLES = 3


def _mode(samples, attribute):
    """取一组样例在某个属性上的最常见值（按段落篇幅加权，并列时保留先出现的）。"""

    return dominant_value((getattr(sample, attribute), getattr(sample, "weight", 1))
                          for sample in samples)


def mode_body_sample(samples) -> BodySample:
    """把同层级的多个正文样例合成一个“典型样例”。"""

    first = samples[0]
    return BodySample(
        index=first.index,
        style_name=first.style_name,
        chinese_font=_mode(samples, "chinese_font"),
        latin_font=_mode(samples, "latin_font"),
        size_pt=_mode(samples, "size_pt"),
        first_line_indent=_mode(samples, "first_line_indent"),
        spacing=_mode(samples, "spacing"),
        weight=sum(getattr(sample, "weight", 1) for sample in samples),
    )


def mode_heading_sample(samples) -> HeadingSample:
    """把同层级的多个标题样例合成一个“典型样例”。"""

    first = samples[0]
    return HeadingSample(
        index=first.index,
        level=first.level,
        spec=HeadingSpec(
            chinese_font=dominant_value((s.spec.chinese_font, 1) for s in samples),
            latin_font=dominant_value((s.spec.latin_font, 1) for s in samples),
            size_pt=dominant_value((s.spec.size_pt, 1) for s in samples),
            bold=dominant_value((s.spec.bold, 1) for s in samples),
            alignment=dominant_value((s.spec.alignment, 1) for s in samples),
            indent=dominant_value((s.spec.indent, 1) for s in samples),
            before=dominant_value((s.spec.before, 1) for s in samples),
            after=dominant_value((s.spec.after, 1) for s in samples),
        ),
    )


def build_body_basis(template_samples, document_body) -> dict[int, BodySample]:
    """返回 ``{论文段号: 比对基准 BodySample}``。

    同一样式下数量相同 → 按“同一位置”逐个对应；数量不同 → 取该样式最常见样例
    （“同一层级”）。模板完全没有该样式时，用整篇模板正文的典型样例兜底。
    """

    samples = list(template_samples)
    body = list(document_body)
    by_style = group_by(samples, lambda item: style_key(item.style_name))
    basis: dict[int, BodySample] = {}
    for style, infos in group_by(body, lambda item: style_key(item.style_name)).items():
        group = by_style.get(style)
        if not group:
            continue
        if len(group) == len(infos):
            basis.update({info.index: group[ordinal] for ordinal, info in enumerate(infos)})
        elif len(group) < MINIMUM_STYLE_SAMPLES and len(samples) >= MINIMUM_STYLE_SAMPLES:
            typical = mode_body_sample(samples)      # 样本太少，用整篇正文的典型样例
            basis.update({info.index: typical for info in infos})
        else:
            typical = mode_body_sample(group)
            basis.update({info.index: typical for info in infos})
    missing = [info for info in body if info.index not in basis]
    if missing and samples:
        typical = mode_body_sample(samples)
        basis.update({info.index: typical for info in missing})
    return basis


def build_heading_basis(template_samples, document_headings) -> dict[int, HeadingSample]:
    """返回 ``{论文标题段号: 比对基准 HeadingSample}``（同一级别 + 同位置优先）。

    Args:
        template_samples: 模板标题样例（``core.models.HeadingSample`` 列表）。
        document_headings: 论文标题 ``[(段号, 级别, 段落), …]``，来自 ``collect_headings``。
    """

    samples = list(template_samples)
    by_level = group_by(samples, lambda item: item.level)
    by_document = group_by(list(document_headings), lambda item: item[1])
    basis: dict[int, HeadingSample] = {}
    for level, items in by_document.items():
        group = by_level.get(level)
        if not group:
            continue
        if len(group) == len(items):
            basis.update({index: group[ordinal] for ordinal, (index, _, _) in enumerate(items)})
        else:
            typical = mode_heading_sample(group)
            basis.update({index: typical for index, _, _ in items})
    return basis
