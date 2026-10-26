"""全文与文末卫生检查（154/160/165）的取证实现。

* 154.1/154.2/154.3：残留批注、残留修订、模板提示语/占位文字；
* 160.1/160.2：致谢部分的字体字号与行距（以模板致谢样例为准）；
* 165.2/165.3：文末隐藏对象、文末残留模板示例。

不涉及分页渲染的结论（“文末空白页”165.1、“非必要空白页”032.1 等）不在本模块。
"""

from __future__ import annotations

import re
from collections import Counter

from thesis_format_checker.core.xmlutil import qn

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import (
    DocxContext,
    effective_paragraph,
    make_finding,
    unknown,
)
from thesis_format_checker.detectors.headings import _heading_level
from thesis_format_checker.detectors.text_runs import effective_text_runs

RULES = ("154", "160", "165")
TITLES = {
    "154": "文档残留批注、修订或模板提示语",
    "160": "致谢标题或正文格式不符合要求",
    "165": "文档末尾存在空白页、隐藏对象或模板示例",
}

#: 模板提示语/占位文字：XXX、××、待填、请填写、“（三号黑体）”这类说明。
_PLACEHOLDER = re.compile(
    r"(?:X{3,}|×{2,}|[xX]{4,}|待填(?:写)?|请(?:自行)?填写|填写示例|"
    r"（\s*[一二三四五六七八九十0-9]+\s*号\s*(?:黑体|宋体|楷体|仿宋)[^）]{0,20}）)")
_ACKNOWLEDGEMENT = re.compile(r"^(?:致\s*谢|acknowledg(?:e)?ments?)$", re.I)
_TAIL_PARAGRAPHS = 20


def _title(code: str) -> str:
    item = FINE_RULES.get(code)
    return item.title if item else TITLES[code.split(".")[0]]


def _unavailable(rule, reason, paragraph=None):
    return unknown(rule, _title(rule), reason, paragraph)


def _section(ctx: DocxContext, pattern) -> list:
    """取某个组成部分（如“致谢”）标题之后、下一个标题之前的段落。"""

    active = False
    result = []
    for paragraph in ctx.paragraphs:
        text = paragraph.text.strip()
        level = _heading_level(paragraph)
        if not active and level and pattern.match(text):
            active = True
            result.append(paragraph)
            continue
        if not active:
            continue
        if level and text:
            break
        result.append(paragraph)
    return result


def _profile(paragraphs) -> dict:
    """一组段落的典型格式：中文字体、字号、行距。"""

    fonts, sizes, lines = Counter(), Counter(), Counter()
    for position, paragraph in enumerate(paragraphs):
        for text, props in effective_text_runs(paragraph):
            if not text.strip():
                continue
            weight = 1 if position == 0 else len(text)
            if props.get("east_asia_font"):
                fonts[props["east_asia_font"]] += weight
            if props.get("size_pt"):
                sizes[round(props["size_pt"], 2)] += weight
        fmt = effective_paragraph(paragraph)
        if fmt.get("line_spacing"):
            lines[(round(fmt["line_spacing"], 2), fmt.get("line_spacing_rule"))] += max(len(paragraph.text), 1)
    return {"font": fonts.most_common(1)[0][0] if fonts else None,
            "size": sizes.most_common(1)[0][0] if sizes else None,
            "line": lines.most_common(1)[0][0] if lines else None}


def extract_hygiene_baseline(ctx: DocxContext) -> dict:
    """模板侧的致谢样例格式；其余规则不需要模板基准。"""

    paragraphs = _section(ctx, _ACKNOWLEDGEMENT)
    if not paragraphs:
        return {"hygiene": {"acknowledgements": None}}
    title, body = paragraphs[:1], paragraphs[1:]
    return {"hygiene": {"acknowledgements": {"title": _profile(title), "body": _profile(body or title)}}}


def _comments(ctx) -> list:
    try:
        return [str(part.partname) for part in ctx.document.part.package.parts
                if str(part.partname).endswith("comments.xml")]
    except Exception:
        return []


def _revised_paragraphs(ctx: DocxContext) -> list:
    tags = (qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"))
    result = []
    for paragraph in ctx.paragraphs:
        if paragraph._p.xpath(".//w:pPrChange|.//w:rPrChange") or any(
            node.tag in tags for node in paragraph._p.iter()
        ):
            result.append(paragraph)
    return result


def _hidden_objects(paragraph) -> bool:
    if paragraph._p.xpath(".//w:drawing|.//w:pict|.//w:object|.//w:txbxContent"):
        return True
    return bool(paragraph._p.xpath(".//w:rPr/w:vanish | .//w:pPr/w:rPr/w:vanish"))


def _placeholder_paragraphs(ctx: DocxContext) -> list:
    return [paragraph for paragraph in ctx.paragraphs
            if paragraph.text.strip() and _PLACEHOLDER.search(paragraph.text)]


def _check_leftovers(ctx: DocxContext) -> list:
    findings = []
    first = next(iter(ctx.paragraphs), None)
    pages = _comments(ctx)
    if pages:
        findings.append(make_finding("154.1", _title("154.1"), "提交前应删除所有批注",
                                     "文档包内仍存在批注部件（comments.xml）", first))
    revised = _revised_paragraphs(ctx)
    if revised:
        findings.append(make_finding("154.2", _title("154.2"), "提交前应接受全部修订",
                                     f"仍存在未接受的修订（首个位置：第 {ctx.index(revised[0]) + 1} 段）",
                                     revised[0]))
    placeholders = _placeholder_paragraphs(ctx)
    if placeholders:
        paragraph = placeholders[0]
        findings.append(make_finding("154.3", _title("154.3"), "不应保留模板提示语或占位文字",
                                     f"仍存在占位文字：{paragraph.text.strip()[:30]}", paragraph))
    return findings


def _check_acknowledgements(baseline: dict, ctx: DocxContext) -> list:
    expected = (baseline.get("hygiene") or {}).get("acknowledgements")
    paragraphs = _section(ctx, _ACKNOWLEDGEMENT)
    findings = []
    if not expected or not paragraphs:
        return findings
    groups = (("title", paragraphs[:1], "160.1", "标题"), ("body", paragraphs[1:] or paragraphs[:1], "160.1", "正文"))
    for kind, items, code, label in groups:
        actual = _profile(items)
        expected_profile = expected.get(kind) or {}
        if actual["font"] and expected_profile.get("font") and actual["font"] != expected_profile["font"]:
            findings.append(make_finding(code, _title(code),
                                         f"致谢{label}中文应为{expected_profile['font']}",
                                         f"当前为{actual['font']}", items[0]))
        if actual["size"] and expected_profile.get("size") and abs(actual["size"] - expected_profile["size"]) > 0.05:
            findings.append(make_finding(code, _title(code),
                                         f"致谢{label}字号应为{expected_profile['size']:g} 磅",
                                         f"当前为{actual['size']:g} 磅", items[0]))
        if actual["line"] and expected_profile.get("line") and actual["line"] != expected_profile["line"]:
            findings.append(make_finding("160.2", _title("160.2"),
                                         f"致谢{label}行距应保持模板样例（{expected_profile['line'][0]:g} 磅）",
                                         f"当前行距为 {actual['line'][0]:g} 磅", items[0]))
    return findings


def _check_tail(ctx: DocxContext) -> list:
    findings = []
    tail = ctx.paragraphs[-_TAIL_PARAGRAPHS:]
    hidden = [paragraph for paragraph in tail if _hidden_objects(paragraph)]
    if hidden:
        findings.append(make_finding("165.2", _title("165.2"), "文末不应存在隐藏对象",
                                     f"文末仍有隐藏图形/隐藏文字（第 {ctx.index(hidden[0]) + 1} 段）",
                                     hidden[0]))
    placeholders = [paragraph for paragraph in tail if paragraph.text.strip() and _PLACEHOLDER.search(paragraph.text)]
    if placeholders:
        paragraph = placeholders[0]
        findings.append(make_finding("165.3", _title("165.3"), "文末不应残留模板示例",
                                     f"文末仍有模板示例文字：{paragraph.text.strip()[:30]}", paragraph))
    return findings


def check_hygiene_errors(baseline: dict, ctx: DocxContext) -> list:
    findings = []
    for rule, check in (("154", lambda: _check_leftovers(ctx)),
                        ("165", lambda: _check_tail(ctx)),
                        ("160", lambda: _check_acknowledgements(baseline, ctx))):
        try:
            findings.extend(check())
        except Exception as exc:
            findings.append(_unavailable(rule, f"读取失败（{type(exc).__name__}）",
                                         next(iter(ctx.paragraphs), None)))
    return findings
