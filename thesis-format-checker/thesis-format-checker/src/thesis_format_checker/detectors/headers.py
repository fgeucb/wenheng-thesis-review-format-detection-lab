"""Template-driven, read-only header/footer additions for rules 034--044.

035 stays in the original header detector. PAGE caches are never treated as
current pagination results. All findings anchor to the corresponding body.
"""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
import re

from docx.oxml import OxmlElement
from thesis_format_checker.core.xmlutil import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, describe_value, effective_paragraph, effective_runs, make_finding, unknown
from thesis_format_checker.detectors.pages import _candidate_sections, _roles, _section_anchor, _section_paragraphs
from thesis_format_checker.detectors.fields import scan_fields, scan_body_fields


_TITLES = {
    "034": "页眉显示配置与模板不一致",
    "036": "页眉格式或下边框与模板不一致",
    "037": "奇偶页页眉内容或位置不符合模板",
    "038": "页眉链接到前一节的关系需核对",
    "039": "页脚文字或线条与模板不一致",
    "040": "页码缺失、重复或连续性需核对",
    "041": "页码数字形式与模板不一致",
    "042": "页码起始或续编号设置与模板不一致",
    "043": "页码位置或文字格式与模板不一致",
    "044": "模板禁止页码的前置区域设置了页码",
}
_VARIANTS = {"default": "普通/奇数页", "first": "首页", "even": "偶数页"}
_FRONT = {"封面", "声明", "中文摘要", "英文摘要", "目录"}
_FONT_LABELS = {"east_asia_font": "中文字体", "ascii_font": "西文字体", "size_pt": "字号"}


class _StoryFindings(list):
    """在产生证据时保存当前变体/承载区域，包括提前 continue 的分支。"""

    variant = None
    story = ""

    def append(self, finding):
        finding.page_variant = self.variant
        finding.page_story = self.story
        if self.variant is not None:
            finding.page_scope = "variant"
            finding.header_kind = {"default": 1, "first": 2, "even": 3}[self.variant]
        super().append(finding)

    def extend(self, findings):
        for finding in findings:
            self.append(finding)


def _limited(rule, reason, anchor):
    return unknown(rule, _title(rule), UNREADABLE + "；" + reason, anchor)


def _fail(rule, expected, actual, anchor):
    return make_finding(rule, _title(rule), expected, actual, anchor)


def _even_pages_text(enabled) -> str:
    """奇偶页设置的中文说明（不把 Word 的 True/False 直接写进批注）。"""

    return ("偶数页页眉页脚与奇数页分开设置（首页仍按首页设置）"
            if enabled else "偶数页页眉页脚与奇数页相同（未单独设置偶数页）")


#: 页眉格式属性 → 细粒度规则编号
_HEADER_FORMAT_CODES = {
    "中文east_asia_font": "036.1",
    "西文ascii_font": "036.2",
    "中文size_pt": "036.3",
    "西文size_pt": "036.3",
    "alignment": "036.4",
    "border": "036.5",
    "left_indent": "036.6",
    "right_indent": "036.6",
    "first_line_indent": "036.6",
}
#: 页码格式属性 → 细粒度规则编号
_PAGE_FORMAT_CODES = {
    "中文east_asia_font": "043.2",
    "西文ascii_font": "043.2",
    "中文size_pt": "043.3",
    "西文size_pt": "043.3",
    "alignment": "043.4",
    "left_indent": "043.1",
    "right_indent": "043.1",
    "first_line_indent": "043.1",
    "carrier": "043.1",
    "distance": "043.1",
}


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else _TITLES[code]


def _split_diff(text: str) -> tuple[str, str]:
    """把“模板 X，当前 Y”形式的差异拆成（应当, 实际）；无法解析时原样保留。"""

    match = re.search(r"模板\s*(.*?)[，,]\s*当前\s*(.*)$", text)
    if match:
        return match.group(1).strip(), match.group(2).strip()
    return "按模板要求", text


def _emit_changes(rule: str, label: str, changes: list, anchor, codes: dict) -> list:
    """把“属性 → 差异文本”按细粒度规则分组：一个属性一条结论。"""

    grouped: dict[str, list[str]] = {}
    for key, text in changes:
        grouped.setdefault(codes.get(key, rule), []).append(text)
    findings = []
    for code, texts in grouped.items():
        want, actual = _split_diff("；".join(texts))
        findings.append(_fail(code, label + "：" + want, actual, anchor))
    return findings


def _flag(root, name):
    node = root.find(qn("w:" + name))
    if node is None:
        return False
    value = node.get(qn("w:val"), "1")
    if value not in {"0", "1", "true", "false", "on", "off"}:
        raise ValueError(name + " 的开关值无效")
    return value not in {"0", "false", "off"}


def _active(section):
    result = ["default"]
    if section["first"]:
        result.append("first")
    if section["even"]:
        result.append("even")
    return result


def _format_runs(ctx, p, elements=None):
    """Resolve cached-result and fldSimple run formats without editing source."""
    allowed = None if elements is None else set(elements)
    result = defaultdict(set)
    for run in p.iter(qn("w:r")):
        if any(a.tag == qn("w:txbxContent") for a in run.iterancestors()):
            continue
        texts = [n for n in run.iter(qn("w:t")) if allowed is None or n in allowed]
        text = "".join(n.text or "" for n in texts)
        if not text.strip():
            continue
        shell = OxmlElement("w:p")
        props = p.find(qn("w:pPr"))
        if props is not None:
            shell.append(deepcopy(props))
        copy_run = deepcopy(run)
        for node in list(copy_run):
            if node.tag != qn("w:rPr"):
                copy_run.remove(node)
        leaf = OxmlElement("w:t")
        leaf.text = text
        copy_run.append(leaf)
        shell.append(copy_run)
        values = effective_runs(Paragraph(shell, ctx.document._body))
        for _, fmt in values:
            if re.search(r"[\u3400-\u9fff]", text):
                for key in ("east_asia_font", "size_pt"):
                    result["中文" + key].add(fmt[key])
            if re.search(r"[A-Za-z0-9\u2160-\u217f]", text):
                for key in ("ascii_font", "size_pt"):
                    result["西文" + key].add(fmt[key])
    return {k: tuple(sorted(v, key=str)) for k, v in result.items()}


def _paragraph_borders(paragraph):
    props = [paragraph._p.pPr]
    style, seen = paragraph.style, set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        props.append(style.element.pPr)
        style = style.base_style
    result = {}
    for side in ("top", "bottom", "left", "right", "between", "bar"):
        for prop in props:
            if prop is None:
                continue
            borders = prop.find(qn("w:pBdr"))
            if borders is None:
                continue
            border = borders.find(qn("w:" + side))
            if border is not None:
                if border.get(qn("w:val"), "nil") not in {"nil", "none"}:
                    result[side] = tuple(sorted((key.rsplit("}", 1)[-1], value) for key, value in border.attrib.items()))
                break
    return result


def _bottom_border(paragraph):
    return _paragraph_borders(paragraph).get("bottom", ())


def _position(paragraph):
    fmt = effective_paragraph(paragraph)
    return {key: fmt[key] for key in ("alignment", "left_indent", "right_indent", "first_line_indent")}


def _empty_story(linked=False):
    return {"present": False, "linked": linked, "source": None, "static": "", "fields": [],
            "pages": [], "paragraphs": [], "objects": False, "complex": False, "errors": [],
            "borders": (), "dynamic": False, "page_uncertain": False, "opaque": False}


def _has_hidden_content(ctx, root):
    for node in root.iter(qn("w:r")):
        p = next((a for a in node.iterancestors() if a.tag == qn("w:p")), None)
        if p is None:
            continue
        paragraph = Paragraph(p, ctx.document._body)
        run = Run(node, paragraph)
        direct = node.find(qn("w:rPr"))
        direct_hidden = direct.find(qn("w:vanish")) if direct is not None else None
        if direct_hidden is not None:
            if _flag(direct, "vanish"):
                return True
            continue
        # Hidden is also a style toggle. When several styles toggle it, its
        # eventual rendering can depend on the chain. Conservatively report
        # the visibility limitation rather than treat the stored run as shown.
        for style in (run.style, paragraph.style):
            seen = set()
            while style is not None and style.style_id not in seen:
                seen.add(style.style_id)
                props = style.element.rPr
                if props is not None and _flag(props, "vanish"):
                    return True
                style = style.base_style
        defaults = ctx.document.styles.element.find(".//" + qn("w:rPrDefault"))
        if defaults is not None:
            props = defaults.find(qn("w:rPr"))
            if props is not None and _flag(props, "vanish"):
                return True
    return False


def _read_story(ctx, sections, index, kind, variant):
    direct = [r for r in sections[index]._sectPr.findall(qn("w:" + kind + "Reference"))
              if r.get(qn("w:type")) == variant]
    if len(direct) > 1:
        raise ValueError("同一种页眉页脚类型存在重复引用")
    linked = index > 0 and not direct
    selected, source = None, None
    for previous in range(index, -1, -1):
        refs = [r for r in sections[previous]._sectPr.findall(qn("w:" + kind + "Reference"))
                if r.get(qn("w:type")) == variant]
        if len(refs) > 1:
            raise ValueError("继承的页眉页脚类型存在重复引用")
        if refs:
            selected, source = refs[0], previous
            break
    result = _empty_story(linked)
    result["source"] = source
    if selected is None:
        return result
    rid = selected.get(qn("r:id"))
    relation = ctx.document.part.rels[rid]
    if relation.is_external or not relation.reltype.endswith("/" + kind):
        raise ValueError("页眉页脚关系类型无效")
    root = relation.target_part.element
    fields = scan_fields(root)
    result["fields"] = fields
    result["errors"] = list(getattr(fields, "errors", []))
    if any(f.get("malformed") for f in fields):
        result["errors"].append("存在未闭合或损坏的域")
    excluded = {n for field in fields for n in field.get("result_elements", [])}
    result["static"] = "\n".join(
        "".join(n.text or "" for n in p.iter(qn("w:t")) if n not in excluded)
        for p in root.iter(qn("w:p")))
    result["objects"] = any(n.tag in {qn("w:drawing"), qn("w:pict"), qn("w:object")}
                            for n in root.iter())
    result["opaque"] = any(n.tag in {qn("w:txbxContent"), qn("w:object"), qn("w:altChunk"),
                                     qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"),
                                     qn("w:rPrChange"), qn("w:pPrChange")}
                           for n in root.iter()) or _has_hidden_content(ctx, root)
    result["complex"] = result["objects"] or bool(root.findall(".//" + qn("w:tbl")))
    borders = []
    for p in root.iter(qn("w:p")):
        paragraph = Paragraph(p, ctx.document._body)
        all_borders = _paragraph_borders(paragraph)
        border = all_borders.get("bottom", ())
        borders.append(tuple(sorted(all_borders.items())))
        if paragraph.text.strip() or border or any(f.get("paragraph_element") is p for f in fields):
            result["paragraphs"].append({"position": _position(paragraph), "border": border,
                                          "font": _format_runs(ctx, p)})
    result["borders"] = tuple(border for border in borders if border)
    result["dynamic"] = any(f.get("kind", "").upper() not in {"PAGE", "NUMPAGES", "SECTIONPAGES"}
                            for f in fields)
    result["page_uncertain"] = bool(result["errors"] or result["opaque"] or
                                      any(f.get("nested") for f in fields) or
                                      any(_flag(r, "vanish") for r in root.iter(qn("w:rPr"))))
    for field in fields:
        if field.get("kind", "").upper() != "PAGE":
            continue
        p = field.get("paragraph_element")
        if p is None:
            result["page_uncertain"] = True
            continue
        paragraph = Paragraph(p, ctx.document._body)
        result["pages"].append({"instruction": field["instruction"],
                                "position": _position(paragraph),
                                "font": _format_runs(ctx, p, field.get("result_elements", [])),
                                "in_table": any(a.tag == qn("w:tc") for a in p.iterancestors()),
                                "nested": bool(field.get("nested"))})
    result["present"] = bool(result["static"].strip() or fields or result["objects"] or result["borders"])
    return result


def _read_sections(ctx):
    sections = list(ctx.document.sections)
    groups = _section_paragraphs(ctx)
    snapshots = []
    body_pages = defaultdict(list)
    for field in scan_body_fields(ctx):
        if field.get("kind") != "PAGE":
            continue
        p = field.get("paragraph_element")
        if p is not None:
            wrapper = Paragraph(p, ctx.document._body)
            body_pages[ctx.section_index(wrapper)].append(field)
    for index, section in enumerate(sections):
        paragraphs = groups[index]
        record = {"index": index, "roles": _roles(paragraphs), "stories": {}, "errors": [],
                  "body_pages": body_pages[index],
                  "anchor": _section_anchor(paragraphs, next(iter(ctx.paragraphs), None))}
        try:
            record["first"] = _flag(section._sectPr, "titlePg")
            record["even"] = _flag(ctx.document.settings.element, "evenAndOddHeaders")
            node = section._sectPr.find(qn("w:pgNumType"))
            # OOXML omission semantics, not a school-specific format requirement.
            record["number_format"] = node.get(qn("w:fmt"), "decimal") if node is not None else "decimal"
            start = node.get(qn("w:start")) if node is not None else None
            record["start"] = ("start", int(start)) if start is not None else (("start", 1) if index == 0 else ("continue", None))
            record["chapter"] = tuple(node.get(qn("w:" + attr)) for attr in ("chapStyle", "chapSep")) if node is not None else (None, None)
            record["header_distance"] = int(section.header_distance) if section.header_distance is not None else None
            record["footer_distance"] = int(section.footer_distance) if section.footer_distance is not None else None
        except Exception as exc:
            record["errors"].append(str(exc))
        for kind in ("header", "footer"):
            for variant in _VARIANTS:
                try:
                    record["stories"][kind + ":" + variant] = _read_story(ctx, sections, index, kind, variant)
                except Exception as exc:
                    record["stories"][kind + ":" + variant] = {"error": str(exc)}
        snapshots.append(record)
    return snapshots


def extract_headers_baseline(ctx):
    result = {"sections": [], "errors": [], "explicit_unlink": False}
    try:
        result["sections"] = _read_sections(ctx)
        result["explicit_unlink"] = bool(re.search(r"页眉.{0,12}(?:必须|应|须).{0,8}(?:取消|断开).{0,6}链接到前一节", ctx.text))
    except Exception as exc:
        result["errors"].append(str(exc))
    return {"headers": result}


def _story_signature(story):
    if "error" in story:
        return ("error",)
    return (story["present"], story["static"], story["linked"], tuple(
        (f.get("kind"), f.get("instruction")) for f in story["fields"]),
        repr(story["paragraphs"]), story["objects"], story["complex"], story["borders"])


def _section_signature(section):
    return (section.get("first"), section.get("even"), section.get("number_format"), section.get("start"),
            section.get("chapter"), section.get("header_distance"), section.get("footer_distance"),
            tuple((f.get("instruction"), f.get("in_textbox")) for f in section.get("body_pages", [])),
            tuple((key, _story_signature(value)) for key, value in section["stories"].items()))


def _comparison_baseline(candidates):
    if not candidates:
        return None
    signature = _section_signature(candidates[0])
    if any(_section_signature(other) != signature for other in candidates[1:]):
        return None
    return candidates[0]


def _placeholder(text):
    return bool(re.search(r"_{2,}|[□×]{2,}|[【\[〈<].{0,20}(?:题目|标题|学校|章名|章节).{0,10}[】\]〉>]|填写|替换", text))


def _font_diffs(expected, actual):
    diffs, uncertain = [], []
    for key, wanted in expected.items():
        current = actual.get(key)
        if len(wanted) != 1 or wanted[0] is None or not current or any(v is None for v in current):
            uncertain.append(key)
        elif any(value != wanted[0] for value in current):
            label = key[:2] + _FONT_LABELS.get(key[2:], key[2:])
            attr = key[2:]
            diffs.append((key, f"{label}：模板 {describe_value(attr, wanted[0])}，"
                               f"当前 {'、'.join(describe_value(attr, v) for v in current)}"))
    return diffs, uncertain


def _position_diffs(expected, actual):
    diffs, uncertain = [], []
    labels = {"alignment": "对齐", "left_indent": "左缩进", "right_indent": "右缩进", "first_line_indent": "首行缩进"}
    for key, wanted in expected.items():
        current = actual.get(key)
        if wanted is None or current is None:
            uncertain.append(key)
        elif current != wanted:
            diffs.append((key, f"{labels[key]}：模板 {describe_value(key, wanted)}，"
                               f"当前 {describe_value(key, current)}"))
    return diffs, uncertain


def _check_header_format(expected, actual, label, anchor):
    findings, changes, uncertain = [], [], []
    if expected["complex"] or actual["complex"]:
        uncertain.append("表格、绘图或浮动对象的实际位置/边框")
    eparas, aparas = expected["paragraphs"], actual["paragraphs"]
    if len(eparas) != len(aparas):
        uncertain.append("页眉段落无法逐一对应")
    else:
        for e, a in zip(eparas, aparas):
            for helper, key in ((_font_diffs, "font"), (_position_diffs, "position")):
                diff, missing = helper(e[key], a[key])
                changes.extend(diff)
                uncertain.extend(missing)
            if e["border"] != a["border"]:
                if expected["complex"] or actual["complex"]:
                    uncertain.append("视觉等效线条可能由不同对象表示")
                else:
                    changes.append(("border", f"下边框：模板 {e['border'] or '无'}，当前 {a['border'] or '无'}"))
    if changes:
        # 字体 / 字号 / 对齐 / 下边框 / 缩进 各自成条
        findings.extend(_emit_changes("036", label, changes, anchor, _HEADER_FORMAT_CODES))
    if uncertain:
        findings.append(_limited("036", label + "；未能完整核对：" + "、".join(dict.fromkeys(uncertain)), anchor))
    return findings


def _check_stories(expected, actual, explicit_unlink):
    """页眉/页脚变体检查；文案只说人话，不暴露 True/False 等配置字段。"""

    findings, anchor = _StoryFindings(), actual["anchor"]
    eactive, aactive = _active(expected), _active(actual)
    if expected["even"] != actual["even"]:
        findings.append(_fail("037.2", _even_pages_text(expected["even"]),
                              _even_pages_text(actual["even"])
                              + "；配置与模板不同，本工具不重新分页，实际效果需在 Word 中核对", anchor))
        findings[-1].page_scope = "even_pages"
        findings[-1].page_skip_first = bool(expected["first"] or actual["first"])
        findings[-1].page_config = f"odd_even:{int(expected['even'])}:{int(actual['even'])}"
    if set(eactive) != set(aactive):
        findings.append(_limited("034", "首页/奇偶页的生效选项不一致；仅比较共同启用的页眉类型，不能确定实际页面显示", anchor))
    for variant in (v for v in eactive if v in aactive):
        findings.variant, findings.story = variant, "header"
        label = f"第{actual['index'] + 1}节{_VARIANTS[variant]}"
        e = expected["stories"]["header:" + variant]
        a = actual["stories"]["header:" + variant]
        if "error" in e or "error" in a:
            for rule in ("034", "036", "038"):
                findings.append(_limited(rule, label + "页眉引用或继承关系无法读取", anchor))
        else:
            if e["errors"] or a["errors"] or e["dynamic"] or a["dynamic"] or e["opaque"] or a["opaque"]:
                findings.append(_limited("034", label + "页眉包含动态或损坏的域，无法判断各页最终是否显示内容", anchor))
            elif e["present"] != a["present"]:
                # 模板有内容而文档没有 → 缺少页眉；反之 → 不应出现页眉
                code = "034.1" if e["present"] else "034.2"
                findings.append(_fail(code, label + "页眉配置为" + ("有内容" if e["present"] else "空白"),
                                      "有效页眉部件" + ("有内容" if a["present"] else "没有内容") + "；这是配置差异，尚未确认该类型页面数量", anchor))
            if e["opaque"] or a["opaque"]:
                findings.append(_limited("036", label + "含文本框或OLE等不可读对象，其内字体/段落位置无法确认", anchor))
            elif e["present"] and a["present"]:
                findings.extend(_check_header_format(e, a, label, anchor))
            if actual["index"] > 0 and e["linked"] != a["linked"]:
                if explicit_unlink and a["linked"]:
                    findings.append(_fail("038", "模板明确要求页眉断开链接到前一节", label + "仍继承前节页眉", anchor))
                elif a["linked"] and not (e["dynamic"] or a["dynamic"] or e["opaque"] or a["opaque"]) and not _placeholder(e["static"]) and e["static"] != a["static"]:
                    findings.append(_fail("038", label + "独立页眉内容：" + e["static"],
                                          f"链接到前一节并继承第{a['source'] + 1}节内容：{a['static']}", anchor))
                elif e["dynamic"] or a["dynamic"]:
                    findings.append(_limited("038", label + "链接状态与模板不同，动态内容无法确定；链接差异本身不等于错误", anchor))
            if expected["even"] and variant in {"default", "even"}:
                # 奇偶页分别对应 037.1（普通/奇数页）与 037.2（偶数页）
                even_code = "037.1" if variant == "default" else "037.2"
                if e["dynamic"] or a["dynamic"] or _placeholder(e["static"]) or e["errors"] or a["errors"] or e["opaque"] or a["opaque"]:
                    findings.append(_limited(even_code, label + "动态内容、模板占位内容或域错误不能按模板示例文字比较", anchor))
                else:
                    changes = []
                    if e["static"] != a["static"]:
                        changes.append(f"模板固定文字“{e['static']}”，当前“{a['static']}”")
                    if len(e["paragraphs"]) == len(a["paragraphs"]):
                        for ep, ap in zip(e["paragraphs"], a["paragraphs"]):
                            diffs, _ = _position_diffs(ep["position"], ap["position"])
                            changes.extend(diffs)
                    if expected["header_distance"] is not None and actual["header_distance"] is not None and expected["header_distance"] != actual["header_distance"]:
                        changes.append("页眉距顶端设置与模板不同")
                    if changes:
                        want, actual_text = _split_diff("；".join(changes))
                        findings.append(_fail(even_code, label + "：" + want, actual_text, anchor))
                    if e["complex"] or a["complex"]:
                        findings.append(_limited(even_code, label + "复杂页眉的实际奇偶页位置无法确认", anchor))
        findings.story = "footer"
        e = expected["stories"]["footer:" + variant]
        a = actual["stories"]["footer:" + variant]
        if "error" in e or "error" in a:
            findings.append(_limited("039", label + "页脚关系无法读取", anchor))
            continue
        if e["errors"] or a["errors"] or e["dynamic"] or a["dynamic"] or _placeholder(e["static"]) or e["opaque"] or a["opaque"]:
            findings.append(_limited("039", label + "页脚包含动态/损坏域或模板占位内容，无法确定最终允许文字", anchor))
        else:
            changes = []
            if e["static"].strip() != a["static"].strip():
                changes.append(("footer_text", f"固定文字：模板“{e['static'].strip()}”，当前“{a['static'].strip()}”"))
            # PAGE and NUMPAGES caches were stripped before this comparison.
            ekinds = sorted(f.get("kind") for f in e["fields"] if f.get("kind") != "PAGE")
            akinds = sorted(f.get("kind") for f in a["fields"] if f.get("kind") != "PAGE")
            if ekinds != akinds:
                changes.append(("fields", f"非当前页码域组成：模板 {ekinds}，当前 {akinds}"))
            if e["borders"] != a["borders"] and not (e["complex"] or a["complex"]):
                changes.append(("fields", "段落边框（含上下横线）与模板不同"))
            if changes:
                findings.extend(_emit_changes("039", label, changes, anchor,
                                              {"footer_text": "039.1", "fields": "039.2"}))
        if e["objects"] or a["objects"] or e["complex"] or a["complex"]:
            findings.append(_limited("039", label + "页脚图形/表格边框的视觉效果无法完整读取", anchor))
    # If the paper enables an extra visible variant that the template leaves
    # disabled, Word uses the template's default story on that page type.
    # Compare the actual variant against that effective default so an even-page
    # alignment error is not hidden behind the configuration warning.
    for variant in (v for v in aactive if v not in eactive):
        findings.variant, findings.story = variant, "header"
        label = f"第{actual['index'] + 1}节{_VARIANTS[variant]}"
        e = expected["stories"]["header:default"]
        a = actual["stories"]["header:" + variant]
        if ("error" not in e and "error" not in a and not e["opaque"] and not a["opaque"]
                and e["present"] and a["present"]):
            findings.extend(_check_header_format(e, a, label, anchor))
    findings.variant, findings.story = None, ""
    return findings


def _page_format(page, section):
    instruction = page["instruction"]
    if page["nested"] or re.search(r"\\#|\\@", instruction):
        return None
    formats = re.findall(r"\\\*\s+([^\s\\]+)", instruction)
    formats = [f for f in formats if f.upper() not in {"MERGEFORMAT", "CHARFORMAT"}]
    if len(formats) > 1:
        return None
    if not formats:
        return section["number_format"]
    names = {"ROMAN": "upperRoman", "roman": "lowerRoman", "Arabic": "decimal", "ARABIC": "decimal",
             "ALPHABETIC": "upperLetter", "alphabetic": "lowerLetter"}
    return names.get(formats[0])


def _page_objects(section, variant):
    found, uncertain, manual = [], False, False
    for kind in ("header", "footer"):
        story = section["stories"][kind + ":" + variant]
        if "error" in story:
            uncertain = True
            continue
        uncertain = uncertain or story["page_uncertain"] or story["dynamic"]
        # Plain digits can be literal page numbers, dates, or identifiers. Do not guess.
        manual = manual or bool(re.search(r"\d|[Ⅰ-ⅿ]", story["static"]) or
                                re.fullmatch(r"\s*[ivxlcdmIVXLCDM]+\s*", story["static"]))
        found.extend((kind, page) for page in story["pages"])
    return found, uncertain, manual


def _start_text(value):
    return "续前节编号" if value[0] == "continue" else f"从{value[1]}开始编号"


def _check_pages(expected, actual, continuation_ambiguous=False, has_render_facts=False):
    findings, anchor = _StoryFindings(), actual["anchor"]
    eactive, aactive = _active(expected), _active(actual)
    if set(eactive) != set(aactive):
        findings.append(_limited("040", "首页/奇偶页选项不一致，无法完整对应生效页码区域", anchor))
    body_pages = bool(expected["body_pages"] or actual["body_pages"])
    any_page = False
    if body_pages:
        for rule in ("040", "043"):
            findings.append(_limited(rule, "正文或文本框含PAGE域，无法由共享页眉页脚推算各页显示次数或实际位置", anchor))
        expected_page_info = [_page_objects(expected, v) for v in eactive]
        expected_header_pages = any(pages for pages, _, _ in expected_page_info)
        expected_uncertain = any(uncertain or manual for _, uncertain, manual in expected_page_info)
        if (actual["body_pages"] and not expected["body_pages"] and not expected_header_pages and
                set(actual["roles"]) & _FRONT and "正文" not in actual["roles"] and actual["roles"] == expected["roles"]):
            if expected_uncertain or any(f.get("in_textbox") or f.get("nested") or f.get("malformed") for f in actual["body_pages"]):
                findings.append(_limited("044", "前置正文存在嵌套、文本框或损坏PAGE域，无法确认可见性", anchor))
            else:
                findings.append(_fail("044", "对应模板前置区域未配置PAGE域", "前置区域正文配置了当前页码PAGE域", anchor))
    for variant in (v for v in eactive if v in aactive):
        findings.variant = variant
        label = f"第{actual['index'] + 1}节{_VARIANTS[variant]}"
        epages, euncertain, emanual = _page_objects(expected, variant)
        apages, auncertain, amanual = _page_objects(actual, variant)
        carriers = {kind for kind, _ in (apages or epages)}
        findings.story = next(iter(carriers)) if len(carriers) == 1 else ""
        any_page = any_page or bool(epages or apages)
        if euncertain or auncertain:
            for rule in ("040", "041", "043"):
                findings.append(_limited(rule, label + "页码包含嵌套域、动态内容、文本框或损坏引用", anchor))
            if set(actual["roles"]) & _FRONT:
                findings.append(_limited("044", label + "前置页面的实际页码显示无法读取", anchor))
            continue
        if epages and not apages:
            if amanual or body_pages:
                findings.append(_limited("040", label + "共享页眉页脚未找到PAGE域，但存在普通数字或正文PAGE域，无法确认各页实际页码", anchor))
            else:
                findings.append(_fail("040.1", label + f"配置{len(epages)}个当前页码PAGE域", "未找到当前页码PAGE域；不能自动显示模板要求的页码", anchor))
        elif epages and len(apages) != len(epages):
            findings.append(_fail("040.2", label + f"配置{len(epages)}个当前页码PAGE域", f"配置{len(apages)}个PAGE域；NUMPAGES/SECTIONPAGES未计为当前页码", anchor))
        if not epages and apages and set(actual["roles"]) & _FRONT:
            if emanual or "正文" in actual["roles"] or actual["roles"] != expected["roles"]:
                findings.append(_limited("044", label + "模板区域/普通数字无法证明对应页面应隐藏页码", anchor))
            else:
                findings.append(_fail("044", label + "的对应模板页眉页脚无当前页码域", f"有效候选页眉页脚配置{len(apages)}个PAGE域；尚未确认该类型实际页数", anchor))
        if epages and len(epages) == len(apages):
            # 前置部分应为罗马数字（041.1），正文应为阿拉伯数字（041.2）
            page_code = "041.1" if set(actual["roles"]) & _FRONT else "041.2"
            for (ekind, ep), (akind, ap) in zip(epages, apages):
                efmt, afmt = _page_format(ep, expected), _page_format(ap, actual)
                if efmt is None or afmt is None:
                    findings.append(_limited(page_code, label + "页码域包含不支持的格式开关，无法确定数字形式", anchor))
                elif efmt != afmt or expected["chapter"] != actual["chapter"]:
                    findings.append(_fail(page_code, label + f"数字形式={efmt}，章节前缀设置={expected['chapter']}",
                                          f"数字形式={afmt}，章节前缀设置={actual['chapter']}", anchor))
                changes, missing = _font_diffs(ep["font"], ap["font"])
                diffs, unresolved = _position_diffs(ep["position"], ap["position"])
                changes.extend(diffs)
                missing.extend(unresolved)
                if ekind != akind:
                    changes.append(("carrier", "承载位置：模板在" + ("页眉" if ekind == "header" else "页脚") + "，当前在" + ("页眉" if akind == "header" else "页脚")))
                if expected[ekind + "_distance"] is not None and actual[akind + "_distance"] is not None and expected[ekind + "_distance"] != actual[akind + "_distance"]:
                    changes.append(("distance", "页码所在页眉/页脚距页面边缘设置不同"))
                if changes:
                    # 页码位置（043.1）/ 字体（043.2）/ 字号（043.3）/ 对齐（043.4）各自成条
                    findings.extend(_emit_changes("043", label, changes, anchor, _PAGE_FORMAT_CODES))
                if missing or not ep["font"] or not ap["font"] or ep["in_table"] or ap["in_table"]:
                    findings.append(_limited("043", label + "部分字体、域结果格式或表格内绝对位置无法确认；域缓存数字不作为实际页码", anchor))
    findings.variant, findings.story = None, ""
    if any_page:
        # 分页事实来自 Word 排版时，页码连续性已由调度层核对（word_effective.pagination），
        # 这里不再重复输出“无法判定”。
        if not has_render_facts:
            findings.append(_limited("040", f"第{actual['index'] + 1}节实际分页后的重复/跳号无法计算；PAGE缓存不是每页当前显示值", anchor))
        if continuation_ambiguous:
            findings.append(_limited("042", "多个论文节映射到同一模板区域，无法把区域起始重启要求套到每个后续节", anchor))
        elif expected["start"] != actual["start"]:
            findings.append(_fail("042", _start_text(expected["start"]), _start_text(actual["start"]) + "；比较的是节编号配置", anchor))
        elif expected["start"][0] == "continue":
            findings.append(_limited("042", "已确认续前节配置一致，但实际起始显示值依赖前节分页，无法计算", anchor))
    return findings


def check_header_footer_errors(baseline, ctx, mapping=None):
    """Compare only readable corresponding configurations; never render caches."""
    data = baseline.get("headers", {})
    findings = []
    try:
        actual_sections = _read_sections(ctx)
    except Exception as exc:
        return [_limited(rule, "待检测文档的分节读取失败：" + str(exc), next(iter(ctx.paragraphs), None)) for rule in _TITLES]
    for actual in actual_sections:
        anchor = actual["anchor"]
        section_findings: list = []
        candidates, error = _candidate_sections(data.get("sections", []), actual, mapping)
        expected = _comparison_baseline(candidates)
        header_mapping = dict(mapping or {})
        if header_mapping.get("header_section_map"):
            header_mapping["section_map"] = header_mapping["header_section_map"]
        header_candidates, header_error = _candidate_sections(
            data.get("sections", []), actual, header_mapping)
        header_expected = _comparison_baseline(header_candidates)
        invalid_common = bool(data.get("errors") or actual["errors"])
        if (header_error or header_expected is None or invalid_common
                or header_expected.get("errors")):
            reason = header_error or "模板页眉基准不唯一、节设置读取失败或模板证据不可得；请按区域核对或提供header_section_map"
            section_findings.extend(_limited(rule, reason, anchor)
                                    for rule in ("034", "035", "036", "037", "038"))
        else:
            try:
                section_findings.extend(
                    f for f in _check_stories(
                        header_expected, actual, data.get("explicit_unlink", False))
                    if str(f.rule_id).split(".")[0] in {"034", "035", "036", "037", "038"})
            except Exception as exc:
                section_findings.extend(_limited(
                    rule, "本节页眉对象读取失败：" + str(exc), anchor)
                                        for rule in ("034", "035", "036", "037", "038"))
        if error or expected is None or invalid_common or expected.get("errors"):
            reason = error or "模板页脚或页码基准不唯一、节设置读取失败或模板证据不可得；请按区域核对或提供section_map"
            section_findings.extend(_limited(rule, reason, anchor)
                                    for rule in ("039", "040", "041", "042", "043", "044"))
        else:
            try:
                section_findings.extend(
                    f for f in _check_stories(
                        expected, actual, data.get("explicit_unlink", False))
                    if str(f.rule_id).split(".")[0] == "039")
                repeated_role = (actual["index"] > 0 and
                                 actual["roles"] == actual_sections[actual["index"] - 1]["roles"])
                explicit_map = (mapping or {}).get("section_map", {})
                mapped = (str(actual["index"] + 1) in explicit_map or
                          actual["index"] + 1 in explicit_map)
                section_findings.extend(_check_pages(
                    expected, actual,
                    repeated_role and not mapped and expected["start"][0] == "start",
                    has_render_facts=bool(getattr(ctx, "render_facts", None)),
                ))
            except Exception as exc:
                section_findings.extend(_limited(
                    rule, "本节页脚或页码对象读取失败：" + str(exc), anchor)
                                        for rule in ("039", "040", "041", "042", "043", "044"))
        findings.extend(_tag_page_scope(section_findings, actual))
    return findings


#: 页眉/页脚规则族 → 结论挂在页眉还是页脚；页码类规则按该节实际承载位置判定。
_HEADER_RULES = frozenset({"034", "035", "036", "037", "038"})
_FOOTER_RULES = frozenset({"039"})
_PAGE_RULES = frozenset({"040", "041", "042", "043", "044"})


def _page_story(actual: dict, family: str) -> str:
    """页码类结论的承载位置：该节页脚里有 PAGE 域就算页脚，否则看页眉。"""

    if family in _FOOTER_RULES:
        return "footer"
    if family in _HEADER_RULES:
        return "header"
    for story in ("footer", "header"):
        for variant in ("default", "first", "even"):
            value = actual["stories"].get(story + ":" + variant) or {}
            if value.get("fields"):
                return story
    return "footer"   # 页码默认在页脚：没有读到域时按页脚落点


def _tag_page_scope(section_findings: list, actual: dict) -> list:
    """给本节结论打上“节 + 页眉/页脚”标记，供调度层把它们锚到该页版心顶部/底部。

    Word 不允许把批注写进页眉页脚，只能锚在正文里；带上这个标记后，批注会落在
    受影响的**那一页**（页眉问题→版心顶部、页脚问题→版心底部），而不是整节首段。
    """

    for finding in section_findings:
        family = str(finding.rule_id).split(".")[0]
        finding.section_index = actual["index"] + 1          # 1 起，与 Word 节号一致
        finding.page_story = finding.page_story or _page_story(actual, family)
    return section_findings


def adapt_legacy_header_findings(findings, baseline, ctx, mapping=None):
    """Retain 035 static detection, downgrade unreadable/dynamic comparisons only.

    The legacy COM checker compares complete cached header strings. A dynamic
    STYLEREF/PAGE result can legitimately differ between template and paper.
    This compatibility layer preserves that evidence without claiming a FAIL.
    """
    adapted = list(findings)
    data = baseline.get("headers", {})
    try:
        sections = _read_sections(ctx)
    except Exception as exc:
        for index, finding in enumerate(adapted):
            if finding.rule_id == "035":
                adapted[index] = unknown("035", "页眉文字动态内容需核对", UNREADABLE +
                                         "；节对象无法读取：" + str(exc) + "；原检测证据：" +
                                         " ".join(finding.message.split()), next(iter(ctx.paragraphs), None))
        if not any(f.rule_id == "035" for f in adapted):
            adapted.append(unknown("035", "页眉文字动态内容需核对", UNREADABLE +
                                   "；节对象无法读取：" + str(exc), next(iter(ctx.paragraphs), None)))
        return adapted
    kinds = {1: "default", 2: "first", 3: "even"}
    emitted = set()
    limitations = {}
    for actual in sections:
        candidates, error = _candidate_sections(data.get("sections", []), actual, mapping)
        for kind, variant in kinds.items():
            if not actual.get("errors") and variant not in _active(actual):
                continue
            story = actual["stories"]["header:" + variant]
            expected_stories = [c["stories"]["header:" + variant] for c in candidates]
            # 035 reads text only. Different page numbering, link state or
            # footer layouts must not disable an otherwise common static header.
            def text_signature(value):
                if "error" in value:
                    return None
                return (value["static"], value["opaque"], value["objects"],
                        tuple((f.get("kind"), f.get("instruction")) for f in value["fields"]), tuple(value["errors"]))
            estory = expected_stories[0] if expected_stories and all(
                text_signature(v) == text_signature(expected_stories[0]) for v in expected_stories) else None
            bad = (error or actual.get("errors") or
                   "error" in story or estory is None or "error" in estory)
            if not bad:
                # Cached results of the same dynamic field (for example
                # STYLEREF in the template and paper) are document-specific
                # and cannot be compared as literal text. A field on only
                # one side, however, is a definite header-content mismatch:
                # the template requires fixed text while the paper derives
                # a chapter title (or vice versa).
                expected_fields = tuple(
                    (f.get("kind"), f.get("instruction")) for f in estory["fields"])
                actual_fields = tuple(
                    (f.get("kind"), f.get("instruction")) for f in story["fields"])
                equivalent_dynamic = bool(expected_fields) and expected_fields == actual_fields
                bad = (equivalent_dynamic or story["opaque"] or estory["opaque"] or
                       story["objects"] or estory["objects"] or _placeholder(estory["static"]))
            if bad:
                limitations[(actual["index"] + 1, kind)] = actual["anchor"]
    for index, finding in enumerate(adapted):
        if finding.rule_id != "035":
            continue
        finding.page_story = "header"
        finding.page_variant = kinds.get(finding.header_kind)
        finding.page_scope = "variant"
        key = (finding.section_index, finding.header_kind)
        if key not in limitations:
            continue
        replacement = unknown("035", "页眉文字动态内容需核对",
                              UNREADABLE + "；页眉包含动态域、图片/文本框、占位或无法匹配的模板基准；"
                              "保留原有比对证据，但缓存文字不能证明实际页眉错误。原检测信息：" + " ".join(finding.message.split()),
                              limitations[key])
        replacement.section_index, replacement.header_kind = key
        adapted[index] = replacement
        emitted.add(key)
    for key, anchor in limitations.items():
        if key in emitted:
            continue
        finding = unknown("035", "页眉文字动态内容需核对", UNREADABLE +
                          "；页眉包含动态域、图片/文本框、占位或无法匹配的模板基准，未执行实际页面文字比对", anchor)
        finding.section_index, finding.header_kind = key
        adapted.append(finding)
    return adapted
