"""字体与段落检测（067–077）的取证实现；旧版 Word 规则保持不变。

067/068/072/074/075 and body 069 reuse the existing checkers. Heading size
reuses 045--047, caption size 079, and table-cell size 084. This module adds
069 for readable footnote/endnote text and 070/071/073/076/077. No library
formatting getter/setter is used to modify either input document.
"""
from __future__ import annotations

import re
from collections import defaultdict

from thesis_format_checker.core.xmlutil import qn
from docx.oxml import parse_xml
from docx.text.paragraph import Paragraph

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import (
    UNREADABLE,
    describe_value,
    effective_paragraph,
    make_finding,
    not_applicable,
    unknown,
)
from thesis_format_checker.core.pairing import dominant_value, group_by
from thesis_format_checker.core.scope import front_end_of
from thesis_format_checker.detectors.headings import _heading_level, _num_pr, _toc_paragraph_ids, collect_heading_records


_TITLES = {
    "069": "注释字号与模板不符", "070": "同用途文字格式不一致", "071": "正文对齐方式不符",
    "073": "标题或题注首行缩进不符", "076": "段落悬挂或左右缩进不符", "077": "段落分页设置不符",
}
_BODY_NAMES = {
    "normal", "正文", "正文1", "普通", "bodytext", "bodytextfirstindent", "正文首行缩进",
}
_FRONT = {"封面", "题名页", "摘要", "中文摘要", "英文摘要", "abstract", "目录", "contents",
          "原创性声明", "独创性声明", "授权使用声明"}
_BACK = {"参考文献", "references", "bibliography", "致谢", "acknowledgements", "acknowledgments",
         "附录", "appendix", "攻读学位期间成果"}
_STYLE_EXCLUDED = ("标题", "heading", "title", "目录", "toc", "摘要", "abstract", "关键词", "keyword",
                   "参考文献", "bibliography", "caption", "题注", "页眉", "页脚", "header", "footer",
                   "footnote", "endnote", "脚注", "尾注", "封面", "cover", "equation", "公式")
_NAMES = {"east_asia_font": "中文字体", "ascii_font": "英文/数字字体", "hansi_font": "扩展西文字体", "size_pt": "字号", "bold": "加粗",
          "color": "颜色", "alignment": "对齐", "left_indent": "左缩进", "right_indent": "右缩进",
          "hanging": "悬挂缩进", "widow_control": "孤行控制", "keep_together": "段中不分页",
          "keep_with_next": "与下段同页"}
#: 段落属性 → 细粒度规则编号（按规则分组）
_PROPERTY_CODES = {
    "071": {"alignment": "071"},
    "076": {"hanging": "076.1", "left_indent": "076.2", "right_indent": "076.3"},
    "077": {"widow_control": "077.1", "keep_together": "077.2", "keep_with_next": "077.3"},
}
#: 同用途文字不一致 → 细粒度规则编号
_CONSISTENCY_CODES = {"east_asia_font": "070.1", "ascii_font": "070.1", "hansi_font": "070.1",
                      "size_pt": "070.2", "color": "070.3", "bold": "070.4"}
def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else _TITLES[code]


def _special_indent_code(role: str) -> str:
    """不应缩进的对象：标题 073.1 / 图表题注 073.2 / 参考文献 073.3。"""

    if "图" in role or "表" in role or "caption" in role.casefold():
        return "073.2"
    if "参考" in role or "reference" in role.casefold():
        return "073.3"
    return "073.1"


def _normal(text):
    return re.sub(r"\s+", "", text or "").casefold()


def _is_body_style(name: str) -> bool:
    return name in _BODY_NAMES or bool(re.fullmatch(r"正文\d+", name))


def _styles(paragraph):
    seen, style = set(), paragraph.style
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def _properties(paragraph):
    if paragraph._p.pPr is not None:
        yield paragraph._p.pPr
    for style in _styles(paragraph):
        if style.element.pPr is not None:
            yield style.element.pPr
    roots = paragraph.part.document.styles.element.xpath("./w:docDefaults/w:pPrDefault/w:pPr")
    yield from roots


def _unavailable(rule, reason, paragraph=None):
    return unknown(rule, _title(rule), UNREADABLE + "；" + reason, paragraph)


def _has_revision(paragraph):
    return bool(paragraph._p.xpath(".//w:pPrChange|.//w:rPrChange|.//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo")) or any(
        p.tag in (qn("w:ins"), qn("w:moveFrom"), qn("w:moveTo"), qn("w:del")) for p in paragraph._p.iterancestors())


def _first_indent(paragraph):
    """Read signed first-line amount without converting character units."""
    if _num_pr(paragraph):
        # 列表编号负责“编号+标题文字”的整体缩进，但段落自身显式设置的首行缩进
        # 仍会额外推开标题文字。只读取段落直设值；没有直设值时以 0 表示
        # “未额外缩进”，不把列表级悬挂量误当成标题首行缩进。
        direct = paragraph._p.pPr.find(qn("w:ind")) if paragraph._p.pPr is not None else None
        if direct is None:
            return ("磅", 0)
        for attribute, unit, factor in (
                ("hangingChars", "字符", -0.01), ("firstLineChars", "字符", 0.01),
                ("hanging", "磅", -0.05), ("firstLine", "磅", 0.05)):
            raw = direct.get(qn("w:" + attribute))
            if raw is not None:
                return (unit, int(raw) * factor)
        return ("磅", 0)
    props = list(_properties(paragraph))
    values = {}
    for attribute in ("firstLineChars", "hangingChars", "firstLine", "hanging"):
        for prop in props:
            node = prop.find(qn("w:ind"))
            if node is not None and node.get(qn("w:" + attribute)) is not None:
                values[attribute] = int(node.get(qn("w:" + attribute)))
                break
    if "hangingChars" in values:
        return ("字符", -values["hangingChars"] / 100)
    if "firstLineChars" in values:
        return ("字符", values["firstLineChars"] / 100)
    if "hanging" in values:
        return ("磅", -values["hanging"] / 20)
    return ("磅", values.get("firstLine", 0) / 20)


def _paragraph_format(paragraph):
    result = effective_paragraph(paragraph)
    # ISO 29500 absence semantics: widow/orphan prevention is enabled when
    # omitted throughout the cascade. This is not a school-format default.
    if result["widow_control"] is None:
        result["widow_control"] = True
    if _num_pr(paragraph):
        result["left_indent"] = result["right_indent"] = None
    result["hanging"] = _first_indent(paragraph)
    for prop in _properties(paragraph):
        bidi = prop.find(qn("w:bidi"))
        if bidi is not None:
            if bidi.get(qn("w:val"), "1") not in ("0", "false", "off"):
                result["alignment"] = None
            break
    return result


def _body_records(ctx, mapping=None, side="document"):
    """Conservative normal-body scope, excluding structural/template regions.

    摘要页之前（封面、题名页、声明页）不属于正文范围：那里的文字只由封面模块
    按模板同位置比对，不参与正文对齐、缩进、字体一致性等通用检测。
    """
    settings = (mapping or {}).get(side, {})
    settings = settings if isinstance(settings, dict) else {}
    chosen = {_normal(str(v)) for v in settings.get("body_styles", [])}
    toc = _toc_paragraph_ids(ctx)
    front_end = front_end_of(ctx)
    heading_roles = {record["paragraph"]._p: record for record in collect_heading_records(ctx)}
    heading_levels = {}
    for p in ctx.paragraphs:
        try:
            heading_levels[p._p] = _heading_level(p)
        except Exception:
            heading_levels[p._p] = None
    first_heading = next((i for i, p in enumerate(ctx.paragraphs)
                          if heading_levels[p._p] and _normal(heading_roles.get(p._p, {}).get("title", p.text)) not in _FRONT | _BACK and p._p not in toc), None)
    active = first_heading is None
    ended = False
    range_active = not settings.get("start_text")
    records = []
    for index, p in enumerate(ctx.paragraphs, 1):
        if 0 < front_end and index <= front_end:
            continue  # 封面/声明等前置区：不进入正文范围（摘要从 front_end+1 段开始）
        text = p.text.strip()
        normalized = _normal(heading_roles.get(p._p, {}).get("title", text))
        if settings.get("start_text") and str(settings["start_text"]) in text:
            range_active = True
            active = True
            continue
        if settings.get("end_text") and str(settings["end_text"]) in text:
            break
        if normalized in _FRONT | _BACK:
            active = False
            ended = ended or normalized in _BACK
            continue
        if ended:
            continue
        if heading_levels[p._p] and p._p not in toc:
            active = True
            continue
        if not range_active or (not active and not chosen) or not text or p._p in toc:
            continue
        if settings.get("start_paragraph") and index < int(settings["start_paragraph"]):
            continue
        if settings.get("end_paragraph") and index > int(settings["end_paragraph"]):
            continue
        if any(a.tag == qn("w:tc") for a in p._p.iterancestors()):
            continue
        names = [_normal(s.name) for s in _styles(p)]
        own = names[0] if names else ""
        if chosen:
            if own not in chosen:
                continue
        elif not any(_is_body_style(name) for name in names) or any(token in own for token in _STYLE_EXCLUDED):
            continue
        # 图题/表题（含英文缩写 Fig./Tab.）属于独立用途，不当作正文比对。
        if re.match(r"^(?:图|表|figure\b|table\b|fig\.|tab\.)\s*\d|^(?:关键词|key\s*words|论文题目|指导教师|作者姓名)\s*[:：]", text, re.I):
            continue
        # 引用、交叉引用等域不会改变段落级对齐/缩进/行距，也不妨碍域外文字
        # 的有效字体读取；不能因为段内有引用域就整段跳过。公式和嵌入对象仍排除。
        if p._p.xpath(".//m:oMath|.//w:object|.//w:drawing|.//w:pict"):
            continue
        role = ("正文" if any(_is_body_style(name) for name in names) or len(chosen) == 1
                else "正文样式：" + (p.style.name or own))
        num = _num_pr(p)
        if num:
            role += "（列表段落，级别" + str(num.get("ilvl", "未明确")) + "）"
        records.append({"paragraph": p, "role": role, "text": text})
    return records


def _special_records(ctx):
    from thesis_format_checker.detectors.figures import collect_caption_records
    records = []
    toc = _toc_paragraph_ids(ctx)
    heading_roles = {record["paragraph"]._p: record for record in collect_heading_records(ctx)}
    reference_nodes, in_references = set(), False
    for p in ctx.paragraphs:
        if p._p in toc or not p.text.strip():
            continue
        level = _heading_level(p)
        title = _normal(heading_roles.get(p._p, {}).get("title", p.text))
        if title in {"参考文献", "references", "bibliography"}:
            in_references = True
        elif title in _FRONT | _BACK:
            in_references = False
        if in_references:
            reference_nodes.add(p._p)
            continue
        # 073参考文献：待开会讨论，本版本暂不开发；不检查文献条目。
        if level is not None:
            role = f"{level}级标题"
            if title in _FRONT | _BACK:
                role += "（" + title + "）"
            records.append({"paragraph": p, "role": role, "text": p.text})
    for item in collect_caption_records(ctx):
        if item["paragraph"]._p in reference_nodes:
            continue
        role = "图题" if item["kind"] == "figure" else "表题"
        role += "（" + str(item.get("language", "")) + "，" + str(item.get("role", "body")) + "）"
        records.append({"paragraph": item["paragraph"], "role": role, "text": item["text"]})
    return records


def _run_groups(record):
    from thesis_format_checker.detectors.text_runs import effective_text_runs
    paragraph = record["paragraph"]
    effective = effective_text_runs(paragraph)
    groups = []
    for text, props in effective:
        if not text.strip():
            continue
        run = props["run"]
        if props.get("revision"):
            raise ValueError("文字含未接受修订，有效用途无法确认")
        if props.get("field"):
            # 只跳过域结果本身（引用号、交叉引用等）；同一段其余普通文字仍可
            # 精确比较字体和字号，不能因一个引用域放弃整段。
            continue
        rstyle = run.style
        style_name = _normal(rstyle.name) if rstyle is not None else ""
        if style_name in ("defaultparagraphfont", "默认段落字体", "默认段落字型"):
            style_name = "普通字符"
        # Hyperlinks, raised note markers and equation scripts are distinct
        # uses, never evidence that plain body font has become inconsistent.
        special = any(a.tag == qn("w:hyperlink") for a in run._r.iterancestors()) or run._r.xpath(
            ".//w:footnoteReference|.//w:endnoteReference|.//w:footnoteRef|.//w:endnoteRef")
        special = special or props.get("vertical_align") in ("superscript", "subscript")
        if special:
            continue
        for language, pattern, font in (("中文", r"[\u3400-\u9fff\U00020000-\U0003134f]", "east_asia_font"),
                                        ("英文", r"[A-Za-z]", "ascii_font"), ("数字", r"[0-9]", "ascii_font"),
                                        ("扩展西文", r"[\u00c0-\u024f]", "hansi_font")):
            if re.search(pattern, text):
                values = {key: props.get(key) for key in (font, "size_pt", "bold", "color")}
                if values[font] is not None:
                    values[font] = str(values[font]).strip().casefold()
                groups.append(((record["role"], style_name, language), text, values))
    return groups


def _note_records(ctx):
    records, issues = [], []
    for kind, label in (("footnote", "脚注"), ("endnote", "尾注")):
        relations = [rel for rel in ctx.document.part.rels.values() if rel.reltype.endswith("/" + kind + "s")]
        refs = {}
        for p in ctx.paragraphs:
            for node in p._p.iter(qn("w:" + kind + "Reference")):
                refs.setdefault(node.get(qn("w:id")), p)
        if refs and not relations:
            issues.append((label + "存在引用标记但对象关系缺失", next(iter(refs.values()))))
        for rel in relations:
            try:
                if rel.is_external:
                    raise ValueError("注释内容关系指向外部")
                root = parse_xml(rel.target_part.blob)
                seen = set()
                for note in root.findall(qn("w:" + kind)):
                    if note.get(qn("w:type"), "normal") != "normal":
                        continue
                    identity = note.get(qn("w:id"))
                    if identity not in refs:
                        continue
                    seen.add(identity)
                    for element in note.iter(qn("w:p")):
                        if any(a.tag == qn("w:txbxContent") for a in element.iterancestors()):
                            continue
                        p = Paragraph(element, ctx.document._body)
                        if p.text.strip():
                            records.append({"paragraph": p, "anchor": refs[identity], "role": label,
                                            "text": p.text, "identity": identity})
                    if any(node.tag in (qn("w:txbxContent"), qn("w:object"), qn("w:drawing")) for node in note.iter()):
                        issues.append((label + "包含无法读取内部文字格式的对象", refs[identity]))
                for identity, p in refs.items():
                    if identity not in seen:
                        issues.append((label + "引用所指内容缺失", p))
            except Exception as exc:
                issues.append((label + "对象读取异常：" + type(exc).__name__, next(iter(refs.values()), None)))
    return records, issues


def _consensus(values):
    if not values or any(value is None for value in values):
        return None
    distinct = set(values)
    return next(iter(distinct)) if len(distinct) == 1 else None


def extract_paragraph_baseline(ctx, mapping=None):
    records = _body_records(ctx, mapping, "template")
    formats, texts, specials, notes = defaultdict(list), defaultdict(list), defaultdict(list), defaultdict(list)
    issues, text_issues = [], set()
    for record in records:
        p = record["paragraph"]
        try:
            if _has_revision(p):
                raise ValueError("模板包含未接受的修订")
            formats[record["role"]].append(_paragraph_format(p))
            for key, _, values in _run_groups(record):
                texts[key].append(values)
        except Exception as exc:
            issues.append(record["role"] + "：" + str(exc))
            text_issues.add(record["role"])
            formats[record["role"]].append({})
    for record in _special_records(ctx):
        try:
            value = None if _has_revision(record["paragraph"]) else _first_indent(record["paragraph"])
            specials[record["role"]].append(value)
        except Exception:
            specials[record["role"]].append(None)
        try:
            if _has_revision(record["paragraph"]):
                raise ValueError("模板标题或题注含修订")
            for key, _, values in _run_groups(record):
                texts[key].append(values)
        except Exception:
            text_issues.add(record["role"])
    note_records, note_issues = _note_records(ctx)
    for record in note_records:
        try:
            for key, _, values in _run_groups(record):
                notes[key].append(values["size_pt"] if not _has_revision(record["paragraph"]) else None)
                texts[key].append(values)
            if _has_revision(record["paragraph"]):
                text_issues.add(record["role"])
        except Exception:
            note_issues.append((record["role"] + "有效字号读取异常", record["anchor"]))
            text_issues.add(record["role"])
    if note_issues:
        text_issues.update(("脚注", "尾注"))
    return {"paragraphs": {"formats": dict(formats), "texts": dict(texts), "specials": dict(specials),
                                   "notes": dict(notes), "issues": issues, "note_issues": note_issues,
                                   "text_issues": text_issues}}


def _format_value(key, value):
    if key == "hanging" and value is not None:
        return f"{value[1]:g}{value[0]}（负值表示悬挂）"
    return describe_value(key, value)


def _basis_sample(samples, ordinal, item_count, keys):
    """取比对基准：数量相同时按“同一位置”逐个对应，否则取同层级最常见样例。

    模板与论文的段落数往往不同（例如模板只放了几段样张），逐位置硬对齐会把
    封面声明错配到摘要说明段；因此只在数量相同时使用位置对应，
    其余情况退回“同一层级”的典型值，而不是要求全篇唯一。
    """

    if not samples:
        return None
    if len(samples) == item_count:
        return samples[ordinal]
    return {key: dominant_value((item.get(key), 1) for item in samples) for key in keys}


def _basis_value(values, ordinal, item_count):
    """列表值（如首行缩进）的同类基准：数量相同按位置，否则取最常见样例。"""

    if not values:
        return None
    if len(values) == item_count:
        return values[ordinal]
    return dominant_value((value, 1) for value in values)


def _check_properties(rule, record, sample, keys):
    """把论文段落与**模板中同一层级、同一位置**的段落逐属性比较。

    ``sample`` 由 :func:`check_paragraph_errors` 按相对位置选出（见 ``core/pairing.py``），
    因此模板里存在多种样例时不再对全篇输出“无法判定”，而是各段比各自对应的那一段。
    """

    p = record["paragraph"]
    if _has_revision(p):
        return [_unavailable(rule, "含未接受的修订，无法确定当前有效段落格式", p)]
    if sample is None:
        return [not_applicable(rule, _TITLES[rule], "模板没有可确认的对应" + record["role"] + "格式样例", p)]
    actual = _paragraph_format(p)
    expected_parts, actual_parts, uncertain = [], [], []
    diffs = []
    for key in keys:
        expected = sample.get(key)
        current = actual.get(key)
        if key == "hanging" and expected is not None and current is not None:
            # Positive first-line indentation remains exclusively in 072.
            if expected[1] >= 0 and current[1] >= 0:
                continue
            if expected[0] != current[0] and expected[1] == current[1] == 0:
                continue
        if expected is None or current is None:
            uncertain.append(_NAMES[key])
        elif current != expected:
            diffs.append((key, _NAMES[key] + "=" + _format_value(key, expected),
                          _NAMES[key] + "=" + _format_value(key, current)))
    findings = []
    for key, want, current_text in diffs:
        code = _PROPERTY_CODES.get(rule, {}).get(key, rule)
        findings.append(make_finding(code, _title(code), record["role"] + "：" + want, current_text, p))
    if uncertain:
        findings.append(_unavailable(rule, "模板同用途存在冲突或下列有效值不可读：" + "、".join(uncertain), p))
    return findings


def _check_consistency(data, records):
    groups, unreadable, findings = defaultdict(list), [], []
    for record in records:
        p = record["paragraph"]
        anchor = record.get("anchor", p)
        try:
            if _has_revision(p):
                raise ValueError("含未接受修订")
            for key, text, values in _run_groups(record):
                groups[key].append((p, text, values))
        except Exception as exc:
            unreadable.append((anchor, str(exc)))
    bad, uncertain = defaultdict(list), defaultdict(set)
    for key, group in groups.items():
        if key[0] in data.get("text_issues", set()):
            for p, _, _ in group:
                uncertain[p._p].add("模板对应用途部分文字读取失败或含修订，基准不完整")
            continue
        template = data.get("texts", {}).get(key, [])
        if not template:
            for p, _, _ in group:
                uncertain[p._p].add("模板缺少对应字符用途：" + "/".join(key))
            continue
        for prop in group[0][2]:
            current_values = {item[2].get(prop) for item in group}
            # 全组一致时属于 067–069（统一格式不符）的范围，070 只指出组内不一致。
            if len(current_values) < 2:
                continue
            if prop == "color" and "auto" in current_values:
                for p, _, _ in group:
                    uncertain[p._p].add("自动颜色的实际显示无法与固定颜色直接等同")
                continue
            # 同一层级：取模板该用途下最常见的取值，不再要求全篇唯一。
            expected = dominant_value((item.get(prop), 1) for item in template)
            if expected is None:
                for p, _, _ in group:
                    uncertain[p._p].add(_NAMES[prop] + "在模板同用途中无法读取")
                continue
            for p, text, values in group:
                current = values.get(prop)
                if current is None:
                    uncertain[p._p].add(_NAMES[prop] + "在模板同用途中无法读取")
                    continue
                if current != expected:
                    bad[p._p].append((key, text, prop, expected, current))
    for record in records:
        p = record["paragraph"]
        anchor = record.get("anchor", p)
        if p._p in bad:
            # 一个属性一条结论：字体 070.1 / 字号 070.2 / 颜色 070.3 / 加粗 070.4
            by_property = defaultdict(list)
            for key, text, prop, exp, cur in bad[p._p]:
                by_property[prop].append((key, text, exp, cur))
            for prop, entries in by_property.items():
                code = _CONSISTENCY_CODES.get(prop, "070")
                expected = list(dict.fromkeys("/".join(key) + "的" + _NAMES[prop] + "=" + _format_value(prop, exp)
                                              for key, _, exp, _ in entries))
                actual = list(dict.fromkeys("“" + text[:50] + "”的" + _NAMES[prop] + "=" +
                                            _format_value(prop, cur)
                                            for _, text, _, cur in entries))
                findings.append(make_finding(code, _title(code), "；".join(expected), "；".join(actual), anchor))
        if p._p in uncertain:
            findings.append(not_applicable("070", _TITLES["070"], "；".join(sorted(uncertain[p._p])), anchor))
    findings.extend(_unavailable("070", reason, p) for p, reason in unreadable)
    return findings


def check_paragraph_errors(baseline, ctx, mapping=None):
    data = baseline.get("paragraphs", {})
    if data.get("unreadable"):
        return [_unavailable(rule, data["unreadable"]) for rule in _TITLES]
    findings = []
    records = _body_records(ctx, mapping)
    for role, group in group_by(records, lambda item: item["role"]).items():
        samples = data.get("formats", {}).get(role, [])
        for ordinal, record in enumerate(group):
            basis_keys = ("alignment", "left_indent", "right_indent", "hanging",
                          "widow_control", "keep_together", "keep_with_next")
            sample = _basis_sample(samples, ordinal, len(group), basis_keys)
            for rule, keys in (("071", ("alignment",)), ("076", ("left_indent", "right_indent", "hanging")),
                               ("077", ("widow_control", "keep_together", "keep_with_next"))):
                try:
                    findings.extend(_check_properties(rule, record, sample, keys))
                except Exception as exc:
                    findings.append(_unavailable(rule, type(exc).__name__, record["paragraph"]))
    special_records = _special_records(ctx)
    notes, note_issues = _note_records(ctx)
    findings.extend(_check_consistency(data, records + special_records + notes))
    for role, group in group_by(special_records, lambda item: item["role"]).items():
        values = data.get("specials", {}).get(role, [])
        for ordinal, record in enumerate(group):
            p = record["paragraph"]
            try:
                # 同级标题/同类题注：数量相同按位置对应，否则取同级最常见样例。
                expected = _basis_value(values, ordinal, len(group))
                actual = _first_indent(p)
                indent_code = _special_indent_code(role)
                if not values:
                    findings.append(not_applicable(indent_code, _title(indent_code),
                                                   "模板没有对应" + role + "缩进基准", p))
                elif expected is None or actual is None or _has_revision(p):
                    findings.append(_unavailable(indent_code, "列表定位、修订、单位或模板样例不可读导致首行缩进无法确认", p))
                elif expected[1] <= 0 < actual[1]:
                    findings.append(make_finding(indent_code, _title(indent_code), role + "首行=" + _format_value("hanging", expected),
                                                 "首行=" + _format_value("hanging", actual), p))
            except Exception as exc:
                findings.append(_unavailable("073", type(exc).__name__, p))
    for record in notes:
        p = record["anchor"]
        bad, uncertain = [], []
        try:
            if _has_revision(record["paragraph"]):
                raise ValueError("注释包含未接受的修订")
            for key, text, values in _run_groups(record):
                expected = None if record["role"] in data.get("text_issues", set()) else _consensus(data.get("notes", {}).get(key, []))
                actual = values["size_pt"]
                if expected is None or actual is None:
                    uncertain.append("/".join(key) + "有效字号或模板对应基准不可确认")
                elif expected != actual:
                    bad.append((text, expected, actual))
            if bad:
                # 注释字号 → 069.4
                findings.append(make_finding("069.4", _title("069.4"), record["role"] + "字号：" + "；".join(
                    dict.fromkeys(f"{exp:g}磅" for _, exp, _ in bad)),
                    record["role"] + "（引用标记处批注）" + "；".join(dict.fromkeys(
                        f"“{text[:50]}”为{actual:g}磅" for text, _, actual in bad)), p))
            if uncertain:
                findings.append(_unavailable("069.4", "；".join(dict.fromkeys(uncertain)), p))
        except Exception as exc:
            findings.append(_unavailable("069", record["role"] + "：" + str(exc), p))
    findings.extend(_unavailable("069", reason, p) for reason, p in note_issues)
    if data.get("note_issues") and notes:
        findings.append(_unavailable("069", "模板注释存在无法读取的内容，相关基准可能不完整", notes[0]["anchor"]))
    return findings
