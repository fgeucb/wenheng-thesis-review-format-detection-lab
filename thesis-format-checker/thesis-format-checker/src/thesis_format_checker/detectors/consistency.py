"""Read-only internal consistency checks 134, 136 and 137.

133、135、138、140、142、143：本条依赖大模型语义，本阶段暂不实现。
参考文献引用判定（含141）：待开会讨论暂不实现。
These checks compare objects within the submitted document, never its thesis
title against the template's sample title. Existing checks remain untouched.
"""
from __future__ import annotations

from collections import defaultdict
import re

from docx.oxml import parse_xml
from thesis_format_checker.core.xmlutil import qn
from docx.text.paragraph import Paragraph

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, make_finding
from thesis_format_checker.detectors.front import _fields, _label, _placeholder
from thesis_format_checker.detectors.fields import scan_fields, scan_body_fields
from thesis_format_checker.detectors.headings import collect_heading_records
from thesis_format_checker.detectors.headers import _flag, _read_story
from thesis_format_checker.detectors.toc import _snapshot, _check_links, _page_expected
from thesis_format_checker.detectors.figures import collect_caption_records, _visible, _revised
from thesis_format_checker.detectors.equations import _read_records
from thesis_format_checker.detectors.notes import _bookmark_text
from thesis_format_checker.core.models import Status


TITLES = {"134": "论文题目在各处表述不一致", "136": "目录与正文标题、编号或页码不一致",
          "137": "正文引用编号与实际对象不一致"}
_REQUIREMENT = "内部一致性要求（非模板固定值）"
_TITLE_BOOKMARK = re.compile(r"^(?:thesis[_ ]?title|论文题目|论文题名)(?:[_ ]?(cn|zh|en|中文|英文))?$", re.I)
_NUMBER = r"[0-9０-９]+(?:[.．\-－][0-9０-９]+)*[a-z]?"
_LITERAL = re.compile(r"(?P<scope>本文|本章|本节)\s*(?:的\s*)?(?:(?P<kind>图|表|公式|式)\s*[（(]?(?P<num>" +
                      _NUMBER + r")(?![A-Za-z0-9０-９.．])[）)]?|第(?P<chapter>[0-9０-９]+(?:[.．][0-9０-９]+)*|[零〇一二两三四五六七八九十百千]+)(?P<unit>章|节))", re.I)
_LABEL = re.compile(r"(?:图|表|公式|式)\s*[（(]?" + _NUMBER + r"|第[0-9零一二三四五六七八九十]+[章节]", re.I)


def _finding(rule, actual, anchor=None, expected="相关对象的可读取内容应一致", status=Status.FAIL):
    def one_line(value):
        return str(value).replace("\r\n", "\n").replace("\r", "\n").replace("\n", "↵")
    return make_finding(rule, _title(rule), _REQUIREMENT + "；" + one_line(expected), one_line(actual), anchor, status)


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else TITLES[code]


def _title_source_code(source: str) -> str:
    """题目所在位置 → 134.3 页眉 / 134.2 摘要正文 / 134.1 封面声明。"""

    value = str(source)
    if "页眉" in value or "header" in value.casefold():
        return "134.3"
    if "摘要" in value or "正文" in value:
        return "134.2"
    return "134.1"


def _reference_code(text) -> str:
    """引用文字/对象类型 → 137.1 图表 / 137.2 章节 / 137.3 公式。"""

    value = str(text)
    if any(token in value for token in ("图", "表", "Figure", "Table", "figure", "table")):
        return "137.1"
    if any(token in value for token in ("章", "节")):
        return "137.2"
    if any(token in value for token in ("式", "公式", "equation")):
        return "137.3"
    return "137"


def _unknown(rule, detail, anchor=None):
    return _finding(rule, "暂无法判定：" + UNREADABLE + "；" + detail, anchor, status=Status.UNKNOWN)


def _fallback(ctx):
    return ctx.paragraphs[0] if ctx.paragraphs else None


def _dynamic(p):
    return bool(p._p.xpath(".//w:fldSimple|.//w:fldChar|.//w:instrText|.//w:object|.//w:drawing|.//w:pict|.//w:txbxContent"))


def _title_text(value):
    # Only layout breaks are ignored; spaces, case, punctuation and subtitles
    # remain meaningful text. In particular no Chinese/English translation.
    return re.sub(r"[\r\n]+", "", value).strip()


def _title_records(ctx):
    result = []
    for record in _fields(ctx):
        if record["role"] not in {"title_cn", "title_en"}:
            continue
        # A generic '题目' after the front matter may be an exercise heading.
        if not record["cover"] and record["label"].strip().casefold() in {"题目", "题名", "title"}:
            continue
        if not record["value"].strip() or _placeholder(record["value"], record["role"]):
            continue
        p = record["paragraph"]
        result.append({**record, "unknown": _dynamic(p) or _revised(p._p), "source": "正文第" + str(record["index"] + 1) + "段"})
    for start in ctx.document.element.body.iter(qn("w:bookmarkStart")):
        name = start.get(qn("w:name"), "")
        match = _TITLE_BOOKMARK.fullmatch(name)
        if not match:
            continue
        pnode = next((a for a in start.iterancestors() if a.tag == qn("w:p")), None)
        if pnode is None or any(r["paragraph"]._p is pnode for r in result):
            continue
        p = Paragraph(pnode, ctx.document._body)
        value = _bookmark_text(ctx, name)
        language = match[1]
        role = "title_en" if language and language.casefold() in {"en", "英文"} else "title_cn"
        if language is None and value and not re.search(r"[\u3400-\u9fff]", value):
            role = "title_en"
        result.append({"role": role, "paragraph": p, "value": value or "", "unknown": value is None,
                       "source": "题目书签“" + name + "”", "cover": False})
    return result


def _header_stories(ctx):
    sections = list(ctx.document.sections)
    seen = set()
    for index, section in enumerate(sections):
        try:
            variants = ["default"]
            if _flag(section._sectPr, "titlePg"):
                variants.append("first")
            if _flag(ctx.document.settings.element, "evenAndOddHeaders"):
                variants.append("even")
        except Exception as exc:
            yield "第" + str(index + 1) + "节页眉", None, type(exc).__name__, index + 1, None
            continue
        for variant in variants:
            try:
                # Reuse the established read-only link-to-previous resolver.
                # Disabled first/even parts and orphan package relationships
                # must never be evidence of a visible thesis-title mismatch.
                story = _read_story(ctx, sections, index, "header", variant)
                if story["source"] is None:
                    continue
                source = sections[story["source"]]._sectPr
                ref = next(r for r in source.findall(qn("w:headerReference")) if r.get(qn("w:type")) == variant)
                relation = ctx.document.part.rels[ref.get(qn("r:id"))]
                name = str(relation.target_part.partname)
                if name in seen:
                    continue
                seen.add(name)
                if story["opaque"]:
                    yield "第" + str(index + 1) + "节" + variant + "页眉（" + name + "）", None, "页眉含隐藏文字、修订或不透明对象，显示题目不能唯一读取", index + 1, variant
                    continue
                blob = relation.target_part.blob
                if b"<!DOCTYPE" in blob.upper() or b"<!ENTITY" in blob.upper():
                    raise ValueError("unsupported XML declaration")
                yield "第" + str(index + 1) + "节" + variant + "页眉（" + name + "）", parse_xml(blob), None, index + 1, variant
            except Exception as exc:
                yield "第" + str(index + 1) + "节" + variant + "页眉", None, type(exc).__name__, index + 1, variant


def extract_consistency_baseline(ctx, mapping=None):
    # Store role evidence only, never the sample thesis title as a requirement.
    titles = _title_records(ctx)
    header_roles = set()
    for _, root, error, _, _ in _header_stories(ctx):
        if error:
            continue
        for element in root.iter(qn("w:p")):
            value = _visible(element).strip()
            label = _label(value)
            if label and label[0] in {"title_cn", "title_en"}:
                header_roles.add(label[0])
            for record in titles:
                if not record["unknown"] and value and _title_text(value) == _title_text(record["value"]):
                    header_roles.add(record["role"])
    return {"consistency": {"header_title_roles": tuple(sorted(header_roles))}}


def _check_titles(base, ctx):
    records = _title_records(ctx)
    findings = []
    grouped = defaultdict(list)
    for record in records:
        if record["unknown"]:
            findings.append(_unknown("134", record["source"] + "包含域、修订或不可读对象，题目不能唯一读取", record["paragraph"]))
        else:
            grouped[record["role"]].append(record)
    for role, group in grouped.items():
        reference = next((r for r in group if r["cover"]), group[0])
        for record in group:
            if _title_text(record["value"]) != _title_text(reference["value"]):
                findings.append(_finding(_title_source_code(record["source"]),
                                         record["source"] + "题目为“" + record["value"] + "”", record["paragraph"],
                                         reference["source"] + "同语种完整题目为“" + reference["value"] + "”"))
    for story, root, error, section_index, variant in _header_stories(ctx):
        if error:
            findings.append(_unknown("134", "页眉故事读取失败：" + error, _fallback(ctx)))
            continue
        for node in root.iter(qn("w:p")):
            value = _visible(node).strip()
            label = _label(value)
            explicit = label is not None and label[0] in {"title_cn", "title_en"}
            relevant_roles = {label[0]} if explicit else set(base.get("header_title_roles", ()))
            if not relevant_roles:
                continue
            # 页眉题目也必须走页边界写入，不把封面题目作为替代锚点。
            anchor = None
            if node.xpath(".//w:fldSimple|.//w:fldChar|.//w:instrText|.//w:txbxContent|.//w:object") or _revised(node):
                findings.append(_unknown("134", story + "题目页眉含动态域、修订或嵌入对象；未刷新域", anchor))
                continue
            if explicit:
                role, _, current, _ = label
                candidates = grouped.get(role, [])
                values = {_title_text(r["value"]) for r in candidates}
                if len(values) != 1 or not current.strip():
                    findings.append(_unknown("134", story + "无唯一的同语种完整论文题目可供对照", anchor))
                elif _title_text(current) not in values:
                    findings.append(_finding(_title_source_code(story),
                                             story + "明确标注的论文题目为“" + current + "”", anchor,
                                             "本论文同语种完整题目为“" + candidates[0]["value"] + "”"))
                    findings[-1].section_index = section_index
                    findings[-1].page_story = "header"
                    findings[-1].page_variant = variant
                    findings[-1].page_scope = "variant"
            elif value:
                candidates = [r for role in relevant_roles for r in grouped.get(role, [])]
                if not any(_title_text(value) == _title_text(r["value"]) for r in candidates):
                    findings.append(_unknown("134", story + "页眉未明确标注完整题目，可能采用简称或章节名，不能断言题目错误", anchor))
    if ctx.has_textboxes:
        findings.append(_unknown("134", "文档含文本框，框内是否有需对照的论文题目不能完整确认", _fallback(ctx)))
    return findings


def _check_directory(ctx, mapping):
    data = _snapshot(ctx)
    findings = []
    if not data["groups"]:
        if data["manual"] or data["titles"]:
            findings.append(_unknown("136", "目录没有可解析TOC域及唯一目标关系，不能按同名文字猜配正文", data["titles"][0] if data["titles"] else data["manual"][0]))
        return findings
    if data["errors"]:
        return [_unknown("136", "目录域结构不完整：" + "；".join(data["errors"]), _fallback(ctx))]
    for group in data["groups"]:
        anchor = group["anchor"] or _fallback(ctx)
        field = group["field"]
        if field["malformed"] or field["in_textbox"] or field["dirty"] or field["locked"]:
            findings.append(_unknown("136", "目录域处于不可解析、待更新或锁定状态，无法可靠核对当前缓存", anchor))
            continue
        # Same submitted TOC on both sides deliberately removes template title
        # values and template scope as an internal consistency dependency.
        for result in _check_links(group, group, ctx, mapping):
            if result.rule_id == "058":
                if result.status == Status.FAIL:
                    findings.append(_finding("136", result.actual, result.anchor,
                                             "目录保存文字/编号与唯一书签对应的当前正文及本目录层级映射一致"))
                else:
                    findings.append(_unknown("136", result.actual, result.anchor))
            elif result.rule_id == "057" and result.status == Status.UNKNOWN:
                findings.append(_unknown("136", result.actual, result.anchor))
        if not group["entries"] or any(_page_expected(group["options"], e["level"]) for e in group["entries"]):
            findings.append(_unknown("136", "python-docx不能重新分页，目录页码缓存不能作为正文实际显示页码；只核对可读取的标题和编号", anchor))
    return findings


def _key(value):
    return value.translate(str.maketrans("０１２３４５６７８９．－", "0123456789.-"))


def _objects(ctx, mapping):
    captions = collect_caption_records(ctx)
    equations, problems = _read_records(ctx)
    headings = collect_heading_records(ctx, mapping)
    records = []
    incomplete = {"figure": ctx.has_textboxes, "table": ctx.has_textboxes,
                  "equation": ctx.has_textboxes or bool(problems), "chapter": ctx.has_textboxes, "section": ctx.has_textboxes}
    if _revised(ctx.document.element.body):
        # Deleted/moved objects may be omitted by the lexical collectors. An
        # unaccepted document version cannot prove any target's absence.
        incomplete = dict.fromkeys(incomplete, True)
    consumed = set()
    for caption in captions:
        p = caption["paragraph"]
        consumed.add(p._p)
        reliable = bool(caption["number"] and not caption["dynamic"] and not caption["revision"] and caption["object"] is not None)
        incomplete[caption["kind"]] |= not reliable
        records.append({"kind": caption["kind"], "number": _key(caption["number"]) if caption["number"] else None,
                        "paragraph": p, "elements": {p._p}, "reliable": reliable})
    covered = {c["object"] for c in captions if c["object"] is not None}
    for p in ctx.paragraphs:
        if p._p.xpath(".//w:drawing|.//w:pict") and p._p not in covered and not p._p.xpath(".//m:oMath|.//w:object"):
            incomplete["figure"] = True
    if any(t not in covered for t in ctx.document.element.body.iter(qn("w:tbl"))):
        incomplete["table"] = True
    for equation in equations:
        p, number_p = equation["paragraph"], equation["number_paragraph"]
        elements = {p._p}
        if number_p is not None:
            elements.add(number_p._p)
        consumed.update(elements)
        reliable = bool(equation["number"] and not any(equation[k] for k in ("dynamic", "revision", "number_unknown", "opaque")))
        incomplete["equation"] |= not reliable
        records.append({"kind": "equation", "number": _key(equation["number"]["value"]) if equation["number"] else None,
                        "paragraph": number_p if number_p is not None else p, "elements": elements, "reliable": reliable})
    for heading in headings:
        p = heading["paragraph"]
        consumed.add(p._p)
        if heading["role"] != "body":
            continue
        kind = "chapter" if heading["level"] == 1 else "section"
        reliable = bool(heading.get("number_known") and heading.get("number") is not None and not _revised(p._p))
        incomplete[kind] |= not reliable
        values = heading.get("number_values") or []
        # For 第n章/第n节 references retain the level's own number. Dotted
        # composite section numbers are not guessed from a bare 第n节.
        number = str(values[0]) if len(values) == 1 else _key(str(heading.get("number") or ""))
        records.append({"kind": kind, "number": number, "paragraph": p, "elements": {p._p}, "reliable": reliable})
    return records, incomplete, consumed, headings


def _bookmark_objects(ctx, records):
    active, starts, ends, targets = {}, defaultdict(list), defaultdict(list), defaultdict(set)
    elements = defaultdict(set)
    for index, record in enumerate(records):
        for element in record["elements"]:
            elements[element].add(index)
    for node in ctx.document.element.body.iter():
        if node.tag == qn("w:bookmarkStart"):
            name, ident = node.get(qn("w:name")), node.get(qn("w:id"))
            if name:
                starts[name].append(ident)
                active[ident] = name
        elif node.tag == qn("w:bookmarkEnd"):
            ident = node.get(qn("w:id"))
            ends[ident].append(node)
            active.pop(ident, None)
        elif node.tag in {qn("w:p"), qn("w:t"), qn("m:oMath")}:
            pnode = node if node.tag == qn("w:p") else next((a for a in node.iterancestors() if a.tag == qn("w:p")), None)
            for name in active.values():
                targets[name].update(elements.get(pnode, set()))
    return starts, ends, targets


def _skip_reference_paragraphs(ctx, headings):
    # Bibliography interpretation remains deferred. Only explicit structural
    # boundaries are used; no meaning or external-source authenticity is read.
    result, inside, level = set(), False, None
    by_element = {h["paragraph"]._p: h for h in headings}
    for p in ctx.paragraphs:
        h = by_element.get(p._p)
        if h and h["role"] == "references":
            inside, level = True, h["level"]
        elif inside and h and level is not None and h["level"] <= level:
            inside = False
        if inside:
            result.add(p._p)
    for f in scan_body_fields(ctx):
        if f["kind"] == "TOC":
            result.update(f["result_paragraph_elements"])
    return result


def _ref_kind(field):
    """Require a label in this REF cache or directly adjoining this field.

    A figure mention elsewhere in the paragraph does not make a date/author
    bookmark reference an object reference. No semantic intent is inferred.
    """
    names = {"图": "figure", "figure": "figure", "表": "table", "table": "table",
             "公式": "equation", "式": "equation", "equation": "equation", "eq": "equation",
             "chapter": "chapter", "section": "section"}
    cache = field["result_text"].strip()
    match = re.match(r"^(图|表|公式|式|Figure|Table|Equation|Eq\.?|Chapter|Section)\s*[（(]?\s*" + _NUMBER,
                     cache, re.I)
    if match:
        return names[match[1].rstrip(".").casefold()]
    chapter = re.match(r"^第[0-9０-９零〇一二两三四五六七八九十百千.．]+([章节])", cache)
    if chapter:
        return "chapter" if chapter[1] == "章" else "section"
    leaves = field["result_elements"]
    pnode = field["paragraph_element"]
    if not leaves or pnode is None or not re.fullmatch(r"[（(]?\s*" + _NUMBER + r"\s*[）)]?", cache, re.I):
        return None
    before, after, phase = "", "", "before"
    for node in pnode.iter():
        if node is leaves[0]:
            phase = "cache"
        if phase == "before" and node.tag == qn("w:t"):
            before += node.text or ""
        elif phase == "after" and node.tag == qn("w:t"):
            after += node.text or ""
        if node is leaves[-1]:
            phase = "after"
    prefix = re.search(r"(图|表|公式|式|Figure|Table|Equation|Eq\.?|Chapter|Section)\s*[（(]?\s*$", before, re.I)
    if prefix:
        return names[prefix[1].rstrip(".").casefold()]
    if re.search(r"第\s*$", before) and re.match(r"\s*[章节]", after):
        return "chapter" if after.lstrip().startswith("章") else "section"
    return None


def _check_references(ctx, mapping):
    records, incomplete, consumed, headings = _objects(ctx, mapping)
    findings, lookup = [], defaultdict(list)
    for record in records:
        if record["reliable"]:
            lookup[(record["kind"], record["number"])].append(record)
    starts, ends, targets = _bookmark_objects(ctx, records)
    skip = consumed | _skip_reference_paragraphs(ctx, headings)
    for p in ctx.paragraphs:
        if p._p in skip:
            continue
        text = _visible(p._p)
        if re.search(r"(?:引用自|引自|摘自|参考文献|文献\s*\[|示例|教程|例如|演示)", text):
            continue  # 参考文献引用判定：待开会讨论暂不实现。
        fields = scan_fields(p._p)
        ref_fields = [f for f in fields if f["kind"] == "REF"]
        if _revised(p._p) or fields.errors:
            if _LITERAL.search(text) or ref_fields:
                findings.append(_unknown("137", "引用段落含修订或未闭合域，不能唯一对应当前显示内容", p))
            continue
        for field in ref_fields:
            match = re.match(r'^REF\s+(?:"([^"]+)"|([^\s\\]+))', field["instruction"], re.I)
            name = (match[1] or match[2]) if match else None
            linked = targets.get(name, set())
            reference_kind = _ref_kind(field)
            if not linked and reference_kind is None:
                continue
            if field["malformed"] or field["in_textbox"] or not name or field["has_nested"] or not field["cache_present"]:
                findings.append(_unknown("137", "REF域结构、嵌套结果或书签名称无法可靠读取", p))
            elif name not in starts:
                findings.append(_finding(_reference_code(name), "REF指向不存在的对象书签“" + name + "”", p, "正文对象引用应指向存在的唯一书签"))
            elif len(starts[name]) != 1 or len(ends[starts[name][0]]) != 1 or len(linked) != 1:
                findings.append(_unknown("137", "REF目标书签重复、未闭合或不唯一对应图/表/公式/章节对象", p))
            else:
                record = records[next(iter(linked))]
                rest = re.sub(r"\\h\b|\\\*\s+(?:MERGEFORMAT|CHARFORMAT)\b", "", field["instruction"][match.end():], flags=re.I).strip()
                target_text = _bookmark_text(ctx, name)
                if not record["reliable"] or rest or field["dirty"] or field["locked"] or target_text is None:
                    findings.append(_unknown("137", "目标编号含动态域、位置/编号开关或对象范围不可完整读取；不以保存缓存推算实时编号", p))
                elif reference_kind is not None and reference_kind != record["kind"]:
                    findings.append(_finding(_reference_code(record["kind"]), "引用标签类型为“" + reference_kind + "”，REF指向“" + record["kind"] + "”对象", p,
                                             "明确引用标签应与其唯一目标对象类型一致"))
                elif target_text != field["result_text"]:
                    findings.append(_finding(_reference_code(record["kind"]), "REF保存文字为“" + field["result_text"] + "”；未更新域", p,
                                             "唯一对象书签当前静态文字为“" + target_text + "”"))
        # Field result text is never also treated as an independently typed
        # literal reference. Mixed field/manual paragraphs remain conservative.
        if fields or _dynamic(p):
            if not ref_fields and _LITERAL.search(text):
                findings.append(_unknown("137", "引用文字包含动态或嵌入对象，不能当作静态编号比对", p))
            continue
        for match in _LITERAL.finditer(text):
            suffix = text[match.end():]
            if re.match(r"\s*(?:[-—~～至、,，和及]|到)\s*(?:图|表|公式|式)?\s*[（(]?[0-9０-９]", suffix) or re.match(r"\s*[（(][a-zA-Z][)）]", suffix):
                findings.append(_unknown("137", "范围、多编号或子编号引用不能唯一拆分目标", p))
                continue
            if match["scope"] != "本文":
                findings.append(_unknown("137", "本章/本节引用的局部重置作用域暂不能唯一确认", p))
                continue
            if match["chapter"]:
                from thesis_format_checker.detectors.headings import _token_value
                value, _ = _token_value(match["chapter"])
                kind = "chapter" if match["unit"] == "章" else "section"
                number = _key(match["chapter"]) if re.search(r"[.．]", match["chapter"]) else str(value)
                if kind == "section" and "." not in number and any(r["kind"] == kind and "." in (r["number"] or "") for r in records):
                    findings.append(_unknown("137", "裸节号与文档的多级节号作用域不能唯一对应", p))
                    continue
            else:
                kind = {"图": "figure", "表": "table", "公式": "equation", "式": "equation"}[match["kind"]]
                number = _key(match["num"])
            candidates = lookup.get((kind, number), [])
            possible = [r for r in records if r["kind"] == kind and r["number"] == number]
            if len(candidates) == 1 and len(possible) == 1:
                continue
            if candidates or incomplete[kind]:
                findings.append(_unknown("137", "引用“" + match[0] + "”存在同号多对象或未读/动态目标，不能断言目标缺失", p))
            else:
                findings.append(_finding(_reference_code(match[0]), "引用“" + match[0] + "”，未找到对应可确认对象编号", p,
                                         "明确的本文对象编号应对应文档内实际对象"))
    return findings


def check_consistency_errors(baseline, ctx, mapping=None):
    findings = []
    base = baseline.get("consistency", {})
    # A malformed story/object must not suppress the other internal checks.
    for rule, check in (("134", lambda: _check_titles(base, ctx)),
                        ("136", lambda: _check_directory(ctx, mapping)),
                        ("137", lambda: _check_references(ctx, mapping))):
        try:
            findings.extend(check())
        except Exception as exc:
            findings.append(_unknown(rule, "内部对象读取失败：" + type(exc).__name__, _fallback(ctx)))
    return findings
