"""Read-only note parts and non-bibliographic cross-reference checks.

100、101、102、109：本条依赖大模型语义，本阶段暂不实现。
103、104：参考文献引用判定【待开会讨论，本版本暂不开发】。
No note part, field cache, bookmark or source character is rewritten here.
"""
from __future__ import annotations

from copy import deepcopy
import re

from docx.oxml import parse_xml
from thesis_format_checker.core.xmlutil import qn
from docx.text.paragraph import Paragraph
from lxml import etree

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, _styles, effective_paragraph, make_finding, unknown
from thesis_format_checker.detectors.fields import scan_fields, scan_body_fields
from thesis_format_checker.detectors.headings import _heading_level, _toc_paragraph_ids
from thesis_format_checker.detectors.text_runs import effective_text_runs
from thesis_format_checker.detectors.bibliography import _heading_text

_TITLES = {"103": "正文引用序号格式不统一", "104": "正文引用序号不连续或顺序错误",
           "105": "注释类型不符合模板", "106": "脚注编号设置不符合模板",
           "107": "脚注文字或分隔线格式不符合模板", "108": "交叉引用机制或缓存不一致",
           "110": "链接或交叉引用显示域错误"}
_REVISION = {qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo"),
             qn("w:rPrChange"), qn("w:pPrChange"), qn("w:sectPrChange")}
_LABEL = re.compile(r"(?:图|表|式|Figure|Table|Equation)\s*[（(]?\d+(?:[.\-．－]\d+)*[）)]?|第\s*\d+\s*[章节]", re.I)
_ERROR = re.compile(r"错误\s*[!！:]\s*(?:未找到引用源|未定义书签|找不到引用源|找不到链接源|无法打开文件|链接无效)|"
                    r"Error\s*[!:]\s*(?:Reference source not found|Bookmark not defined|Not a valid link|Cannot open file|Not a valid filename)", re.I)
_FIELDS = {"REF", "PAGEREF", "NOTEREF", "HYPERLINK", "INCLUDETEXT", "INCLUDEPICTURE", "LINK"}
_PARA_KEYS = ("first_line_indent", "left_indent", "right_indent", "line_spacing", "line_spacing_rule")


def _fallback(ctx):
    return ctx.paragraphs[0] if ctx.paragraphs else None


def _uncertain(rule, detail, anchor=None):
    return unknown(rule, _title(rule), UNREADABLE + "；" + detail, anchor)


#: 脚注段落属性 → 细粒度规则编号
_FOOTNOTE_FORMAT_CODES = {
    "line_spacing": "107.3", "line_spacing_rule": "107.3",
    "space_before": "107.3", "space_after": "107.3",
    "first_line_indent": "107.4", "left_indent": "107.4", "right_indent": "107.4",
    "alignment": "107.4",
    "markers": "107.5", "borders": "107.5",
}


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else _TITLES[code]


def _footnote_format_code(key: str, paragraph_keys) -> str:
    """脚注格式差异键 → 细粒度编号（字体组归 107.1，其余按属性分派）。"""

    if key in _FOOTNOTE_FORMAT_CODES:
        return _FOOTNOTE_FORMAT_CODES[key]
    if key in paragraph_keys:
        return "107.3"
    return "107.1"


def _has_revision(element):
    return any(n.tag in _REVISION for n in element.iter()) or any(a.tag in _REVISION for a in element.iterancestors())


def _paragraph(ctx, element):
    return next((p for p in ctx.paragraphs if p._p is element), _fallback(ctx))


def _note_data(ctx):
    records, errors, refs = [], [], []
    for p in ctx.paragraphs:
        for kind in ("footnote", "endnote"):
            for node in p._p.iter(qn("w:" + kind + "Reference")):
                if any(a.tag in {qn("w:del"), qn("w:moveFrom")} for a in node.iterancestors()):
                    continue
                refs.append({"kind": kind, "id": node.get(qn("w:id")), "anchor": p,
                             "custom": node.get(qn("w:customMarkFollows"), "0") not in ("0", "false", "off"),
                             "revision": _has_revision(p._p)})
    for kind in ("footnote", "endnote"):
        if any(any(a.tag == qn("w:txbxContent") for a in node.iterancestors())
               for node in ctx.document.element.body.iter(qn("w:" + kind + "Reference"))):
            errors.append((kind, "文本框中的注释标记不能可靠关联正文批注位置"))
    for kind in ("footnote", "endnote"):
        related = [r for r in ctx.document.part.rels.values() if r.reltype.endswith("/" + kind + "s")]
        if len(related) != 1:
            if any(r["kind"] == kind for r in refs):
                errors.append((kind, "注释部件关系缺失或重复"))
            continue
        try:
            relation = related[0]
            if relation.is_external:
                raise ValueError("external note part")
            # Parse into a detached tree; the package's original bytes stay intact.
            blob = relation.target_part.blob
            if b"<!DOCTYPE" in blob.upper() or b"<!ENTITY" in blob.upper():
                raise ValueError("unsafe XML declaration")
            root = parse_xml(blob)
            if root.tag != qn("w:" + kind + "s"):
                raise ValueError("unexpected note root")
            for element in root:
                if element.tag != qn("w:" + kind):
                    continue
                note_id = element.get(qn("w:id"))
                matches = [r for r in refs if r["kind"] == kind and r["id"] == note_id]
                note_type = element.get(qn("w:type"), "normal")
                if note_type == "normal" and not matches:
                    continue
                records.append({"kind": kind, "id": note_id, "type": note_type,
                                "element": element,
                                "anchor": matches[0]["anchor"] if matches else next((r["anchor"] for r in refs if r["kind"] == kind), _fallback(ctx))})
        except Exception as exc:
            errors.append((kind, "注释部件XML读取失败：" + type(exc).__name__))
    for ref in refs:
        matches = [r for r in records if r["kind"] == ref["kind"] and r["id"] == ref["id"] and r["type"] == "normal"]
        if len(matches) != 1:
            errors.append((ref["kind"], "正文注释标记无法唯一关联注释内容"))
    return {"records": records, "errors": list(dict.fromkeys(errors)), "refs": refs}


def _number_config(ctx, section):
    # These are OOXML absence semantics, not a school format standard.
    values = {"numFmt": "decimal", "numStart": "1", "numRestart": "continuous"}
    roots = [ctx.document.settings.element.find(qn("w:footnotePr"))]
    sections = list(ctx.document.sections)
    if section < len(sections):
        if _has_revision(sections[section]._sectPr):
            raise ValueError("revised section settings")
        roots.append(sections[section]._sectPr.find(qn("w:footnotePr")))
    for root in roots:
        if root is None:
            continue
        if _has_revision(root):
            raise ValueError("revised note settings")
        for key in values:
            children = root.findall(qn("w:" + key))
            if len(children) > 1:
                raise ValueError("duplicate note numbering property")
            if children:
                value = children[0].get(qn("w:val"))
                if not value:
                    raise ValueError("missing note numbering value")
                values[key] = value
                if key == "numFmt" and value == "custom":
                    values["customFormat"] = children[0].get(qn("w:format"))
    if not values["numStart"].isdigit() or int(values["numStart"]) < 1:
        raise ValueError("invalid note numbering start")
    if values["numRestart"] not in {"continuous", "eachSect", "eachPage"}:
        raise ValueError("unknown note numbering restart")
    return values


def _font_keys(text):
    keys = ["size_pt"]
    if re.search(r"[\u3400-\u9fff]", text):
        keys.append("east_asia_font")
    if re.search(r"[A-Za-z0-9]", text):
        keys.append("ascii_font")
    if re.search(r"[^\x00-\x7f\u3400-\u9fff\u3000-\u303f\uff00-\uffef\s]", text):
        keys.append("hansi_font")
    return keys


def _note_profiles(ctx, note):
    result = []
    for index, element in enumerate(note["element"].iter(qn("w:p"))):
        p = Paragraph(deepcopy(element), ctx.document._body)
        if _has_revision(element):
            raise ValueError("注释包含未决修订")
        if element.xpath(".//w:object | .//w:drawing | .//w:txbxContent | .//w:fldChar | .//w:fldSimple"):
            raise ValueError("注释含对象或域，无法确认文字实际格式")
        if any(a.tag == qn("w:tbl") for a in element.iterancestors()):
            raise ValueError("注释位于表格中，表格条件样式不能可靠叠加")
        if any(prop is not None and prop.find(qn("w:numPr")) is not None for prop in
               [p._p.pPr] + [s.element.pPr for s in _styles(p.style)]):
            raise ValueError("注释使用自动列表，编号级别缩进不能可靠叠加")
        role = "first" if index == 0 else "continuation"
        profile = {"role": role, "paragraph": {k: effective_paragraph(p)[k] for k in _PARA_KEYS}, "fonts": {}}
        for text, fmt in effective_text_runs(p):
            if not text.strip():
                continue
            if fmt["run"]._r.xpath(".//w:footnoteRef | .//w:endnoteRef"):
                raise ValueError("注释编号标记与正文文字混于同一run")
            for key in _font_keys(text):
                group = (fmt["character_style"], key)
                profile["fonts"].setdefault(group, set()).add(fmt[key])
        # Separator line geometry rendered by Word is not inferred from pixels.
        if note["type"] != "normal":
            profile["markers"] = tuple(etree.QName(n).localname for n in element.iter() if n.tag in {qn("w:separator"), qn("w:continuationSeparator")})
            borders = element.find("./" + qn("w:pPr") + "/" + qn("w:pBdr"))
            if borders is not None and any(any("theme" in etree.QName(k).localname.lower() for k in n.attrib) for n in borders):
                raise ValueError("分隔线边框使用主题颜色，无法确认实际外观")
            profile["borders"] = tuple((etree.QName(n).localname, tuple(sorted((etree.QName(k).localname, v) for k, v in n.attrib.items()))) for n in borders) if borders is not None else ()
        result.append(profile)
    return result


def _field_label(field):
    cache = field["result_text"]
    match = _LABEL.fullmatch(cache.strip())
    if match:
        return match.group()
    # Word commonly stores only the number in REF and the label before it.
    p = field["paragraph_element"]
    leaves = field["result_elements"]
    if p is None or not leaves:
        return None
    before = ""
    for node in p.iter():
        if node is leaves[0]:
            break
        if node.tag == qn("w:t"):
            before += node.text or ""
    prefix = re.search(r"(?:图|表|式|Figure|Table|Equation)\s*[（(]?\s*$", before, re.I)
    if prefix:
        candidate = prefix.group() + cache
        return candidate if _LABEL.fullmatch(candidate.strip()) else None
    return None


def _pattern(label):
    return re.sub(r"\d+", lambda _: r"\d+", re.escape(label))


def extract_notes_baseline(ctx, mapping=None):
    data = _note_data(ctx)
    baseline = {"types": sorted({r["kind"] for r in data["refs"] if not r["revision"]}),
                "note_errors": data["errors"], "types_unreadable": bool(data["errors"]) or any(r["revision"] for r in data["refs"]),
                "numbering": [], "profiles": {}, "profile_errors": {}, "ref_patterns": [],
                "manual_reference_required": bool(re.search(r"(?:必须|应|须)[^。\n]{0,16}(?:使用|采用)[^。\n]{0,8}(?:自动交叉引用|交叉引用域|REF域)|"
                                                            r"(?:交叉引用|图表引用|公式引用)[^。\n]{0,16}(?:必须|应|须)[^。\n]{0,8}(?:自动|使用域)", ctx.text, re.I))}
    for index in range(len(ctx.document.sections)):
        try:
            baseline["numbering"].append(_number_config(ctx, index))
        except Exception as exc:
            baseline["numbering"].append({"unreadable": type(exc).__name__})
    for note in data["records"]:
        if note["kind"] != "footnote":
            continue
        try:
            baseline["profiles"].setdefault(note["type"], []).extend(_note_profiles(ctx, note))
        except Exception as exc:
            baseline["profile_errors"][note["type"]] = str(exc)
    fields = scan_body_fields(ctx)
    # A malformed or revised template field must not silently disappear while
    # the remaining samples are treated as a complete mandatory baseline.
    baseline["ref_patterns_unreadable"] = bool(fields.errors) or (
        _has_revision(ctx.document.element.body) and bool(ctx.document.element.body.xpath(".//w:fldSimple | .//w:instrText")))
    for field in fields:
        if field["kind"] != "REF":
            continue
        if (field["malformed"] or field["in_textbox"] or field["has_nested"] or not field["cache_present"] or
                field["locked"] or field["dirty"] or _has_revision(field["element"])):
            baseline["ref_patterns_unreadable"] = True
            continue
        label = _field_label(field)
        if label:
            baseline["ref_patterns"].append(_pattern(label))
    baseline["ref_patterns"] = sorted(set(baseline["ref_patterns"]))
    return {"notes": baseline}


def _check_notes(base, ctx):
    data, findings = _note_data(ctx), []
    if not data["refs"] and not data["errors"]:
        return findings
    for ref in data["refs"]:
        if ref["revision"]:
            findings.append(_uncertain("105", "注释标记含未决修订", ref["anchor"]))
        elif base.get("types_unreadable"):
            findings.append(_uncertain("105", "模板注释类型示例包含读取异常或未决修订", ref["anchor"]))
        elif not base.get("types"):
            findings.append(_uncertain("105", "模板没有可确认的注释类型示例", ref["anchor"]))
        elif ref["kind"] not in base["types"]:
            # 模板只允许一种注释类型却出现另一种 → 脚注尾注混用（105.1）
            code = "105.1" if len(base["types"]) == 1 else "105.2"
            findings.append(make_finding(code, _title(code), "允许类型：" + "、".join(base["types"]), ref["kind"], ref["anchor"]))
    for kind, reason in data["errors"]:
        anchor = next((r["anchor"] for r in data["refs"] if r["kind"] == kind), _fallback(ctx))
        findings.append(_uncertain("105", reason, anchor))
        if kind == "footnote":
            findings.append(_uncertain("107", reason, anchor))
    checked_sections = set()
    for ref in (r for r in data["refs"] if r["kind"] == "footnote"):
        section = ctx.section_index(ref["anchor"])
        if ref["custom"] or ref["revision"]:
            findings.append(_uncertain("106", "自定义注释标记或未决修订，不能由内部ID推算显示编号", ref["anchor"]))
        if section in checked_sections:
            continue
        checked_sections.add(section)
        expected_all = base.get("numbering", [])
        if base.get("types_unreadable") or "footnote" not in base.get("types", []):
            findings.append(_uncertain("106", "模板没有脚注编号示例", ref["anchor"]))
            continue
        if expected_all and all(item == expected_all[0] for item in expected_all):
            expected = expected_all[0]
        else:
            findings.append(_uncertain("106", "模板和论文分节无法唯一对应脚注编号设置", ref["anchor"]))
            continue
        try:
            actual = _number_config(ctx, section)
            if expected.get("unreadable"):
                raise ValueError("模板编号设置不可读")
            if expected != actual:
                findings.append(make_finding("106.3", _title("106.3"), str(expected), str(actual), ref["anchor"]))
            if actual["numRestart"] == "eachPage" or actual["numFmt"] == "custom":
                findings.append(_uncertain("106.3", "已读取编号设置，但按页重启/自定义显示编号需要Word分页渲染；未将注释ID当作编号", ref["anchor"]))
        except Exception as exc:
            findings.append(_uncertain("106", "脚注编号设置无法读取：" + str(exc), ref["anchor"]))
    footnotes = [n for n in data["records"] if n["kind"] == "footnote"]
    for note in footnotes:
        try:
            if any(kind == "footnote" for kind, _ in base.get("note_errors", [])):
                raise ValueError("模板脚注部件读取不完整")
            if base.get("profile_errors", {}).get(note["type"]):
                raise ValueError("模板注释格式含未决对象或修订")
            expected_profiles = base.get("profiles", {}).get(note["type"], [])
            actual_profiles = _note_profiles(ctx, note)
            if not expected_profiles:
                raise ValueError("模板没有对应的脚注文字/分隔线格式示例")
            differences, limits = [], []
            for profile in actual_profiles:
                candidates = [p for p in expected_profiles if p["role"] == profile["role"]]
                if not candidates:
                    limits.append("缺少相应脚注段落角色示例")
                    continue
                for key, actual in profile["paragraph"].items():
                    values = {p["paragraph"][key] for p in candidates}
                    if len(values) != 1 or None in values or actual is None:
                        limits.append(key + "有效值不唯一或不可读")
                    elif actual not in values:
                        differences.append((key, next(iter(values)), actual))
                for group, actuals in profile["fonts"].items():
                    values = set().union(*(p["fonts"].get(group, set()) for p in candidates))
                    if any(group not in p["fonts"] for p in candidates) or len(values) != 1 or None in values or None in actuals:
                        limits.append(str(group) + "字体字号基准不唯一或不可读")
                    elif actuals != values:
                        differences.append((str(group), next(iter(values)), sorted(str(v) for v in actuals)))
                for key in ("markers", "borders"):
                    if key not in profile:
                        continue
                    values = {p[key] for p in candidates}
                    if len(values) != 1:
                        limits.append("分隔线基准不唯一")
                    elif profile[key] not in values:
                        differences.append((key, next(iter(values)), profile[key]))
            if differences:
                # 字体 / 行距 / 缩进 / 分隔线 各自成条
                grouped: dict[str, list] = {}
                for diff in differences:
                    grouped.setdefault(_footnote_format_code(str(diff[0]), _PARA_KEYS), []).append(diff)
                for code, items in grouped.items():
                    findings.append(make_finding(code, _title(code),
                                                 "；".join(str(k) + "=" + str(e) for k, e, _ in items),
                                                 "；".join(str(k) + "=" + str(a) for k, _, a in items),
                                                 note["anchor"]))
            if limits:
                findings.append(_uncertain("107", "；".join(sorted(set(limits))), note["anchor"]))
        except Exception as exc:
            findings.append(_uncertain("107", "脚注文字/分隔线读取受限：" + str(exc), note["anchor"]))
    normal = next((n for n in footnotes if n["type"] == "normal"), None)
    if normal and not any(n["type"] == "separator" for n in footnotes):
        findings.append(_uncertain("107", "脚注分隔线未在部件中显式保存，无法确认Word内置分隔线外观", normal["anchor"]))
    return findings


def _bookmark_text(ctx, name):
    starts = [n for n in ctx.document.element.body.iter(qn("w:bookmarkStart")) if n.get(qn("w:name")) == name]
    if len(starts) != 1:
        return None
    start = starts[0]
    start_p = next((a for a in start.iterancestors() if a.tag == qn("w:p")), None)
    if start_p is None:
        return None
    paragraph = Paragraph(start_p, ctx.document._body)
    if _has_revision(start_p) or any(prop is not None and prop.find(qn("w:numPr")) is not None for prop in
                                   [start_p.pPr] + [s.element.pPr for s in _styles(paragraph.style)]):
        return None
    note_id = start.get(qn("w:id"))
    ends = [n for n in ctx.document.element.body.iter(qn("w:bookmarkEnd")) if n.get(qn("w:id")) == note_id]
    if len(ends) != 1:
        return None
    leaves, active = [], False
    for node in ctx.document.element.body.iter():
        if node is start:
            active = True
            continue
        if node is ends[0]:
            return "".join(leaves) if active else None
        if not active:
            continue
        if node.tag == qn("w:p"):
            # Cross-paragraph bookmarks can include inherited numbering and
            # paragraph marks; a renderer is needed for an exact REF result.
            return None
        if _has_revision(node) or node.tag in {qn("w:fldChar"), qn("w:fldSimple"), qn("w:object"), qn("w:drawing"), qn("w:numPr"), qn("m:oMath")}:
            return None
        if node.tag == qn("w:t"):
            leaves.append(node.text or "")
        elif node.tag in {qn("w:tab"), qn("w:br"), qn("w:cr")}:
            return None
    return None


def _check_fields(base, ctx):
    findings, fields = [], []
    if base.get("ref_patterns_unreadable"):
        findings.append(_uncertain("108", "模板交叉引用示例包含不可读域或未决修订，不能据剩余示例判定手工引用错误", _fallback(ctx)))
    stories = [(ctx.document.element.body, None, "正文")]
    notes = _note_data(ctx)
    stories.extend((n["element"], n["anchor"], "注释") for n in notes["records"] if n["type"] == "normal")
    for relation in ctx.document.part.rels.values():
        if not relation.reltype.endswith(("/header", "/footer")):
            continue
        try:
            if relation.is_external:
                raise ValueError("external header/footer")
            blob = relation.target_part.blob
            if b"<!DOCTYPE" in blob.upper() or b"<!ENTITY" in blob.upper():
                raise ValueError("unsafe XML declaration")
            stories.append((parse_xml(blob), _fallback(ctx), "页眉页脚"))
        except Exception as exc:
            findings.append(_uncertain("110", "页眉页脚故事域不可读取：" + type(exc).__name__, _fallback(ctx)))
    for root, story_anchor, story_name in stories:
        scanned = scan_fields(root)
        for item in scanned:
            item["story_anchor"] = story_anchor
        fields.extend(scanned)
        for reason in scanned.errors:
            findings.append(_uncertain("110", story_name + "域结构解析失败：" + reason, story_anchor or _fallback(ctx)))
    for field in fields:
        if field["kind"] not in _FIELDS:
            continue
        anchor = field["story_anchor"] or _paragraph(ctx, field["paragraph_element"])
        if field["in_textbox"] or field["malformed"] or _has_revision(field["element"]):
            findings.append(_uncertain("110", "域结构、文本框或修订阻止可靠解析", anchor))
            if field["kind"] in {"REF", "PAGEREF", "NOTEREF"}:
                findings.append(_uncertain("108", "交叉引用域结构不能可靠解析", anchor))
            continue
        if field["has_nested"] or not field["cache_present"]:
            findings.append(_uncertain("110", "嵌套域或没有保存结果缓存，不能确认显示内容", anchor))
            if field["kind"] in {"REF", "PAGEREF", "NOTEREF"}:
                findings.append(_uncertain("108", "交叉引用缓存无法确认", anchor))
            continue
        if _ERROR.search(field["result_text"]):
            findings.append(make_finding("110", _title("110"), "域可读取且不显示Word域错误结果", "已保存域结果：" + field["result_text"], anchor))
            continue
        if field["kind"] != "REF":
            # No page renderer and no footnote numbering evaluator are used.
            if field["kind"] in {"PAGEREF", "NOTEREF"}:
                findings.append(_uncertain("108", "PAGEREF/NOTEREF显示值依赖分页或注释编号；仅读取缓存", anchor))
            continue
        label = _field_label(field)
        if not label:
            # Ambiguous and bibliography REF fields are deliberately deferred.
            continue
        instruction = field["instruction"]
        match = re.match(r'^REF\s+(?:"([^\"]+)"|([^\s\\]+))', instruction, re.I)
        if not match:
            findings.append(_uncertain("108", "REF书签指令无法解析", anchor))
            continue
        rest = instruction[match.end():]
        # Only plain textual REF plus hyperlink/MERGEFORMAT switches has a
        # dependable text comparison; numbering/position switches need Word.
        rest = re.sub(r"\\h\b|\\\*\s+(?:MERGEFORMAT|CHARFORMAT)\b", "", rest, flags=re.I).strip()
        if rest or field["locked"] or field["dirty"]:
            findings.append(_uncertain("108", "REF编号/位置开关、锁定或待更新状态，不能据缓存断言当前显示值", anchor))
            continue
        target = _bookmark_text(ctx, match.group(1) or match.group(2))
        if target is None:
            findings.append(_uncertain("108", "REF目标书签缺失、跨对象或含动态内容，无法确认缓存同步", anchor))
        elif target != field["result_text"]:
            findings.append(make_finding("108.2", _title("108.2"), "可读取的目标书签文字：" + target,
                                         "已保存REF结果：" + field["result_text"] + "；尚未刷新域，不能断言修改历史", anchor))
    if base.get("ref_patterns"):
        field_leaves = {id(n) for field in fields for n in field["result_elements"]}
        in_bibliography, bibliography_level = False, None
        toc = _toc_paragraph_ids(ctx)
        for p in ctx.paragraphs:
            if p._p in toc:
                continue
            text = p.text.strip()
            compact = _heading_text(text)
            level = _heading_level(p)
            if compact in {"参考文献", "references", "bibliography"}:
                in_bibliography, bibliography_level = True, level
                continue
            back_heading = re.match(r"^(?:致谢|附录|攻读学位期间|acknowledg(?:e)?ments|appendix)(?:$|[a-z一二三四五六七八九十\d：:])", compact)
            # Bibliography subheadings remain protected. If its heading level
            # is unknown, only an explicit following back-matter title closes it.
            if in_bibliography and (back_heading or (level and bibliography_level is not None and level <= bibliography_level)):
                in_bibliography = False
            if in_bibliography or p.style.name.lower().startswith(("caption", "题注")) or re.search(r"(?:示例|例如|演示|教程|域代码|显示错误|输入文字|手工输入)", text):
                continue
            if _has_revision(p._p):
                continue
            fragments = []
            for node in p._p.iter(qn("w:t")):
                ancestors = list(node.iterancestors())
                if id(node) in field_leaves or any(a.tag in {qn("w:hyperlink"), qn("w:txbxContent")} for a in ancestors):
                    fragments.append("\x00")
                else:
                    fragments.append(node.text or "")
            plain = "".join(fragments)
            for pattern in base["ref_patterns"]:
                found = re.search(r"(?:参见|参照|见|如)\s*(" + pattern + r")(?!\d)", plain, re.I)
                if found:
                    if base.get("manual_reference_required") and not base.get("ref_patterns_unreadable"):
                        findings.append(make_finding("108.1", _title("108.1"), "模板明确要求使用自动交叉引用，且同类编号示例为REF域",
                                                     "明确引用文字“" + found.group(1) + "”未位于交叉引用域中；仅判断手工引用，未断言编号历史", p))
                    else:
                        reason = "模板交叉引用基准不完整" if base.get("ref_patterns_unreadable") else "模板同类编号使用REF域，但未明确排除手工引用"
                        findings.append(_uncertain("108", reason + "；正文“" + found.group(1) + "”为普通文字，需人工确认", p))
                    break
    return findings


# ---------------------------------------------------------------------------
# 103/104：正文引用序号 [1] / [1,2] / [1-3] 的格式与顺序
# ---------------------------------------------------------------------------

_CITATION = re.compile(r"([\[［])(\d+(?:\s*[,，\-–—~～]\s*\d+)*)([\]］])")


def _citation_numbers(match) -> list[int]:
    """展开引用序号；``[34-37]`` 表示 34、35、36、37，不只是两个端点。"""

    values: list[int] = []
    for item in re.split(r"\s*[,，]\s*", match.group(2)):
        range_match = re.fullmatch(r"(\d+)\s*[\-–—~～]\s*(\d+)", item)
        if range_match:
            start, end = map(int, range_match.groups())
            if start <= end and end - start <= 1000:
                values.extend(range(start, end + 1))
            else:
                values.extend((start, end))
        else:
            values.extend(int(value) for value in re.findall(r"\d+", item))
    return values


def _citation_format(paragraph, start: int, end: int) -> tuple[bool | None, tuple]:
    """读取一个引用标记自身的上标、字体和字号，不把整段格式混在一起。"""

    from thesis_format_checker.detectors.text_runs import effective_text_runs

    superscripts, profiles = set(), set()
    offset = 0
    for text, props in effective_text_runs(paragraph):
        stop = offset + len(text)
        if stop > start and offset < end:
            superscripts.add(props.get("vertical_align") == "superscript")
            profiles.add((props.get("ascii_font"), props.get("size_pt")))
        offset = stop
        if offset >= end:
            break
    superscript = next(iter(superscripts)) if len(superscripts) == 1 else None
    return superscript, tuple(sorted(profiles, key=str))


def _marker_finding(rule, expected, actual, marker):
    finding = make_finding(
        rule, _title(rule), expected, actual, marker["record"]["paragraph"])
    finding.start = marker["match"].start()
    finding.end = marker["match"].end()
    finding.excerpt = marker["match"].group(0)
    return finding


def _outliers(markers, key):
    """返回偏离主要写法的标记；并列时以文中最先出现的写法为基准。"""

    from collections import Counter

    counts = Counter(marker[key] for marker in markers if marker[key] is not None)
    if len(counts) <= 1:
        return [], None
    order = {value: next(i for i, marker in enumerate(markers)
                         if marker[key] == value) for value in counts}
    expected = min(counts, key=lambda value: (-counts[value], order[value]))
    return [marker for marker in markers if marker[key] not in (None, expected)], expected


def _missing_groups(values: list[int]) -> list[list[int]]:
    groups: list[list[int]] = []
    for value in values:
        if not groups or value != groups[-1][-1] + 1:
            groups.append([value])
        else:
            groups[-1].append(value)
    return groups


def _citation_records(ctx):
    """收集正文里的引用序号；目录与参考文献表本身不参与。"""

    toc = _toc_paragraph_ids(ctx)
    result, in_references = [], False
    for index, paragraph in enumerate(ctx.paragraphs, 1):
        raw_text = paragraph.text
        text = raw_text.strip()
        compact = re.sub(r"\s+", "", text).casefold()
        if compact in {"参考文献", "references", "bibliography"}:
            in_references = True
            continue
        if in_references:
            if _heading_level(paragraph) and text:
                in_references = False
            else:
                continue
        if not text or paragraph._p in toc:
            continue
        if paragraph._p.xpath(".//w:instrText | .//w:fldSimple | .//w:fldChar"):
            continue  # 域缓存里的编号不能当作正文引用序号
        # 编号型参考文献从 1 起；含 0 的 ``[0,1]``、``[0,100]`` 是数值区间，
        # 不能拿来制造“非上标引用”和大段缺号。
        matches = [match for match in _CITATION.finditer(raw_text)
                   if 0 not in _citation_numbers(match)]
        if matches:
            result.append({
                "paragraph": paragraph, "index": index,
                "matches": matches, "text": raw_text,
            })
    return result


def _check_citations(baseline, ctx):
    """103.1–103.3 引用写法统一性；104.1–104.3 引用序号连续性/顺序。"""

    records = _citation_records(ctx)
    if not records:
        return []
    findings = []
    markers, first_order, repeated = [], [], []
    for record in records:
        for match in record["matches"]:
            numbers = _citation_numbers(match)
            superscript, profile = _citation_format(
                record["paragraph"], match.start(), match.end())
            marker = {
                "record": record, "match": match, "numbers": numbers,
                "brackets": match.group(1) + match.group(3),
                "superscript": superscript, "profile": profile,
            }
            markers.append(marker)
            first_order.append((numbers[0], marker))
            if len({value for value in numbers if numbers.count(value) > 1}):
                repeated.append(marker)
    outliers, expected_brackets = _outliers(markers, "brackets")
    for marker in outliers:
        findings.append(_marker_finding(
            "103.1", f"引用序号括号形式应全文统一（主要使用 {expected_brackets}）",
            f"“{marker['match'].group(0)}”使用了 {marker['brackets']}", marker))
    outliers, expected_superscript = _outliers(markers, "superscript")
    for marker in outliers:
        expected_label = "上标" if expected_superscript else "非上标"
        actual_label = "上标" if marker["superscript"] else "非上标"
        findings.append(_marker_finding(
            "103.2", f"引用序号是否上标应全文统一（主要为{expected_label}）",
            f"“{marker['match'].group(0)}”为{actual_label}", marker))
    outliers, expected_profile = _outliers(markers, "profile")
    for marker in outliers:
        findings.append(_marker_finding(
            "103.3", "引用序号字体字号应全文统一",
            f"“{marker['match'].group(0)}”的字体字号与主要写法不同", marker))
    numbers = sorted({value for marker in markers for value in marker["numbers"]})
    if numbers:
        missing = [value for value in range(1, numbers[-1] + 1) if value not in numbers]
        for group in _missing_groups(missing):
            later = [marker for marker in markers
                     if any(value > group[-1] for value in marker["numbers"])]
            if not later:
                continue
            marker = min(later, key=lambda item: min(
                value for value in item["numbers"] if value > group[-1]))
            findings.append(_marker_finding(
                "104.1", "正文引用序号应连续",
                f"在“{marker['match'].group(0)}”之前缺少序号："
                + "、".join(str(value) for value in group), marker))
    if repeated:
        for marker in repeated:
            findings.append(_marker_finding(
                "104.2", "同一个引用方括号内不应重复出现同一序号",
                f"“{marker['match'].group(0)}”重复标注了同一文献", marker))
    policies = ((baseline.get("bibliography") or {}).get("policies") or {}).get("sorting") or set()
    if "number" in policies:
        order, ordered_markers = [], []
        for value, marker in first_order:
            if value not in order:
                order.append(value)
                ordered_markers.append(marker)
        for position, value in enumerate(order[1:], 1):
            if value < order[position - 1]:
                marker = ordered_markers[position]
                findings.append(_marker_finding(
                    "104.3", "参考文献按正文首次引用顺序编号",
                    f"“{marker['match'].group(0)}”的首次出现序号小于此前的"
                    f"{order[position - 1]}", marker))
    return findings


def check_note_reference_errors(baseline, ctx, mapping=None):
    base = baseline.get("notes", {})
    findings = []
    # Isolation ensures an unreadable footnotes part cannot suppress field errors.
    for rules, checker in ((("105", "106", "107"), _check_notes), (("108", "110"), _check_fields)):
        try:
            findings.extend(checker(base, ctx))
        except Exception as exc:
            findings.extend(_uncertain(rule, "读取异常：" + type(exc).__name__, _fallback(ctx)) for rule in rules)
    try:
        # 103/104 是正文引用序号的写法与顺序检查，不需要注释部件基准。
        findings.extend(_check_citations(baseline, ctx))
    except Exception as exc:
        findings.append(_uncertain("103", "读取异常：" + type(exc).__name__, _fallback(ctx)))
    return findings
