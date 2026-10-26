"""Read-only caption and table checks, rules 078--084.

082, 085, 086, 087, 088: 本条依赖大模型语义，本阶段暂不实现。
No image quality, rendered pagination, reference intent or table-data semantics
are inferred.  Field caches never supply current computed sequence values.
"""
from __future__ import annotations

import re
from collections import defaultdict

from thesis_format_checker.core.xmlutil import qn, visible_text
from docx.table import Table
from docx.text.paragraph import Paragraph

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, effective_paragraph, make_finding, unknown, describe_value, memo
from thesis_format_checker.detectors.fields import scan_fields, scan_body_fields
from thesis_format_checker.detectors.headings import collect_heading_records


TITLES = {"078": "图表题注相对位置不符合模板", "079": "图表题注格式不符合模板",
          "080": "图表编号序列不符合模板", "081": "图表编号形式不符合模板",
          "083": "表格有效边框不符合模板", "084": "表头表身合计行格式不符合模板"}
LABELS = {"east_asia_font": "中文字体", "ascii_font": "西文字体", "size_pt": "字号",
          "bold": "加粗", "color": "文字颜色", "alignment": "段落对齐",
          "space_before": "段前间距", "space_after": "段后间距", "height": "行高及高度规则",
          "vertical": "单元格垂直对齐", "shading": "单元格底纹"}
#: 题注标签：中文、英文全称与常见缩写（Fig./Tab.）。
#: 顺序敏感——`Figure|Table` 必须排在 `Fig.|Tab.` 之前，否则会先匹配到缩写。
CAPTION = re.compile(r"^(图|表|Figure|Table|Fig\.?|Tab\.?)([ \t\u3000]*)"
                     r"([0-9０-９]+(?:[.．\-—－][0-9０-９]+)*)([ \t\u3000]*)(.*)$", re.I)
TOTAL = re.compile(r"^(?:合计|总计|小计|Total|Subtotal)(?:[：:]\s*)?$", re.I)

#: 题注标签 → 图/表；未知标签历史上按表处理，保持不变。
_LABEL_KINDS = {"图": "figure", "figure": "figure", "fig": "figure",
                "表": "table", "table": "table", "tab": "table"}


def _label_kind(label) -> str:
    return _LABEL_KINDS.get(str(label or "").casefold().rstrip("."), "table")


def _notice(rule, reason, paragraph=None):
    return unknown(rule, _title(rule), UNREADABLE + "；" + reason, paragraph)


#: 题注/表格属性 → 细粒度规则编号
_PROPERTY_CODES = {
    "079": {"east_asia_font": "079.1", "ascii_font": "079.1", "hansi_font": "079.1", "color": "079.1",
            "bold": "079.1", "size_pt": "079.2", "alignment": "079.3",
            "space_before": "079.4", "space_after": "079.4",
            "height": "079.4", "vertical": "079.4", "shading": "079.4"},
    "084": {"east_asia_font": "084.2", "ascii_font": "084.2", "hansi_font": "084.2", "color": "084.2",
            "bold": "084.2", "size_pt": "084.2", "alignment": "084.2",
            "space_before": "084.2", "space_after": "084.2",
            "height": "084.4", "shading": "084.5", "vertical": "084.2"},
}
#: 表格边框方向 → 细粒度规则编号（083.1 竖线 / 083.2 横线）
_BORDER_CODES = {"left": "083.1", "right": "083.1",
                 "top": "083.2", "bottom": "083.2", "tl2br": "083.2", "tr2bl": "083.2"}


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else TITLES[code]


def _visible(p):
    """元素内可见文字（跳过删除内容/文本框，一次遍历；见 core.xmlutil.visible_text）。"""

    return visible_text(p)


def _styles(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def _object_kind(node):
    if node.tag == qn("w:tbl"):
        return "table"
    if node.tag == qn("w:p") and node.xpath(".//wp:inline|.//wp:anchor|.//w:pict|.//w:object"):
        return "figure"
    return None


def _adjacent_objects(paragraph, kind):
    """Only same-container adjacent block associations, never visual geometry."""
    found = []
    for direction, attr in (("above", "getprevious"), ("below", "getnext")):
        node = getattr(paragraph._p, attr)()
        while node is not None:
            candidate = _object_kind(node)
            if candidate:
                if candidate == kind:
                    found.append((direction, node))
                break
            if node.tag != qn("w:p"):
                break
            text = _visible(node).strip()
            if node.xpath(".//w:br|.//w:sectPr|.//w:lastRenderedPageBreak"):
                break
            match = CAPTION.match(text)
            same_caption = bool(match) and (_label_kind(match[1]) == kind)
            if text and not same_caption:
                break
            node = getattr(node, attr)()
    return found


def _caption_role(ctx, paragraph, headings):
    earlier = [h for h in headings if ctx.index(h["paragraph"]) <= ctx.index(paragraph) and h["level"] == 1]
    heading = earlier[-1] if earlier else None
    role = heading["role"] if heading else "body"
    chapter = heading["number_values"][0] if heading and heading["number_known"] and heading["number_values"] else None
    return role, chapter


def _revised(element):
    return bool(element.xpath(".//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo|.//w:rPrChange|.//w:pPrChange|.//w:tblPrChange|.//w:trPrChange|.//w:tcPrChange") or
                any(a.tag in (qn("w:ins"), qn("w:moveTo")) for a in element.iterancestors()))


def _collect_caption_records_impl(ctx):
    """Return conservative caption records for use by paragraph detectors too."""
    records = []
    headings = collect_heading_records(ctx)
    fields = scan_body_fields(ctx)
    toc_nodes = {p for f in fields if f["kind"] == "TOC" for p in f["result_paragraph_elements"]}
    for paragraph in ctx.paragraphs:
        p = paragraph._p
        if p in toc_nodes or any(a.tag in (qn("w:del"), qn("w:moveFrom")) for a in p.iterancestors()):
            continue
        text = _visible(p).strip()
        match = CAPTION.match(text)
        local_fields = [f for f in fields if f["paragraph_element"] is p]
        sequence_fields = [f for f in local_fields if f["kind"] == "SEQ"]
        style_caption = any(re.search(r"caption|题注|图题|表题", s.name or "", re.I) for s in _styles(paragraph.style))
        if not style_caption and not sequence_fields and re.search(r"[。！？!?；;]", text):
            # An ordinary sentence starting with a figure label is not positive
            # evidence of a caption; explicit Caption/SEQ evidence wins.
            continue
        if not match:
            # Retain labelled SEQ captions with an empty result, and explicit
            # caption-style paragraphs using unsupported numbering alphabets.
            label = re.match(r"^(图|表|Figure|Table|Fig\.?|Tab\.?)"
                             r"(?=\s|[（(A-Za-z零一二三四五六七八九十0-9０-９]|$)", text, re.I)
            if not label or not (style_caption or sequence_fields):
                continue
            label_text = label[1]
        else:
            label_text = match[1]
        kind = _label_kind(label_text)
        adjacent = _adjacent_objects(paragraph, kind)
        if not (adjacent or sequence_fields or style_caption):
            continue
        role, chapter = _caption_role(ctx, paragraph, headings)
        record = {"paragraph": paragraph, "kind": kind, "text": text,
                  "language": "en" if label_text.isascii() else "zh", "role": role,
                  "chapter": chapter, "fields": local_fields,
                  "dynamic": bool(local_fields or p.xpath(".//w:fldSimple|.//w:fldChar|.//w:instrText")),
                  "revision": _revised(p) or any(_revised(s.element) for s in _styles(paragraph.style)),
                  "adjacent": adjacent, "position": None, "object": None,
                  "number": match[3] if match else None, "values": [], "shape": None,
                  "number_end": match.end(3) if match else 0, "title_start": match.start(5) if match else 0}
        if match:
            record["values"] = [int(n) for n in re.findall(r"[0-9０-９]+", match[3])]
            record["shape"] = (match[1], match[2], re.sub(r"[0-9０-９]+", "#", match[3]), match[4],
                               "fullwidth" if any("０" <= c <= "９" for c in match[3]) else "ascii")
        if len(adjacent) == 1:
            direction, obj = adjacent[0]
            record["object"] = obj
            record["position"] = "below" if direction == "above" else "above"
        records.append(record)
    return records


def _runs(paragraph):
    from thesis_format_checker.detectors.text_runs import effective_text_runs
    return effective_text_runs(paragraph)


def _text_signature(paragraph, start=0, end=None):
    """Compare scripts separately; mixed styles are not flattened to a mode."""
    result = defaultdict(set)
    offset = 0
    for text, props in _runs(paragraph):
        stop = offset + len(text)
        clip = text[max(0, start - offset):None if end is None else max(0, end - offset)] if stop > start else ""
        if end is not None and offset >= end:
            break
        offset = stop
        if not clip.strip():
            continue
        if re.search(r"[\u3400-\u9fff]", clip):
            result["east_asia_font"].add(props.get("east_asia_font"))
        if re.search(r"[A-Za-z0-9]", clip):
            result["ascii_font"].add(props.get("ascii_font"))
        for key in ("size_pt", "bold", "color"):
            value = props.get(key)
            result[key].add(str(value) if isinstance(value, (dict, list)) else value)
    return dict(result)


def _caption_format(record):
    paragraph = record["paragraph"]
    return {"paragraph": {k: effective_paragraph(paragraph)[k] for k in ("alignment", "space_before", "space_after")},
            "number": _text_signature(paragraph, 0, record["number_end"]),
            "title": _text_signature(paragraph, record["title_start"])}


def _key(record):
    return record["kind"], record["language"], record["role"]


def _same(values):
    return bool(values) and all(v == values[0] for v in values)


def _shape_description(shape):
    label, gap, number, suffix, digits = shape
    def spaces(value):
        if not value:
            return "无间隔"
        return value.replace(" ", "〈半角空格〉").replace("\t", "〈制表符〉").replace("\u3000", "〈全角空格〉")
    return f"标签“{label}”，标签后{spaces(gap)}，编号形态“{number.replace('#', '数字')}”，编号后{spaces(suffix)}，{'全角' if digits == 'fullwidth' else '半角'}数字"


def _border_description(value):
    if value == ("none",):
        return "无可见边框"
    if value is None:
        return "有效边框不能确认"
    style, width, color, space, shadow = value
    line = {"single": "单线", "double": "双线", "dotted": "点线", "dashed": "虚线"}.get(style, style)
    try:
        line_width = f"{float(width) / 8:g}磅" if width is not None else "未声明宽度"
    except (TypeError, ValueError):
        line_width = "无法解析宽度"
    return f"{line}，{line_width}，颜色{'自动' if color == 'auto' else color}，线外间距{space}磅，阴影{shadow}"


def _format_compare(rule, expected, actual, paragraph, prefix, keys=None):
    findings = []
    for key in expected.keys() if keys is None else keys:
        exp, got = expected.get(key), actual.get(key)
        code = _PROPERTY_CODES.get(rule, {}).get(key, rule)
        if exp is None or got is None or (isinstance(exp, set) and (None in exp or len(exp) != 1)) or (isinstance(got, set) and None in got):
            findings.append(_notice(code, prefix + LABELS.get(key, key) + "的有效值或模板唯一要求不能确认", paragraph))
        elif exp != got:
            def show(value):
                if isinstance(value, set):
                    return "、".join(describe_value(key, v) for v in sorted(value, key=str))
                return describe_value(key, value)
            findings.append(make_finding(code, _title(code), prefix + LABELS.get(key, key) + "为" + show(exp),
                                         "读取为" + show(got), paragraph))
    return findings


def _style_elements(table):
    return [s.element for s in _styles(table.style)]


def _property(sources, child, attribute="val", default=None):
    for source in sources:
        if source is None:
            continue
        item = source.find(qn("w:" + child))
        if item is not None:
            value = item.get(qn("w:" + attribute))
            if value is not None:
                return value
    return default


def _border(node):
    if node is None:
        return ("none",)
    value = node.get(qn("w:val"))
    if value in ("nil", "none"):
        return ("none",)
    if not value or any(node.get(qn("w:" + a)) is not None for a in ("themeColor", "themeTint", "themeShade")):
        return None
    # These are OOXML omission values, not template/school constants.
    return (value, node.get(qn("w:sz")), node.get(qn("w:color"), "auto"),
            node.get(qn("w:space"), "0"), node.get(qn("w:shadow"), "0"))


def _border_node(sources, name):
    for source in sources:
        if source is not None:
            node = source.find(qn("w:" + name))
            if node is not None:
                return node
    return None


def _row_role(row):
    props = row._tr.trPr
    marker = props.find(qn("w:tblHeader")) if props is not None else None
    if marker is not None and marker.get(qn("w:val"), "1") not in ("0", "false", "off"):
        return "header"
    nonempty = [cell.text.strip() for cell in row.cells if cell.text.strip()]
    if nonempty and TOTAL.fullmatch(nonempty[0]):
        return "total"
    return "body"


def _table_snapshot(table, caption):
    styles = _style_elements(table)
    tr_nodes = list(table._tbl.findall(qn("w:tr")))
    tc_nodes = [tc for tr in tr_nodes for tc in tr.findall(qn("w:tc"))]
    shape_issue = bool(table._tbl.xpath("./w:tr/w:tc/w:tcPr/w:gridSpan|./w:tr/w:tc/w:tcPr/w:vMerge|./w:tr/w:trPr/w:gridBefore|./w:tr/w:trPr/w:gridAfter|./w:tblPr/w:tblCellSpacing|./w:tblPr/w:bidiVisual|./w:tblPr/w:tblpPr|./w:tr/w:tc/w:tbl|.//w:tcBorders/w:start|.//w:tcBorders/w:end|.//w:tblBorders/w:start|.//w:tblBorders/w:end"))
    shape_issue = shape_issue or _revised(table._tbl) or any(_revised(s) or s.xpath("./w:tblPr/w:tblCellSpacing|./w:tblPr/w:bidiVisual|./w:tblPr/w:tblpPr") for s in styles)
    conditional_border = any(s.xpath("./w:tblStylePr//w:tblBorders|./w:tblStylePr//w:tcBorders") for s in styles)
    conditional_text = any(s.xpath("./w:tblStylePr//w:rPr|./w:tblStylePr//w:pPr|./w:tblStylePr//w:shd|./w:tblStylePr//w:vAlign|./w:tblStylePr//w:trHeight") for s in styles)
    # Non-conditional table style rPr/pPr is below paragraph/run formatting in
    # Word. The shared text resolver marks unresolved table-style effects.
    table_props = [table._tbl.tblPr] + [s.find(qn("w:tblPr")) for s in styles]
    style_tc = [s.find(qn("w:tcPr")) for s in styles]
    style_tr = [s.find(qn("w:trPr")) for s in styles]
    rows = []
    for ri, row in enumerate(table.rows):
        role = _row_role(row)
        props = [row._tr.trPr] + style_tr
        height = _property(props, "trHeight", "val")
        height_rule = _property(props, "trHeight", "hRule", "auto")
        row_record = {"role": role, "height": (height, height_rule), "cells": []}
        exception = row._tr.find(qn("w:tblPrEx"))
        tbl_border_sources = [p.find(qn("w:tblBorders")) for p in [exception] + table_props if p is not None]
        for ci, cell in enumerate(row.cells):
            sources = [cell._tc.tcPr] + style_tc
            cell_border_sources = [p.find(qn("w:tcBorders")) for p in sources if p is not None]
            borders = {}
            for side in ("top", "bottom", "left", "right", "tl2br", "tr2bl"):
                direct = _border_node(cell_border_sources, side)
                if direct is None:
                    edge = side
                    if side == "top" and ri:
                        edge = "insideH"
                    elif side == "bottom" and ri < len(table.rows) - 1:
                        edge = "insideH"
                    elif side == "left" and ci:
                        edge = "insideV"
                    elif side == "right" and ci < len(row.cells) - 1:
                        edge = "insideV"
                    direct = _border_node(tbl_border_sources, edge)
                borders[side] = _border(direct)
            shade = next((p.find(qn("w:shd")) for p in sources if p is not None and p.find(qn("w:shd")) is not None), None)
            if shade is None:
                shade = next((p.find(qn("w:shd")) for p in table_props if p is not None and p.find(qn("w:shd")) is not None), None)
            shading = ("clear", "auto", "auto") if shade is None else (
                shade.get(qn("w:val"), "clear"), shade.get(qn("w:color"), "auto"), shade.get(qn("w:fill"), "auto"))
            if shade is not None and any("theme" in k for k in shade.attrib):
                shading = None
            paragraphs = []
            for paragraph in cell.paragraphs:
                fmt = _text_signature(paragraph)
                fmt["alignment"] = effective_paragraph(paragraph)["alignment"]
                style_alignment = any(s.xpath("./w:pPr/w:jc") for s in styles)
                explicit_alignment = paragraph._p.xpath("./w:pPr/w:jc") or any(s.element.xpath("./w:pPr/w:jc") for s in _styles(paragraph.style))
                if style_alignment and not explicit_alignment:
                    fmt["alignment"] = None
                paragraphs.append((paragraph, fmt))
            row_record["cells"].append({"borders": borders, "vertical": _property(sources, "vAlign", default="top"),
                                        "shading": shading, "paragraphs": paragraphs})
        rows.append(row_record)
    # At zero cell spacing Word resolves shared-edge conflicts. Identical
    # declarations and one visible side versus nil are unambiguous; two
    # different visible borders need the full Word conflict algorithm.
    border_conflict = False
    def shared(left, left_edge, right, right_edge):
        nonlocal border_conflict
        first, second = left["borders"][left_edge], right["borders"][right_edge]
        if first == second:
            return
        if first == ("none",):
            left["borders"][left_edge] = second
        elif second == ("none",):
            right["borders"][right_edge] = first
        else:
            border_conflict = True
    if not shape_issue:
        for ri, row in enumerate(rows):
            for ci, cell in enumerate(row["cells"]):
                if ci:
                    shared(row["cells"][ci-1], "right", cell, "left")
                if ri and ci < len(rows[ri-1]["cells"]):
                    shared(rows[ri-1]["cells"][ci], "bottom", cell, "top")
    return {"table": table, "caption": caption, "rows": rows, "shape_issue": shape_issue,
            "border_issue": shape_issue or conditional_border or border_conflict,
            "format_issue": shape_issue or conditional_text,
            "style_name": table.style.name if table.style else None}


def _table_records(ctx, captions):
    result, seen = [], set()
    for caption in captions:
        obj = caption["object"]
        if caption["kind"] != "table" or obj is None or obj in seen:
            continue
        seen.add(obj)
        try:
            table = Table(obj, ctx.document._body)
            result.append(_table_snapshot(table, caption))
        except Exception as exc:
            result.append({"caption": caption, "error": str(exc)})
    return result


def extract_figure_table_baseline(ctx, mapping=None):
    del mapping
    records = collect_caption_records(ctx)
    groups = defaultdict(list)
    for record in records:
        try:
            record["format"] = _caption_format(record)
        except Exception as exc:
            record["format_error"] = str(exc)
        groups[_key(record)].append(record)
    return {"figures": {"captions": records, "groups": dict(groups),
                                "tables": _table_records(ctx, records), "text": ctx.text,
                                "has_textboxes": ctx.has_textboxes}}


def _sequence_baseline(candidates):
    manual = [r for r in candidates if not r["dynamic"] and r["values"]]
    if len(manual) < 2 or not _same([len(r["values"]) for r in manual]):
        return None
    length = len(manual[0]["values"])
    if length > 2:
        return None
    grouped = defaultdict(list)
    for r in manual:
        grouped[tuple(r["values"][:-1])].append(r["values"][-1])
    pairs = [b - a for group in grouped.values() for a, b in zip(group, group[1:])]
    if not pairs or not _same(pairs) or pairs[0] <= 0:
        return None
    starts = [g[0] for g in grouped.values()]
    known = [r for r in manual if r["chapter"] is not None]
    linked = length == 2 and len({r["chapter"] for r in known}) >= 2 and all(r["values"][0] == r["chapter"] for r in known)
    return {"step": pairs[0], "start": starts[0] if _same(starts) else None, "length": length,
            "chapter_link": linked, "scoped": length == 2}


def _position_check(record, candidates):
    if not candidates:
        return [_notice("078", "模板没有可匹配图表类型、语言及区域的题注位置示例", record["paragraph"])]
    positions = [r["position"] for r in candidates]
    obj = record["object"]
    uncertain_template = any(r["object"] is not None and (r["object"].xpath(".//wp:anchor|.//w:pict|.//w:object|.//w:txbxContent") or _revised(r["object"])) for r in candidates)
    uncertain_actual = obj is not None and (obj.xpath(".//wp:anchor|.//w:pict|.//w:object|.//w:txbxContent") or _revised(obj))
    if not _same(positions) or positions[0] is None or record["position"] is None or uncertain_template or uncertain_actual:
        return [_notice("078", "题注关联不唯一或存在浮动/嵌入对象，无法将存储顺序认定为视觉位置", record["paragraph"])]
    if record["position"] != positions[0]:
        names = {"above": "图表对象上方", "below": "图表对象下方"}
        # 表题应在表格上方（078.2），图题应在图片下方（078.1）
        code = "078.2" if re.match(r"^\s*(?:表|Table)", record["text"], re.I) else "078.1"
        return [make_finding(code, _title(code), names[positions[0]], names[record["position"]],
                             record["paragraph"])]
    return []


def _caption_checks(data, records, missing_objects=False):
    findings, sequence_states = [], {}
    uncertain_groups = {_key(r) for r in records if r["revision"] or r["dynamic"] or not r["values"] or
                        re.search(r"续表|续图|continued|[（(][a-zA-Z][）)]", r["text"], re.I) or
                        (r["object"] is not None and r["object"].xpath(".//wp:anchor|.//w:pict|.//w:object"))}
    for record in records:
        p = record["paragraph"]
        candidates = data["groups"].get(_key(record), [])
        if record["revision"] or any(c["revision"] for c in candidates):
            for rule in ("078", "079", "080", "081"):
                findings.append(_notice(rule, "题注含未接受的修订，最终文字、格式或编号视图不能确认", p))
            continue
        findings.extend(_position_check(record, candidates))
        if not candidates:
            for rule in ("079", "080", "081"):
                findings.append(_notice(rule, "模板缺少可匹配的图表题注基准", p))
            continue
        try:
            formats = [c.get("format") for c in candidates]
            if not _same(formats) or not formats[0]:
                findings.append(_notice("079", "模板同角色题注存在不同格式或格式读取失败，不能唯一选择", p))
            elif record["dynamic"]:
                findings.append(_notice("079", "动态题注域的当前文字边界及最终显示格式需Word解析", p))
            else:
                actual = _caption_format(record)
                for part in ("paragraph", "number", "title"):
                    findings.extend(_format_compare("079", formats[0][part], actual[part], p,
                                                    {"paragraph": "题注段落", "number": "题注标签编号", "title": "题名"}[part],
                                                    tuple(k for k in ("east_asia_font", "ascii_font", "size_pt") if k in formats[0][part]) if part != "paragraph" else None))
        except Exception as exc:
            findings.append(_notice("079", "题注有效格式读取失败：" + str(exc), p))
        if record["dynamic"] or any(c["dynamic"] for c in candidates):
            for rule in ("080", "081"):
                findings.append(_notice(rule, "SEQ/STYLEREF等域只能读取指令和旧缓存，当前计算编号及完整形式需Word更新后复核；未更新原文件", p))
            continue
        shapes = [c["shape"] for c in candidates]
        if not record["shape"] or not _same(shapes) or shapes[0] is None:
            findings.append(_notice("081", "编号采用未支持的数字形式或模板同角色示例形式不唯一", p))
        elif record["shape"] != shapes[0]:
            findings.append(make_finding("081", TITLES["081"], _shape_description(shapes[0]),
                                         _shape_description(record["shape"]), p))
        sequence = _sequence_baseline(candidates)
        if _key(record) in uncertain_groups or missing_objects or data.get("has_textboxes"):
            findings.append(_notice("080", "同一编号序列含动态域、未接受修订、子图/续表、浮动或文本框等未能确认的节点，不能跳过该节点推断后续编号", p))
            continue
        obj = record["object"]
        if obj is not None and obj.xpath(".//wp:anchor|.//w:pict|.//w:object"):
            findings.append(_notice("080", "浮动图表存储顺序不等于阅读顺序，不能确认编号顺序", p))
            continue
        if sequence is None or not record["values"] or len(record["values"]) != sequence["length"]:
            findings.append(_notice("080", "模板缺少可确认的连续编号样例，或当前存在子图/续表/不同编号层次", p))
            continue
        if re.search(r"续表|续图|continued|[（(][a-zA-Z][）)]", record["text"], re.I):
            findings.append(_notice("080", "子图或续表编号需要明确独立序列范围", p))
            continue
        prefix = tuple(record["values"][:-1])
        key = (_key(record), prefix)
        previous = sequence_states.get(key)
        current = record["values"][-1]
        expected = previous + sequence["step"] if previous is not None else sequence["start"]
        if expected is not None and current != expected:
            findings.append(make_finding("080.1", _title("080.1"), f"模板样例递增{sequence['step']}，此处序号应为{expected}",
                                         f"当前序号为{current}（编号{record['number']}）", p))
        if sequence["chapter_link"]:
            if record["chapter"] is None:
                findings.append(_notice("080", "模板含按章编号样例，但当前所属一级标题章号不能确认", p))
            elif record["values"][0] != record["chapter"]:
                findings.append(make_finding("080.3", _title("080.3"), f"章节前缀与所属章号{record['chapter']}一致",
                                             f"题注章节前缀为{record['values'][0]}", p))
        sequence_states[key] = current
    return findings


def _table_candidates(data, actual):
    candidates = [t for t in data["tables"] if _key(t["caption"]) == _key(actual["caption"])]
    if len(candidates) > 1 and "style_name" in actual:
        matching = [t for t in candidates if t.get("style_name") == actual["style_name"]]
        if matching:
            candidates = matching
    return candidates


def _table_checks(data, tables):
    findings = []
    for actual in tables:
        p = actual["caption"]["paragraph"]
        candidates = _table_candidates(data, actual)
        if not candidates or "error" in actual or any("error" in t for t in candidates):
            for rule in ("083", "084"):
                findings.append(_notice(rule, "表格读取失败或没有可匹配的模板表格", p))
            continue
        for rule, issue in (("083", "border_issue"), ("084", "format_issue")):
            if actual[issue] or any(t[issue] for t in candidates):
                findings.append(_notice(rule, "合并/嵌套/浮动表格、单元格间距或条件表样式使有效属性不能可靠解析", p))
                continue
            for ri, row in enumerate(actual["rows"]):
                role = row["role"]
                template_rows = [r for t in candidates for r in t["rows"] if r["role"] == role]
                if not template_rows:
                    findings.append(_notice(rule, "模板没有可确认的" + {"header": "表头", "body": "表身", "total": "合计"}[role] + "行角色；表头仅依据重复标题行标记识别，合计仅依据明确标签", p))
                    continue
                for ci, cell in enumerate(row["cells"]):
                    anchor = cell["paragraphs"][0][0] if cell["paragraphs"] else p
                    # Keep boundary cells distinct; allow a variable count of
                    # uniform body rows by matching top/middle/bottom roles.
                    selected = []
                    for table in candidates:
                        for rj, tr in enumerate(table["rows"]):
                            if tr["role"] != role:
                                continue
                            if rule == "083" and ((ri == 0) != (rj == 0) or (ri == len(actual["rows"]) - 1) != (rj == len(table["rows"]) - 1)):
                                continue
                            if len(tr["cells"]) == len(row["cells"]):
                                selected.append((tr, tr["cells"][ci]))
                            elif rule == "084":
                                selected.extend((tr, tc) for tc in tr["cells"])
                    if not selected:
                        findings.append(_notice(rule, f"第{ri+1}行第{ci+1}列无法匹配模板同角色及边界位置", anchor))
                        continue
                    prefix = f"第{ri+1}行第{ci+1}列"
                    if rule == "083":
                        for edge in ("top", "bottom", "left", "right", "tl2br", "tr2bl"):
                            expected = [tc["borders"][edge] for _, tc in selected]
                            got = cell["borders"][edge]
                            if not _same(expected) or expected[0] is None or got is None:
                                findings.append(_notice(rule, prefix + "的有效边框存在模板冲突或主题色等未解析属性", anchor))
                            elif got != expected[0]:
                                # 竖线问题 083.1，横线问题 083.2
                                code = _BORDER_CODES.get(edge, rule)
                                findings.append(make_finding(code, _title(code), prefix + {"top":"上", "bottom":"下", "left":"左", "right":"右", "tl2br":"左上至右下对角", "tr2bl":"右上至左下对角"}[edge] + "边框为" + _border_description(expected[0]),
                                                             "当前边框为" + _border_description(got), anchor))
                    else:
                        for key in ("height", "vertical", "shading"):
                            expected = [tr[key] if key == "height" else tc[key] for tr, tc in selected]
                            got = row[key] if key == "height" else cell[key]
                            if not _same(expected):
                                findings.append(_notice(rule, prefix + LABELS[key] + "模板同角色基准不唯一", anchor))
                            else:
                                findings.extend(_format_compare(rule, {key: expected[0]}, {key: got}, anchor, prefix))
                        template_formats = [fmt for _, tc in selected for tp, fmt in tc["paragraphs"] if _visible(tp._p).strip()]
                        for ap, fmt in cell["paragraphs"]:
                            if not _visible(ap._p).strip():
                                continue
                            if not _same(template_formats):
                                findings.append(_notice(rule, prefix + "模板同角色文字格式不唯一", ap))
                            elif template_formats:
                                findings.extend(_format_compare(rule, template_formats[0], fmt, ap, prefix))
    return findings


def check_figure_table_errors(baseline, ctx, mapping=None):
    del mapping
    data = baseline.get("figures", {})
    if not data:
        return []
    findings = []
    try:
        records = collect_caption_records(ctx)
        findings.extend(_caption_checks(data, records, ctx.has_textboxes))
        tables = _table_records(ctx, records)
        findings.extend(_table_checks(data, tables))
        associated = {r["object"] for r in records if r["kind"] == "table" and r["object"] is not None}
        headings = collect_heading_records(ctx)
        for table in ctx.document.tables:
            if table._tbl in associated or table._tbl.xpath(".//m:oMath|.//m:oMathPara|.//w:object|.//w:drawing"):
                continue
            paragraphs = [p for p in ctx.paragraphs if any(a is table._tbl for a in p._p.iterancestors())]
            if not paragraphs:
                continue
            anchor = paragraphs[0]
            earlier = [h for h in headings if ctx.index(h["paragraph"]) < ctx.index(anchor) and h["level"] == 1]
            # Known front matter and reference tables are not data-table roles.
            if earlier and earlier[-1]["role"] != "body":
                continue
            joined = " ".join(p.text for p in paragraphs)
            if re.search(r"论文题目|指导教师|作者姓名|提交日期|学位授予|密级|分类号", joined):
                continue
            for rule in ("083", "084"):
                findings.append(_notice(rule, "此表没有可可靠关联的题注，不能从模板任意选择表格用途及格式基准", anchor))
        if ctx.has_textboxes:
            anchor = ctx.paragraphs[0] if ctx.paragraphs else None
            for rule in ("078", "079", "080", "081"):
                findings.append(_notice(rule, "文档含文本框，框内题注及其对象关联未纳入可读题注清单", anchor))
    except Exception as exc:
        anchor = ctx.paragraphs[0] if ctx.paragraphs else None
        for rule in TITLES:
            findings.append(_notice(rule, "图表对象局部读取异常：" + str(exc), anchor))
    return findings

def collect_caption_records(ctx):
    """带缓存的入口：同一份文档只扫描一次（见 core.docx_context.memo）。"""

    return memo(ctx, "caption_records", lambda: _collect_caption_records_impl(ctx))
