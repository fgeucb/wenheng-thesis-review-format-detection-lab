"""Read-only TOC checks (056-066); never refresh a source field or paginate it.

The field cache is evidence of its current state, not a Word rendering engine.
Only explicit template switches and styles establish the baseline. Ambiguous
scope, unsupported fields and physical page results produce UNKNOWN findings.
"""
from __future__ import annotations

import re
from collections import defaultdict

from thesis_format_checker.core.xmlutil import qn, visible_text

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, effective_paragraph, effective_runs, make_finding, unknown, describe_value, memo
from thesis_format_checker.detectors.fields import scan_fields, scan_body_fields


TITLES = {
    "056": "目录自动生成结构不符合模板", "057": "目录收录标题缺失或多余",
    "058": "目录标题、编号或层级与正文不一致", "059": "目录页码对应关系",
    "060": "目录层级显示范围不符合模板", "061": "目录条目格式不符合模板",
    "062": "目录引导线不符合模板", "063": "目录页码对齐不符合模板",
    "064": "目录标题格式不符合模板", "065": "目录条目结构异常", "066": "目录与当前正文同步状态",
}
_PARAGRAPH_KEYS = ("alignment", "left_indent", "right_indent", "first_line_indent")
_ENTRY_PARAGRAPH_KEYS = ("left_indent", "right_indent", "first_line_indent")
_TITLE_KEYS = _PARAGRAPH_KEYS + ("space_before", "space_after")
_LABELS = {"alignment": "对齐", "left_indent": "左缩进", "right_indent": "右缩进",
           "first_line_indent": "首行或悬挂缩进", "space_before": "段前间距", "space_after": "段后间距",
           "east_asia_font": "中文字体", "ascii_font": "西文字体", "size_pt": "字号", "bold": "加粗"}
_TAB_ALIGNMENT = {"left": "左对齐", "right": "右对齐", "center": "居中", "decimal": "小数点对齐", "bar": "竖线"}
_TAB_LEADER = {"none": "无", "dot": "点线", "hyphen": "短横线", "underscore": "下划线",
               "heavy": "粗线", "middleDot": "中点线"}


def _notice(rule, reason, paragraph=None):
    return unknown(rule, _title(rule), UNREADABLE + "；" + reason, paragraph)


#: 目录条目格式属性 → 细粒度规则编号
_ENTRY_FORMAT_CODES = {
    "east_asia_font": "061.1", "ascii_font": "061.1", "hansi_font": "061.1",
    "size_pt": "061.2", "bold": "061.3",
    "alignment": "061.4", "first_line_indent": "061.4",
    "left_indent": "061.4", "right_indent": "061.4", "color": "061.4",
}
#: “目录”标题格式属性 → 细粒度规则编号
_TITLE_FORMAT_CODES = {
    "east_asia_font": "064.1", "ascii_font": "064.1", "hansi_font": "064.1",
    "color": "064.1", "bold": "064.5",
    "size_pt": "064.2",
    "alignment": "064.3", "first_line_indent": "064.3",
    "left_indent": "064.3", "right_indent": "064.3",
    "space_before": "064.4", "space_after": "064.4",
}


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else TITLES[code]


def _norm(text):
    return re.sub(r"\s+", " ", text or "").strip()


def _visible(p):
    """元素内可见文字（跳过删除内容/文本框，一次遍历；见 core.xmlutil.visible_text）。"""

    return visible_text(p)


def _style_chain(paragraph):
    seen = set()
    style = paragraph.style
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style
        style = style.base_style


def _toc_level(paragraph):
    for style in _style_chain(paragraph):
        for name in (style.name, style.style_id):
            match = re.fullmatch(r"(?:TOC|目录)\s*([1-9])", name or "", re.I)
            if match:
                return int(match.group(1))
    return None


def _options(instruction):
    tokens = re.findall(r'"[^"]*"|[^\s]+', instruction.strip())
    opts = {}
    problems = []
    for token in tokens[1:]:
        if token.startswith("\\"):
            key = token[1:].lower()
            if key in opts:
                problems.append("同一目录开关重复出现")
            opts[key] = None
        elif opts:
            if opts[key] is not None:
                problems.append("目录开关参数无法唯一解析")
            opts[key] = token.strip('"')
        else:
            problems.append("目录域含无法解析的参数")
    levels = None
    if "o" in opts:
        match = re.fullmatch(r"([1-9])-([1-9])", opts["o"] or "")
        if match and int(match.group(1)) <= int(match.group(2)):
            levels = (int(match.group(1)), int(match.group(2)))
        else:
            problems.append("目录大纲层级范围无法读取")
    style_map = {}
    if "t" in opts:
        parts = (opts["t"] or "").split(",")
        if len(parts) % 2:
            problems.append("目录自定义样式映射格式无法读取")
        else:
            for name, level in zip(parts[::2], parts[1::2]):
                if not name.strip() or not re.fullmatch("[1-9]", level.strip()):
                    problems.append("目录自定义样式映射格式无法读取")
                else:
                    style_map[name.strip().casefold()] = int(level.strip())
    unsupported = set(opts) - {"o", "u", "t", "h", "z", "n", "p", "b", "*"}
    if unsupported:
        problems.append("目录包含暂不解析的来源/显示开关：" + "、".join(sorted(unsupported)))
    if "b" in opts:
        problems.append("局部书签范围目录需核对范围边界")
    if not (levels or style_map or "u" in opts):
        problems.append("模板未显式给出可解析的标题来源范围")
    suppress = None
    if "n" in opts and opts["n"]:
        match = re.fullmatch(r"([1-9])-([1-9])", opts["n"])
        if match and int(match.group(1)) <= int(match.group(2)):
            suppress = (int(match.group(1)), int(match.group(2)))
        else:
            problems.append("目录隐藏页码的层级范围无法读取")
    return {"raw": opts, "levels": levels, "styles": style_map, "outline": "u" in opts,
            "problems": problems, "suppressed": suppress}


def _page_expected(options, level):
    if "n" not in options["raw"]:
        return True
    bounds = options["suppressed"]
    if bounds is None:
        return False
    return level is None or not bounds[0] <= level <= bounds[1]


def _source_signature(options):
    return (options["levels"], options["outline"], tuple(sorted(options["styles"].items())))


def _targets(paragraph, all_fields):
    targets = {n.get(qn("w:anchor")) for n in paragraph._p.iter(qn("w:hyperlink")) if n.get(qn("w:anchor"))}
    for field in all_fields:
        if field.get("paragraph_element") is not paragraph._p:
            continue
        instruction = field.get("instruction", "")
        kind = field.get("kind", "").upper()
        if kind == "PAGEREF":
            match = re.match(r'\s*PAGEREF\s+(?:"([^"]+)"|([^\s\\]+))', instruction, re.I)
            if match:
                targets.add(match.group(1) or match.group(2))
        elif kind == "HYPERLINK":
            match = re.search(r'\\l\s+(?:"([^"]+)"|([^\s\\]+))', instruction, re.I)
            if match:
                targets.add(match.group(1) or match.group(2))
    return targets


def __snapshot_impl(ctx):
    fields = scan_body_fields(ctx)
    toc_fields = [f for f in fields if f.get("kind", "").upper() == "TOC"]
    by_element = {p._p: p for p in ctx.paragraphs}
    groups = []
    for field in toc_fields:
        options = _options(field.get("instruction", ""))
        elements = list(field.get("result_paragraph_elements", []))
        if not elements:
            for node in field.get("result_elements", []):
                p = next((a for a in node.iterancestors() if a.tag == qn("w:p")), None)
                if p is not None and p not in elements:
                    elements.append(p)
        entries = []
        for element in elements:
            paragraph = by_element.get(element)
            if paragraph is None:
                continue
            text = _visible(element)
            if not text.strip() and element is field.get("paragraph_element"):
                continue
            level = _toc_level(paragraph)
            targets = _targets(paragraph, fields)
            # Field-only boundary paragraphs and spacing paragraphs are not entries.
            if level is None and not targets:
                continue
            separator = options["raw"].get("p") or "\t"
            has_separator = separator in text and _page_expected(options, level)
            label, page = text.rsplit(separator, 1) if has_separator else (text, None)
            entries.append({"paragraph": paragraph, "text": text, "label": _norm(label), "page": page,
                            "level": level, "targets": targets, "separator": separator,
                            "has_separator": has_separator})
        groups.append({"field": field, "options": options, "entries": entries,
                       "anchor": by_element.get(field.get("paragraph_element"))})
    titles = [p for p in ctx.paragraphs if re.sub(r"\s+", "", _visible(p._p)).casefold() in
              {"目录", "contents", "tableofcontents"}]
    manual = [p for p in ctx.paragraphs if _toc_level(p) is not None]
    return {"groups": groups, "titles": titles, "manual": manual, "errors": list(getattr(fields, "errors", [])),
            "fields": fields, "ctx": ctx}


def _tabs(paragraph):
    values = {}
    defaults = paragraph.part.document.styles.element.xpath("./w:docDefaults/w:pPrDefault/w:pPr")
    props = defaults + [s.element.pPr for s in reversed(list(_style_chain(paragraph)))] + [paragraph._p.pPr]
    for prop in props:
        if prop is None:
            continue
        tabs = prop.find(qn("w:tabs"))
        if tabs is None:
            continue
        for tab in tabs:
            try:
                pos = int(tab.get(qn("w:pos")))
            except (TypeError, ValueError):
                return None
            kind = tab.get(qn("w:val"))
            if kind == "clear":
                values.pop(pos, None)
            elif kind is None:
                return None
            else:
                values[pos] = (kind, tab.get(qn("w:leader"), "none"))
    return tuple((pos, *spec) for pos, spec in sorted(values.items()))


def _format(paragraph, entry=False):
    profile = effective_paragraph(paragraph)
    raw = _visible(paragraph._p)
    split = raw.rfind("\t") if entry else -1
    offset = 0
    values = defaultdict(set)
    runs = effective_runs(paragraph)
    if "".join(text for text, _ in runs) != raw:
        profile["run_problem"] = "嵌套域或特殊对象的有效文字格式无法完整读取"
    for text, props in runs:
        for char in text:
            zone = "页码" if split >= 0 and offset > split else "条目文字"
            offset += 1
            if char.isspace():
                continue
            script = "中文" if "\u3400" <= char <= "\u9fff" else "西文"
            for key in ("east_asia_font" if script == "中文" else "ascii_font", "size_pt", "bold"):
                values[(zone, script, key)].add(props[key])
    profile["runs"] = {key: next(iter(value)) if len(value) == 1 else None for key, value in values.items()}
    profile["tabs"] = _tabs(paragraph)
    return profile


def _consensus(profiles, key):
    values = [p.get(key) for p in profiles]
    return values[0] if values and values[0] is not None and all(v == values[0] for v in values) else None


def _compare_format(rule, templates, paragraph, entry=False, codes=None):
    """比较目录条目/“目录”标题格式：**一个属性一条结论**。"""

    findings = []
    current = _format(paragraph, entry)
    diffs, uncertain = [], []
    for key in _ENTRY_PARAGRAPH_KEYS if entry else _TITLE_KEYS:
        expected = _consensus(templates, key)
        actual = current.get(key)
        if expected is None or actual is None:
            uncertain.append(_LABELS[key])
        elif expected != actual:
            diffs.append((key, _LABELS[key] + "=" + describe_value(key, expected),
                          _LABELS[key] + "=" + describe_value(key, actual)))
    run_keys = set().union(*(set(p["runs"]) for p in templates))
    if current.get("run_problem") or any(p.get("run_problem") for p in templates):
        run_keys = set()
    for key in sorted(run_keys):
        values = [p["runs"].get(key) for p in templates if key in p["runs"]]
        expected = values[0] if values and values[0] is not None and all(v == values[0] for v in values) else None
        if key not in current["runs"]:
            continue
        actual = current["runs"][key]
        label = "/".join(key[:2]) + _LABELS[key[2]]
        if expected is None or actual is None:
            uncertain.append(label)
        elif expected != actual:
            diffs.append((key[2], label + "=" + describe_value(key[2], expected),
                          label + "=" + describe_value(key[2], actual)))
    if current.get("run_problem") or any(p.get("run_problem") for p in templates):
        uncertain.append("特殊对象中的文字有效格式")
    for key, want, actual_text in diffs:
        code = (codes or {}).get(key, rule)
        findings.append(make_finding(code, _title(code), want, actual_text, paragraph))
    if uncertain:
        findings.append(_notice(rule, "模板未形成唯一有效值或对象不可读：" + "、".join(uncertain), paragraph))
    return findings


def _bookmarks(ctx):
    bookmarks = defaultdict(list)
    for paragraph in ctx.paragraphs:
        for bookmark in paragraph._p.iter(qn("w:bookmarkStart")):
            name = bookmark.get(qn("w:name"))
            if name:
                bookmarks[name].append(paragraph)
    return bookmarks


def _source_level(record, options):
    paragraph = record["paragraph"]
    style = paragraph.style
    for name in (style.name, style.style_id):
        if name.casefold() in options["styles"]:
            return options["styles"][name.casefold()]
    level = record["level"]
    bounds = options["levels"]
    if bounds and bounds[0] <= level <= bounds[1]:
        if options["outline"]:
            return level
        if any(re.fullmatch(r"(?:Heading|标题)\s*[1-9]", name or "", re.I)
               for name in (style.name, style.style_id)):
            return level
    elif options["outline"] and bounds is None:
        return level
    return None


def _records(ctx, options, mapping=None):
    from thesis_format_checker.detectors.headings import collect_heading_records
    records = collect_heading_records(ctx, mapping)
    # A TOC custom style may legitimately have no outline level at all.
    known = {record["paragraph"]._p for record in records}
    for paragraph in ctx.paragraphs:
        if paragraph._p in known or _toc_level(paragraph) is not None:
            continue
        style = paragraph.style
        level = next((options["styles"][name.casefold()] for name in (style.name, style.style_id)
                      if name.casefold() in options["styles"]), None)
        if level is not None:
            records.append({"paragraph": paragraph, "level": level, "text": _visible(paragraph._p),
                            "title": _visible(paragraph._p), "number": None, "automatic": False,
                            "numbering_unknown": None})
    return records


def extract_toc_baseline(ctx, mapping=None):
    snapshot = _snapshot(ctx)
    formats = defaultdict(list)
    for group in snapshot["groups"]:
        for entry in group["entries"]:
            if entry["level"] is not None:
                formats[entry["level"]].append(_format(entry["paragraph"], True))
    snapshot["formats"] = dict(formats)
    snapshot["title_formats"] = [_format(p) for p in snapshot["titles"]]
    return {"toc": snapshot}


def _check_tabs(templates, entry, expected_page):
    if not expected_page:
        return []
    paragraph = entry["paragraph"]
    expected = _consensus(templates, "tabs")
    actual = _tabs(paragraph)
    if expected is None or actual is None:
        return [_notice(rule, "目录制表位基准冲突或无法读取", paragraph) for rule in ("062", "063")]
    if not expected:
        return [_notice(rule, "模板没有可确认的页码制表位基准，不能推定点引导线或右对齐", paragraph)
                for rule in ("062", "063")]
    if len(expected) != 1 or (actual and len(actual) != 1):
        return [_notice(rule, "目录条目有多个制表位，无法唯一确认页码采用哪一个制表位", paragraph)
                for rule in ("062", "063")]
    expected_tab = expected[-1]
    expected_text = (f"页码制表位位置={expected_tab[0] / 20:g}磅，对齐=" +
                     _TAB_ALIGNMENT.get(expected_tab[1], expected_tab[1]) + "，引导符=" +
                     _TAB_LEADER.get(expected_tab[2], expected_tab[2]))
    if not actual or not entry["has_separator"]:
        # 缺少页码分隔/制表位：引导线（062.1）与页码对齐（063.1）各一条
        return [make_finding(code, _title(code), expected_text,
                             "未找到对应页码分隔制表符或有效制表位", paragraph)
                for code in ("062.1", "063.1")]
    current_tab = actual[-1]
    findings = []
    if expected_tab[2] != current_tab[2]:
        findings.append(make_finding("062.2", _title("062.2"), "引导符=" + _TAB_LEADER.get(expected_tab[2], expected_tab[2]),
                                     "引导符=" + _TAB_LEADER.get(current_tab[2], current_tab[2]), paragraph))
    if expected_tab[:2] != current_tab[:2]:
        findings.append(make_finding("063.1", _title("063.1"), expected_text,
                                     f"制表位位置={current_tab[0] / 20:g}磅，对齐=" + _TAB_ALIGNMENT.get(current_tab[1], current_tab[1]), paragraph))
    return findings


def _check_links(template_group, current_group, ctx, mapping=None):
    options = template_group["options"]
    anchor = current_group["anchor"]
    if any(next(ctx.document.element.iter(qn("w:" + name)), None) is not None
           for name in ("ins", "del", "moveFrom", "moveTo")):
        return [_notice(rule, "文档含未接受的修订，目录来源/书签与当前显示版本不能唯一对应", anchor)
                for rule in ("057", "058", "066")]
    if not current_group["entries"]:
        return [_notice(rule, "目录域没有可确认的结果条目；未刷新域，无法判定当前来源收录结果", anchor)
                for rule in ("057", "058", "066")]
    if options["problems"] or current_group["options"]["problems"]:
        reason = "；".join(options["problems"] + current_group["options"]["problems"])
        return [_notice(rule, reason, anchor) for rule in ("057", "058", "066")]
    if _source_signature(options) != _source_signature(current_group["options"]):
        return [_notice(rule, "目录来源配置与模板不同，先核对060后再判断内容同步", anchor) for rule in ("057", "058", "066")]
    records = _records(ctx, options, mapping)
    by_element = {record["paragraph"]._p: record for record in records}
    wanted = {record["paragraph"]._p: (record, _source_level(record, options)) for record in records
              if _source_level(record, options) is not None}
    bookmarks = _bookmarks(ctx)
    linked, findings, incomplete = set(), [], False
    for entry in current_group["entries"]:
        paragraph = entry["paragraph"]
        if len(entry["targets"]) != 1:
            incomplete = True
            continue
        target = next(iter(entry["targets"]))
        destinations = bookmarks.get(target, [])
        if len(destinations) != 1:
            incomplete = True
            findings.append(_notice("057", "目录目标书签缺失或重复，不能断言正文标题不存在：" + target, paragraph))
            continue
        destination = destinations[0]
        record = by_element.get(destination._p)
        if record is None:
            incomplete = True
            findings.append(_notice("057", "目录书签目标不是可确认的标题/样式来源", paragraph))
            continue
        level = _source_level(record, options)
        if level is None:
            findings.append(make_finding("057", TITLES["057"], "仅收录模板明确来源范围中的标题",
                                         "条目指向当前收录范围外标题：" + record["text"], paragraph))
            continue
        linked.add(destination._p)
        diffs = []
        current_title = _norm(record["text"])
        if record.get("numbering_unknown"):
            findings.append(_notice("058", "正文标题含动态域或编号当前显示结果无法可靠计算", paragraph))
            current_title = None
        elif record.get("automatic"):
            number = record.get("number")
            if record.get("numbering_unknown") or number is None:
                findings.append(_notice("058", "正文自动编号的当前显示结果无法可靠计算", paragraph))
                current_title = None
            else:
                current_title = _norm(record.get("display_text") or
                                      (str(number) + " " + str(record.get("title", record["text"]))))
        if current_title is not None and entry["label"] != current_title:
            # Layout whitespace around an automatic number is not title content.
            if not (record.get("automatic") and re.sub(r"\s", "", entry["label"]) == re.sub(r"\s", "", current_title)):
                diffs.append("文字/编号应为“" + current_title + "”，缓存为“" + entry["label"] + "”")
        if entry["level"] is not None and entry["level"] != level:
            diffs.append(f"映射目录级别应为{level}，当前为{entry['level']}")
        if diffs:
            description = "；".join(diffs)
            findings.append(make_finding("058", TITLES["058"], "与书签对应正文标题及模板级别映射一致", description, paragraph))
            findings.append(make_finding("066", TITLES["066"], "目录缓存应与当前正文来源同步",
                                         description + "；确认当前不同步，未更新域，不能推定用户操作历史", paragraph))
    if incomplete:
        findings.append(_notice("057", "存在无唯一可用书签的条目，不能可靠判定漏收或同名标题关联", anchor))
    else:
        for element, (record, level) in wanted.items():
            if element not in linked:
                findings.append(make_finding("057", TITLES["057"], f"模板来源应收录映射级别{level}的标题",
                                             "目录未包含：" + record["text"], record["paragraph"]))
    return findings


def check_toc_errors(baseline, ctx, mapping=None):
    template = baseline.get("toc")
    if template is None:
        return [_notice("056", "没有提取到目录模板基准")]
    actual = _snapshot(ctx)
    anchor = actual["titles"][0] if actual["titles"] else (ctx.paragraphs[0] if ctx.paragraphs else None)
    findings = []
    if template["title_formats"]:
        if len(actual["titles"]) == 1:
            findings.extend(_compare_format("064.1", template["title_formats"], actual["titles"][0],
                                            codes=_TITLE_FORMAT_CODES))
            findings.append(_notice("064", "已核对可读取的段落格式；目录标题在实际页面中的绝对位置需要Word排版才能确认",
                                    actual["titles"][0]))
        else:
            findings.append(_notice("064", "目录标题缺失、位于不可读对象中或存在多个候选，不能唯一定位", anchor))
    template_groups, groups = template["groups"], actual["groups"]
    if not template_groups:
        if template["titles"] or template["manual"]:
            findings.append(_notice("056", "模板未提供可识别TOC自动域，不能确定自动生成要求", anchor))
        return findings
    if any(g["field"].get("malformed") for g in template_groups) or template["errors"]:
        return findings + [_notice(rule, "模板目录域结构不完整或域嵌套异常", anchor) for rule in TITLES if rule != "064"]
    if any(g["field"].get("malformed") for g in groups) or actual["errors"]:
        return findings + [_notice(rule, "待检测目录域结构不完整或域嵌套异常", anchor) for rule in TITLES if rule != "064"]
    if not groups:
        if ctx.has_textboxes:
            return findings + [_notice("056", "未发现正文TOC域，但文本框可能包含目录", anchor)]
        return findings + [make_finding("056.1", _title("056.1"), "模板采用原生TOC自动目录域",
                             "未找到对应TOC域；目录缺失或为普通文字", anchor)]
    if len(template_groups) != 1 or len(groups) != 1:
        return findings + [_notice(rule, "存在多个目录，无法唯一匹配各自来源与范围", anchor) for rule in TITLES if rule != "064"]
    template_group, group = template_groups[0], groups[0]
    if template_group["options"]["problems"] or group["options"]["problems"]:
        findings.append(_notice("060", "；".join(template_group["options"]["problems"] + group["options"]["problems"]), anchor))
    elif _source_signature(template_group["options"]) != _source_signature(group["options"]):
        findings.append(make_finding("060", TITLES["060"], template_group["field"]["instruction"], group["field"]["instruction"], anchor))
    else:
        # 060.1/.2：目录显示的层级范围应与模板一致（读 TOC 域的 \o 开关，不需要 Word）。
        expected_levels = template_group["options"].get("levels")
        actual_levels = group["options"].get("levels")
        if expected_levels and actual_levels and expected_levels[1] != actual_levels[1]:
            if actual_levels[1] > expected_levels[1]:
                findings.append(make_finding(
                    "060.1", _title("060.1"),
                    f"目录显示到第 {expected_levels[1]} 级标题（模板 \\o \"{expected_levels[0]}-{expected_levels[1]}\"）",
                    f"当前目录显示到第 {actual_levels[1]} 级（\\o \"{actual_levels[0]}-{actual_levels[1]}\"）",
                    anchor))
            else:
                findings.append(make_finding(
                    "060.2", _title("060.2"),
                    f"目录显示到第 {expected_levels[1]} 级标题（模板 \\o \"{expected_levels[0]}-{expected_levels[1]}\"）",
                    f"当前目录只显示到第 {actual_levels[1]} 级（\\o \"{actual_levels[0]}-{actual_levels[1]}\"）",
                    anchor))
    findings.extend(_check_links(template_group, group, ctx, mapping))
    seen_targets = set()
    seen_entry_format = set()
    for entry in group["entries"]:
        paragraph = entry["paragraph"]
        profiles = template["formats"].get(entry["level"])
        entry_findings = []
        if profiles:
            entry_findings.extend(_compare_format("061.1", profiles, paragraph, True,
                                                  codes=_ENTRY_FORMAT_CODES))
            if entry["separator"] == "\t":
                entry_findings.extend(_check_tabs(
                    profiles, entry, _page_expected(template_group["options"], entry["level"])))
            else:
                entry_findings.extend(_notice(
                    rule, "使用非制表符页码分隔，无法从制表位确认引导线和对齐", paragraph)
                                      for rule in ("062", "063"))
        else:
            entry_findings.append(_notice("061", "模板没有该目录级别的唯一条目格式样例", paragraph))
            if _page_expected(template_group["options"], entry["level"]):
                entry_findings.extend(_notice(
                    rule, "模板没有该目录级别的缓存样例，无法确认页码制表位基准", paragraph)
                                      for rule in ("062", "063"))
        style = paragraph.style.style_id if paragraph.style is not None else ""
        for item in entry_findings:
            key = (item.rule_id, item.status, entry["level"], style,
                   str(item.expected), str(item.actual))
            if key in seen_entry_format:
                continue
            seen_entry_format.add(key)
            findings.append(item)
        if not entry["label"] and entry["targets"]:
            findings.append(make_finding("065.2", _title("065.2"), "目录条目应具有可见标题文字", "条目有目标但标题缓存为空", paragraph))
        if len(entry["targets"]) == 1:
            target = next(iter(entry["targets"]))
            if target in seen_targets:
                findings.append(make_finding("065.3", _title("065.3"), "同一目录正文目标不重复收录",
                                             "多条目录条目指向同一书签：" + target, paragraph))
            seen_targets.add(target)
        if "\n" in entry["text"]:
            findings.append(_notice("065.4", "条目含手工换行；模板未明确禁止，无法据此判定错误换行", paragraph))
    if any(_page_expected(group["options"], entry["level"]) for entry in group["entries"]) or not group["entries"]:
        findings.append(_notice("063", "已按可确认基准核对制表位配置；页码渲染后的实际右边缘和换行位置需要Word排版确认", anchor))
        findings.append(_notice("059", "python-docx不能重新分页并计算正文实际显示页码；不以域缓存当作实际页码", anchor))
        findings.append(_notice("066", "未在Word中刷新目录，无法验证页码缓存是否同步；标题缓存差异另行报告", anchor))
    # 065 的孤立条目/错误换行/跨页分离已由 Word 排版流程判定（word_effective.pagination），
    # 这里只在没有排版事实时保留限制说明。
    if not getattr(ctx, "render_facts", None):
        findings.append(_notice("065", "未取得Word排版结果：目录孤立条目、错误换行与条目页码跨页分离无法判定", anchor))
    return findings

def _snapshot(ctx):
    """带缓存的入口：同一份文档只扫描一次（见 core.docx_context.memo）。"""

    return memo(ctx, "toc_snapshot", lambda: __snapshot_impl(ctx))
