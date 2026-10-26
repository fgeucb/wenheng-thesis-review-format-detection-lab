"""参考文献检测（111–121）的取证实现（只读，以模板为准）。

112/116/119: 本条依赖大模型语义，本阶段暂不实现。
120 and the first-citation-order branch of 121: 待开会讨论，本版本暂不开发。
This parser does not infer true source type, author identities or internet validity.
It only compares unambiguous, same-marker/same-language written structures.
"""
from __future__ import annotations

import re
from collections import defaultdict

from thesis_format_checker.core.xmlutil import qn, visible_text

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, make_finding, unknown
from thesis_format_checker.detectors.fields import scan_fields, scan_body_fields
from thesis_format_checker.detectors.headings import _heading_level, _toc_paragraph_ids


RULES = ("111", "112", "113", "114", "115", "117", "118", "121")
_TITLES = {
    "111": "参考文献著录结构与模板不符", "113": "参考文献作者著录格式不符",
    "112": "文献类型标识缺失或使用不当",
    "114": "参考文献必填字段缺失", "115": "期刊年卷期页码结构不符",
    "117": "DOI、URL或引用日期格式待核对", "118": "参考文献结构标点不符",
    "121": "参考文献排列顺序不符",
}
_FIELD_NAMES = {
    "authors": "作者", "title": "题名", "outlet": "期刊名", "publisher": "出版社",
    "place": "出版地", "year": "年份", "volume": "卷", "issue": "期", "pages": "页码",
    "doi": "DOI", "url": "URL", "access_date": "引用日期",
}
_ALIASES = {
    "作者": "authors", "题名": "title", "标题": "title", "期刊名": "outlet", "刊名": "outlet",
    "出版社": "publisher", "出版地": "place", "年份": "year", "出版年": "year",
    "年": "year", "卷号": "volume", "卷": "volume", "期号": "issue", "期": "issue",
    "起止页码": "pages", "页码": "pages", "doi": "doi", "url": "url", "引用日期": "access_date",
}
_MARKER = re.compile(r"[\[［]([A-Za-z]+(?:/[A-Za-z]+)?)[\]］]")
_NUMBER = re.compile(r"^(?:(?P<left>[\[［（(])(?P<n>\d+)(?P<right>[\]］）)])|(?P<plain>\d+)(?P<suffix>[.．、]))(?P<gap>\s*)")
_YEAR = re.compile(r"(?<!\d)(?:\d{4}|年份|出版年)(?!\d)")
_VALUE = r"(?:\d+|卷号|卷|期号|期|起止页码|页码)"
_PLACEHOLDERS = {"作者", "作者姓名", "题名", "标题", "期刊名", "刊名", "出版社", "出版地",
                 "年份", "出版年", "卷", "卷号", "期", "期号", "页码", "起止页码", "待填", "待填写", "xxx", "…", "..."}


def _unavailable(rule, reason, paragraph=None):
    return unknown(rule, _title(rule), UNREADABLE + "；" + reason, paragraph)


#: 参考文献字段 → 细粒度规则编号（114）
_FIELD_CODES = {"title": "114.1", "outlet": "114.2", "publisher": "114.3", "place": "114.4",
                "year": "114.5", "volume": "114.6", "issue": "114.7", "pages": "114.8"}
#: DOI/URL/引用日期 → 细粒度规则编号（117）
_DIGITAL_CODES = {"doi": "117.1", "url": "117.2", "access_date": "117.3", "DOI": "117.1",
                  "URL": "117.2", "引用日期": "117.3"}


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else _TITLES[code]


def _text(paragraph):
    # Include hyperlink text without normalising characters or changing the XML.

    return visible_text(paragraph._p)


def _heading_text(text):
    text = re.sub(r"\s+", "", text).casefold()
    return re.sub(r"^(?:第[一二三四五六七八九十百\d]+章|\d+[.、．]?)", "", text)


def _bibliography_paragraphs(ctx):
    toc = _toc_paragraph_ids(ctx)
    active, level, anchor = False, None, None
    result = []
    for paragraph in ctx.paragraphs:
        if paragraph._p in toc:
            continue
        text = _text(paragraph)
        compact = _heading_text(text)
        p_level = _heading_level(paragraph)
        if compact in {"参考文献", "references", "bibliography"}:
            active, level, anchor = True, p_level, paragraph
            continue
        back_heading = re.match(r"^(?:致谢|附录|攻读学位期间|acknowledg(?:e)?ments|appendix)(?:$|[a-z一二三四五六七八九十\d：:])", compact)
        if active and (back_heading or (p_level and (level is None or p_level <= level))):
            active = False
        if active and text.strip():
            result.append(paragraph)
    return result, anchor


def _language(authors):
    cjk = bool(re.search(r"[\u3400-\u9fff]", authors))
    latin = bool(re.search(r"[A-Za-z]", re.sub(r"\bet\s+al\.?", "", authors, flags=re.I)))
    return "zh" if cjk and not latin else "en" if latin and not cjk else None


def _authors(value):
    omitted = re.search(r"(?:\bet\s+al\.?|等)\s*$", value, re.I)
    marker = omitted.group().strip() if omitted else None
    visible = value[:omitted.start()].rstrip(" ,，;；、") if omitted else value
    parts = re.split(r"\s*[,，;；、]\s*", visible)
    separators = tuple(re.findall(r"[,，;；、]\s*", visible))
    language = _language(visible)
    # A comma inside a Western surname/given-name form is ambiguous.
    safe = bool(parts) and all(parts) and (language == "zh" or (
        language == "en" and all(len(p.split()) >= 2 for p in parts)))
    return {"text": value, "marker": marker, "parts": parts, "separators": separators,
            "language": language, "safe_count": safe}


def _digital_fields(text):
    doi = re.search(r"(?i)(?P<label>DOI\s*[:：]\s*|https?://(?:dx\.)?doi\.org/)(?P<value>[^\s，；<>]+)", text)
    if doi and re.match(r"https?://|URL[:：]|引用日期|访问日期", doi.group("value"), re.I):
        # A blank DOI label followed by the next URL/date field must not absorb
        # that following field and become a spurious malformed DOI value.
        doi = None
    links = list(re.finditer(r"https?://[^\s<>]+", text, re.I))
    url = next((m for m in links if not re.match(r"https?://(?:dx\.)?doi\.org/", m.group(), re.I)), None)
    date = re.search(r"(?P<open>[\[［])(?P<year>\d{4})(?P<s1>[-/\.])(?P<month>\d{1,2})(?P<s2>[-/\.])(?P<day>\d{1,2})(?P<close>[\]］])", text)
    positions = [m.start() for m in (doi, url, date) if m]
    label = re.search(r"\bDOI\s*[:：]|\bURL\s*[:：]|引用日期\s*[:：]|访问日期\s*[:：]", text, re.I)
    if label:
        positions.append(label.start())
    return {
        "doi": doi.group("value").rstrip(".。;；") if doi else None,
        "doi_style": doi.group("label") if doi else None,
        "url": url.group().rstrip(".。;；，,]］") if url else None,
        "url_style": url.group().split(":", 1)[0] + "://" if url else None,
        "access_date": date.group() if date else None,
        "date_style": (date.group("open"), date.group("s1"), date.group("s2"), date.group("close"),
                       len(date.group("month")), len(date.group("day"))) if date else None,
        "date_components": (date.group("year"), date.group("month"), date.group("day")) if date else None,
        "hints": {
            "doi": bool(re.search(r"\bdoi\b|doi\.org", text, re.I)),
            "url": bool(re.search(r"\bURL\b|https?(?=[:/])|www\.", re.sub(r"https?://(?:dx\.)?doi\.org/\S+", "", text, flags=re.I), re.I)),
            "access_date": bool(re.search(r"引用日期|访问日期|[\[［]\d{4}[^\]］]*[\]］]", text)),
        },
        "start": min(positions) if positions else len(text),
    }


def _parse_entry(paragraph):
    text = _text(paragraph).strip()
    number = _NUMBER.match(text)
    entry = {"paragraph": paragraph, "text": text, "number": None, "number_style": None,
             "error": None, "fields": {}, "separators": {}, "digital": _digital_fields(text)}
    if number:
        entry["number"] = int(number.group("n") or number.group("plain"))
        entry["number_style"] = (number.group("left"), number.group("right"), number.group("suffix"), number.group("gap"))
        text = text[number.end():]
    marker = _MARKER.search(text)
    if marker is None:
        entry["error"] = "未找到可区分类型示例的文献类型标识；不推断真实文献类型"
        return entry
    entry["type"] = marker.group(1).upper()
    entry["type_style"] = (marker.group()[0], marker.group()[-1])
    lead = text[:marker.start()]
    stops = list(re.finditer(r"[.。．]\s*", lead))
    if len(stops) != 1:
        entry["error"] = "作者与题名的句点边界不唯一（缩写、机构名或复杂题名需要人工核对）"
        return entry
    split = stops[0]
    authors = _authors(lead[:split.start()].strip())
    title = lead[split.end():].strip()
    entry["authors"] = authors
    entry["language"] = authors["language"]
    entry["fields"].update(authors=authors["text"], title=title)
    entry["separators"]["作者→题名"] = split.group()
    after = text[marker.end():]
    boundary = re.match(r"[.。．,，;；:\s]*", after).group()
    entry["separators"]["类型标识→出版项"] = boundary
    after = after[len(boundary):]
    digital = _digital_fields(after)
    tail = after[:digital["start"]].rstrip()
    year_matches = list(_YEAR.finditer(tail))
    entry["numeric_signature"] = None
    entry["numeric_error"] = None
    if len(year_matches) != 1:
        entry["numeric_error"] = "无法唯一识别出版年份，不能把未解析字段当作缺失"
        entry["fields"].update(outlet=None, publisher=None, place=None, year=None, volume=None, issue=None, pages=None)
        return entry
    year = year_matches[0]
    # Identify the publication/numeric boundary, independently from later number order.
    # Digits inside a journal/publisher name are content, not volume numbers.
    # Only a recognisable volume(issue), prefix immediately before the year is
    # independently identified as a reordered numeric group.
    swapped = re.search(r"[,，;；]\s*(?P<group>" + _VALUE + r"\s*[（(]" + _VALUE + r"[）)][,，;；]\s*)$", tail[:year.start()])
    entry["numeric_start"] = swapped.start("group") if swapped else year.start()
    publication = tail[:entry["numeric_start"]]
    separator = re.search(r"[，,;；:\s]+$", publication)
    entry["separators"]["出版项→数字字段"] = separator.group() if separator else ""
    publication = publication[:separator.start()] if separator else publication
    entry["fields"]["outlet"] = publication.strip()
    place_publisher = re.fullmatch(r"(.*?)([:：]\s*)(.*)", publication)
    entry["fields"].update(place=None, publisher=None)
    if place_publisher:
        entry["fields"]["place"] = place_publisher.group(1).strip()
        entry["fields"]["publisher"] = place_publisher.group(3).strip()
        entry["separators"]["出版地→出版社"] = place_publisher.group(2)
    entry["fields"]["year"] = year.group()
    numeric = tail[entry["numeric_start"]:]
    signature = numeric[:year.start() - entry["numeric_start"]] + "〈年份〉" + numeric[year.end() - entry["numeric_start"]:]
    signature = re.sub(_VALUE, "〈数值〉", signature)
    entry["numeric_signature"] = signature
    # These field labels are recognised only for a canonical readable position;
    # arbitrary number sequences remain unparsed instead of guessing identities.
    following = tail[year.end():]
    shape = re.fullmatch(r"(?P<sep>[,，;；]\s*)(?P<volume>" + _VALUE + r")\s*(?P<open>[（(])(?P<issue>" + _VALUE + r")(?P<close>[）)])(?P<page_sep>[:：]\s*)(?P<pages>" + _VALUE + r"(?:\s*[-–—－~～]\s*" + _VALUE + r")?)(?P<end>[.。．]?\s*)", following)
    entry["fields"].update(volume=None, issue=None, pages=None)
    if shape:
        entry["fields"].update(volume=shape.group("volume"), issue=shape.group("issue"), pages=shape.group("pages"))
    # Explicit empty slots can be identified without assigning a guessed value.
    empties = re.fullmatch(r"[,，;；]\s*(?P<volume>" + _VALUE + r")?\s*[（(](?P<issue>" + _VALUE + r")?[）)][:：]\s*(?P<pages>" + _VALUE + r"(?:\s*[-–—－~～]\s*" + _VALUE + r")?)?[.。．]?", following)
    if empties:
        entry["fields"].update({key: empties.group(key) or "" for key in ("volume", "issue", "pages")})
    return entry


def _is_instruction(text):
    # A paper title may itself discuss a standard or formatting. Only recognise
    # instruction-shaped paragraphs, never a clearly numbered ordinary entry.
    text = text.strip()
    prefix = _NUMBER.match(text)
    if prefix:
        remainder = text[prefix.end():]
        # Word templates often number their instructions. These explicit forms
        # remain instructions; a numbered citation's title is never inspected
        # as a rule declaration.
        if re.match(r"(?:[\[［][A-Za-z]+(?:/[A-Za-z]+)?[\]］]\s*(?:必填|必须包含|应包含)|(?:中文|英文)?作者(?:人数)?超过\d+人|参考文献按)", remainder):
            text = remainder
        else:
            return False
    marker = _MARKER.search(text)
    if marker and re.search(r"[.。．]", text[:marker.start()]):
        return False
    return bool(re.search(r"必填|必须包含|应包含|按.*排序|按.*排列|作者超过|作者人数|列前|列出前|GB/T|著录要求|格式要求|参考文献要求", text, re.I))


def _instruction_rules(paragraphs):
    required, authors, sorting = defaultdict(set), {}, set()
    for paragraph in paragraphs:
        text = _text(paragraph).strip()
        marker = _MARKER.search(text)
        fields = re.search(r"(?:必填字段|必填|必须包含|应包含)\s*[:：]?\s*(.+)", text)
        if marker and fields:
            # Required-field declarations use explicit field lists. No common
            # national-standard field set is substituted here.
            for item in re.split(r"[、,，;；和及\s]+", fields.group(1).rstrip("。.")):
                if item.casefold() in _ALIASES:
                    required[marker.group(1).upper()].add(_ALIASES[item.casefold()])
        policy = re.search(r"作者(?:人数)?超过\s*(\d+)\s*人.*?列(?:出)?前\s*(\d+)\s*人.*?(?:后加|加|使用)\s*[“\"‘]?\s*(等|et\s+al\.?)", text, re.I)
        if policy:
            key = "en" if re.search(r"英文|english|et\s+al", text, re.I) else "zh"
            rule = (int(policy.group(1)), int(policy.group(2)), policy.group(3))
            authors.setdefault(key, set()).add(rule)
        if re.search(r"按(?:第一)?作者(?:的)?(?:英文)?字母(?:升序|顺序)?(?:排序|排列)", text):
            sorting.add("author")
        if re.search(r"按(?:文末)?(?:条目)?编号(?:升序|顺序|递增)?(?:排序|排列)", text):
            sorting.add("number")
        if re.search(r"按(?:正文)?(?:首次)?引用(?:出现)?顺序(?:排序|排列)", text):
            sorting.add("citation_deferred")
    return {"required": dict(required), "authors": authors, "sorting": sorting}


def _field_paragraphs(ctx):
    fields = scan_body_fields(ctx)
    affected = set()
    for field in fields:
        if field["kind"] in {"CITATION", "BIBLIOGRAPHY"} or field["malformed"]:
            if field["paragraph_element"] is not None:
                affected.add(field["paragraph_element"])
            affected.update(field["result_paragraph_elements"])
    return affected


def _unsafe(paragraph, field_nodes):
    return paragraph._p in field_nodes or bool(paragraph._p.xpath(
        ".//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo|.//w:rPrChange|.//w:pPrChange|.//w:object|.//w:drawing"
    )) or any(a.tag in {qn("w:ins"), qn("w:del"), qn("w:moveFrom"), qn("w:moveTo")} for a in paragraph._p.iterancestors())


def extract_bibliography_baseline(ctx, mapping=None):
    paragraphs, anchor = _bibliography_paragraphs(ctx)
    policies = _instruction_rules(paragraphs)
    field_nodes = _field_paragraphs(ctx)
    groups = defaultdict(list)
    issues = []
    blocked_groups, blocked_all = set(), False
    for paragraph in paragraphs:
        text = _text(paragraph)
        if _unsafe(paragraph, field_nodes):
            entry = _parse_entry(paragraph)
            if entry.get("type") and entry.get("language"):
                blocked_groups.add((entry["type"], entry["language"]))
            else:
                blocked_all = True
            issues.append("模板含域、修订或不可读对象，无法确定当前采用的文献基准")
            continue
        if _is_instruction(text):
            continue
        entry = _parse_entry(paragraph)
        if entry["error"] or not entry.get("language"):
            issues.append(entry["error"] or "模板作者语言无法唯一归组")
            # Dropping an unreadable counterexample could make the remaining
            # examples look falsely unanimous. Prevent such a false baseline.
            if entry.get("type") and entry.get("language"):
                blocked_groups.add((entry["type"], entry["language"]))
            else:
                blocked_all = True
            continue
        groups[(entry["type"], entry["language"])].append(entry)
    return {"bibliography": {"groups": dict(groups), "policies": policies,
                                     "issues": issues, "present": anchor is not None,
                                     "blocked_groups": blocked_groups, "blocked_all": blocked_all}}


def _one(entries, getter):
    values = [getter(entry) for entry in entries]
    if not values or any(value is None for value in values):
        return None
    return values[0] if all(value == values[0] for value in values) else None


def _display(value):
    return repr(value)


def _check_structure(entry, templates):
    findings, p = [], entry["paragraph"]
    structural = []
    for key, label in (("number_style", "条目前编号形式"), ("type_style", "类型标识括号形式")):
        expected = _one(templates, lambda item: item.get(key))
        if expected is not None and entry.get(key) != expected:
            structural.append((label, expected, entry.get(key)))
    # 用 sorted 固定遍历顺序：集合的迭代顺序依赖进程哈希种子，会让同一份文档的
    # 报告文案时好时坏（对回归比对尤其致命）。
    for key in sorted(set().union(*(t["separators"].keys() for t in templates))):
        expected = _one(templates, lambda item: item["separators"].get(key))
        actual = entry["separators"].get(key)
        if expected is not None and actual is not None and expected != actual:
            structural.append((key, expected, actual))
    if structural:
        expected = "；".join(f"{name}={_display(a)}" for name, a, b in structural)
        actual = "；".join(f"{name}={_display(b)}" for name, a, b in structural)
        findings.append(make_finding("111", _TITLES["111"], expected, actual, p))
        # 118 is reserved for punctuation repertoire changes. Spaces/order
        # alone remain under 111; title/author/internal URL punctuation is ignored.
        punctuation = [(n, a, b) for n, a, b in structural if isinstance(a, str) and isinstance(b, str)
                       and re.sub(r"\s", "", a) != re.sub(r"\s", "", b)]
        if punctuation:
            # 中文文献误用英文标点（118.1）/ 英文文献误用中文标点（118.2）
            code = "118.1" if entry.get("language") == "zh" else "118.2"
            findings.append(make_finding(code, _title(code),
                "；".join(f"{n}={_display(a)}" for n, a, b in punctuation),
                "；".join(f"{n}={_display(b)}" for n, a, b in punctuation), p))
    if entry["type"] == "J":
        expected = _one(templates, lambda item: item.get("numeric_signature"))
        actual = entry.get("numeric_signature")
        if expected is None or actual is None:
            findings.append(_unavailable("115", entry.get("numeric_error") or "模板同类年卷期页数字结构不唯一或未提供", p))
        elif expected != actual:
            # 数字已被掩码成〈数值〉：标点不同 → 115.2，字段结构不同 → 115.1
            punctuation = lambda value: re.findall(r"[,，;；:：.。()（）\[\]［］\-–—~～]", str(value))
            code = "115.2" if punctuation(expected) != punctuation(actual) else "115.1"
            findings.append(make_finding(code, _title(code), expected, actual, p))
    return findings


def _check_authors(entry, templates, policies):
    findings, p = [], entry["paragraph"]
    actual = entry["authors"]
    current = actual["marker"]
    markers = {t["authors"]["marker"] for t in templates if t["authors"]["marker"]}
    policy = policies["authors"].get(entry["language"], set())
    if len(policy) == 1:
        threshold, retain, marker = next(iter(policy))
        if not actual["safe_count"]:
            findings.append(_unavailable("113", "作者分隔存在姓名/机构内部标点歧义，不能推算作者人数", p))
        else:
            count = len(actual["parts"])
            if (current and (count != retain or current.casefold().rstrip(".") != marker.casefold().rstrip("."))) or (not current and count > threshold):
                findings.append(make_finding("113.1", _title("113.1"), f"作者超过{threshold}人时列前{retain}人后加{marker}", actual["text"], p))
    elif len(policy) > 1:
        findings.append(_unavailable("113", "模板同语言作者人数/省略条件相互冲突", p))
    elif current and len(markers) == 1 and current.casefold().rstrip(".") != next(iter(markers)).casefold().rstrip("."):
        findings.append(make_finding("113.3", _title("113.3"), "已有省略标记的写法为" + next(iter(markers)), actual["text"], p))
    delimiters = set()
    for template in templates:
        author = template["authors"]
        if author["safe_count"] and author["separators"]:
            delimiters.update(author["separators"])
    if actual["safe_count"] and len(delimiters) == 1 and any(x not in delimiters for x in actual["separators"]):
        findings.append(make_finding("113.2", _title("113.2"), "作者间分隔符为" + repr(next(iter(delimiters))), actual["text"], p))
    return findings


def _check_required(entry, policies):
    findings, missing, unparsed = [], [], []
    required = policies["required"].get(entry["type"], set())
    if not required:
        return [_unavailable("114", "模板未明确此类型必填字段；示例中出现字段不能独立证明它是必填项", entry["paragraph"])]
    for field in sorted(required - {"doi", "url", "access_date"}):
        value = entry["fields"].get(field)
        if value is None:
            unparsed.append(_FIELD_NAMES[field])
        elif not value.strip() or value.strip().casefold() in _PLACEHOLDERS:
            missing.append(_FIELD_NAMES[field])
    if missing:
        # 一个字段一条结论：题名/期刊名/出版社/出版地/年份/卷/期/页码
        by_code: dict[str, list[str]] = {}
        for field in sorted(required):
            name = _FIELD_NAMES[field]
            if name in missing:
                by_code.setdefault(_FIELD_CODES.get(field, "114.1"), []).append(name)
        for code, names in by_code.items():
            findings.append(make_finding(code, _title(code),
                                         "此类型必填：" + "、".join(_FIELD_NAMES[f] for f in sorted(required)),
                                         "缺失或未替换占位文字：" + "、".join(names), entry["paragraph"]))
    if unparsed:
        findings.append(_unavailable("114", "以下必填字段无法可靠划分，不能认定缺失：" + "、".join(unparsed), entry["paragraph"]))
    return findings


def _check_digital(entry, templates, policies):
    findings, p = [], entry["paragraph"]
    digital = entry["digital"]
    # A hyperlink with a label such as “访问地址” still contains an address.
    # Do not report a missing field merely because the visible text is not a URI.
    for link in p._p.xpath(".//w:hyperlink"):
        relation = p.part.rels.get(link.get(qn("r:id")))
        if relation is not None and relation.is_external:
            target = relation.target_ref
            if re.match(r"https?://(?:dx\.)?doi\.org/", target, re.I):
                digital["hints"]["doi"] = True
            elif re.match(r"https?://", target, re.I):
                digital["hints"]["url"] = True
    required = policies["required"].get(entry["type"], set()) & {"doi", "url", "access_date"}
    unreadable = [key for key in sorted(required) if not digital[key] and digital["hints"][key]]
    missing = [key for key in sorted(required) if not digital[key] and not digital["hints"][key]]
    if missing:
        for key in missing:
            code = _DIGITAL_CODES.get(key, "117.1")
            findings.append(make_finding(code, _title(code),
                                         "此类型明确必填：" + _FIELD_NAMES[key],
                                         "未找到对应字段或标记：" + _FIELD_NAMES[key], p))
    if unreadable:
        findings.append(_unavailable("117", "以下字段的文字/标签存在，但结构未能可靠解析，不能认定缺失：" + "、".join(_FIELD_NAMES[k] for k in unreadable), p))
    differences = []
    for key, prop in (("doi", "doi_style"), ("url", "url_style"), ("access_date", "date_style")):
        matching = [t for t in templates if t["digital"][key]]
        expected = _one(matching, lambda item: item["digital"][prop])
        actual_style = digital[prop]
        if key == "access_date" and expected is not None and actual_style is not None:
            # Brackets and delimiters are direct characters. A two-digit month
            # like 12 is not evidence of a zero-padding requirement. Need two
            # distinct template date examples actually exhibiting each zero.
            widths = []
            for offset, component_index in ((4, 1), (5, 2)):
                padded = {t["digital"]["access_date"] for t in matching
                          if t["digital"]["date_components"][component_index].startswith("0")
                          and len(t["digital"]["date_components"][component_index]) == 2}
                widths.append(len(padded) >= 2)
            expected = expected[:4] + tuple(expected[i] if widths[i - 4] else None for i in (4, 5))
            actual_style = actual_style[:4] + tuple(actual_style[i] if widths[i - 4] else None for i in (4, 5))
        if digital[key] and expected is not None and actual_style != expected:
            differences.append((_FIELD_NAMES[key], expected, actual_style))
    if differences:
        for name, want, got in differences:
            code = _DIGITAL_CODES.get(name, "117.1")
            findings.append(make_finding(code, _title(code), f"{name}={_display(want)}",
                                         f"{name}={_display(got)}", p))
    if digital["doi"]:
        # This grammar is activated by matching template examples; it does not
        # substitute a fixed DOI registrar prefix or validate registration.
        doi_examples = [t["digital"]["doi"] for t in templates if t["digital"]["doi"]]
        shape = r"\d+\.\d+/[^\s/].*"
        if doi_examples and all(re.fullmatch(shape, value) for value in doi_examples) and not re.fullmatch(shape, digital["doi"]):
            findings.append(make_finding("117.1", _title("117.1"), "模板DOI为数字段.数字段/非空后缀", digital["doi"], p))
    # URI/DOI content is deliberately not treated as a normative standard.
    # Compare structural syntax to the template; never visit a link or invent
    # an access date, expiry threshold, source identity, or source metadata.
    if digital["doi"] or digital["url"]:
        findings.append(_unavailable("117", "本阶段只比对可读写法；未联网验证DOI/URL可达性或目标真实性", p))
    for link in p._p.xpath(".//w:hyperlink"):
        relation_id = link.get(qn("r:id"))
        if not relation_id:
            continue
        relation = p.part.rels.get(relation_id)
        if relation is None or not relation.is_external:
            findings.append(_unavailable("117", "超链接关系缺失或不是可读取的外部地址", p))
            continue
        shown = "".join(node.text or "" for node in link.iter(qn("w:t")))
        if re.fullmatch(r"https?://\S+", shown) and shown != relation.target_ref:
            # No assertion that either target is factually incorrect.
            findings.append(_unavailable("117", "可见URL与超链接实际目标不同，模板不能证明哪个是正确来源", p))
    return findings


def _check_order(entries, policies, anchor):
    mode = policies["sorting"] - {"citation_deferred"}
    if not mode:
        return [] if "citation_deferred" in policies["sorting"] else [_unavailable("121", "模板未明确作者字母或文末编号排序方式，不能仅凭短示例推断", anchor)]
    if len(mode) != 1 or "citation_deferred" in policies["sorting"]:
        return [_unavailable("121", "模板排序方式存在冲突，无法选定唯一规则", anchor)]
    mode = next(iter(mode))
    findings, previous = [], None
    for entry in entries:
        if entry is None or entry["error"]:
            previous = None
            continue
        if mode == "number":
            key = entry["number"]
            if key is None:
                findings.append(_unavailable("121", "手工条目编号缺失或自动列表编号不可读取", entry["paragraph"]))
                previous = None
                continue
        else:
            author = entry["authors"]
            if author["language"] != "en" or not author["safe_count"] or not re.fullmatch(r"[A-Za-z \-]+", author["parts"][0]):
                findings.append(_unavailable("121", "作者排序键包含中文、标点、机构/姓名歧义；模板未给出转写/特殊字符排序细则", entry["paragraph"]))
                previous = None
                continue
            # Plain Latin first-author keys only. No transliteration, accents,
            # same-name/year tiebreak or guessed personal-name rearrangement.
            key = author["parts"][0].casefold()
        if previous is not None and key < previous[0]:
            findings.append(make_finding("121", _TITLES["121"], "按" + ("文末编号升序" if mode == "number" else "作者字母顺序") + "排列",
                f"当前排序键{key!r}位于较大排序键{previous[0]!r}之后", entry["paragraph"]))
        previous = key, entry
    return findings


def check_bibliography_errors(baseline, ctx, mapping=None):
    config = baseline.get("bibliography", {})
    groups = config.get("groups", {})
    policies = config.get("policies", {"required": {}, "authors": {}, "sorting": set()})
    paragraphs, anchor = _bibliography_paragraphs(ctx)
    if not paragraphs:
        return []
    field_nodes = _field_paragraphs(ctx)
    entries, findings = [], []
    allowed_types = {key[0] for key in groups if key and key[0]}
    for paragraph in paragraphs:
        if _is_instruction(_text(paragraph)):
            continue
        if _unsafe(paragraph, field_nodes):
            entries.append(None)
            for rule in RULES:
                findings.append(_unavailable(rule, "条目含引用/书目域、修订或不可读对象；不能从缓存或图片推断著录字段", paragraph))
            continue
        entry = _parse_entry(paragraph)
        entries.append(entry)
        if entry.get("type") is None:
            findings.append(make_finding("112.1", _title("112.1"),
                                         "每条文献应带类型标识（如 [J]/[M]/[C]/[D]）",
                                         "条目前缺少文献类型标识", paragraph))
        elif allowed_types and entry.get("type") not in allowed_types:
            findings.append(make_finding("112.2", _title("112.2"),
                                         "文献类型标识应使用模板出现过的：" + "、".join(sorted(allowed_types)),
                                         f"当前使用 [{entry['type']}]", paragraph))
        templates = groups.get((entry.get("type"), entry.get("language")), [])
        if config.get("blocked_all") or (entry.get("type"), entry.get("language")) in config.get("blocked_groups", set()):
            for rule in RULES:
                findings.append(_unavailable(rule, "模板相关示例含不可读/未解析内容，不能丢弃这些示例后将剩余条目当作唯一基准", paragraph))
            entries[-1] = None
            continue
        if entry["error"] or not templates:
            reason = entry["error"] or "模板没有可唯一匹配的同标识、同作者语言示例"
            for rule in RULES[:-1]:
                findings.append(_unavailable(rule, reason, paragraph))
            continue
        findings.extend(_check_structure(entry, templates))
        findings.extend(_check_authors(entry, templates, policies))
        findings.extend(_check_required(entry, policies))
        findings.extend(_check_digital(entry, templates, policies))
    findings.extend(_check_order(entries, policies, anchor))
    return findings
