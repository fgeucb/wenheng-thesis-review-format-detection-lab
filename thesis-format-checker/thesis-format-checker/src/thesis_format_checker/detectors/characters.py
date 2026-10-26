"""字符、空格与日期形式检测（122–132）的取证实现。

122/123：语义扩展本阶段暂不实现，保留原检测；不重复已有标点检测。
126/127/130/131/132：本条依赖大模型语义，本阶段暂不实现。
Reference-identification business remains deferred; no bibliography inference.
"""
from __future__ import annotations

from collections import defaultdict
import re

from thesis_format_checker.core.xmlutil import qn, visible_text

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, effective_paragraph, make_finding, unknown
from thesis_format_checker.detectors.paragraphs import _body_records, _has_revision
from thesis_format_checker.detectors.headings import _heading_level


TITLES = {"124": "全半角字符形式不符", "125": "中英文或数字边界空格不符",
          "126": "空格使用不规范", "127": "英文与数字空格使用不规范",
          "128": "连续空行与模板不符", "129": "日期时间年份或编号字符形式不符",
          "130": "计量单位写法不规范", "131": "符号使用不规范", "132": "文档包含乱码或不可见字符"}
# Unicode counterparts are character identities, not prescribed school formats.
_PUNCT = {",": "逗号", "，": "逗号", ":": "冒号", "：": "冒号", ";": "分号", "；": "分号",
          "!": "叹号", "！": "叹号", "?": "问号", "？": "问号",
          "(": "左圆括号", "（": "左圆括号", ")": "右圆括号", "）": "右圆括号",
          "[": "左方括号", "［": "左方括号", "]": "右方括号", "］": "右方括号"}
_BLANKS = " \u00a0\u3000\u2009\u202f\t"
_CJK = r"\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
_LETTERS = "A-Za-zＡ-Ｚａ-ｚ"
_DIGITS = "0-9０-９"
_LINK = re.compile(r"(?:https?://|www\.|doi\s*:|10\.\d{4,9}/)\S+|[\w.+-]+@[\w.-]+", re.I)
_DATE_LABEL = re.compile(r"^\s*([^:：\d０-９\n]{0,24}?(?:日期|时间|年份|年度|编号))(?:\s*[:：]\s*|\s+|(?=[0-9０-９]))(.*?)\s*$")
_SPECIAL_STYLE = re.compile(r"code|代码|程序|verbatim|公式|equation|参考文献|bibliography|reference", re.I)
_BACK_HEAD = re.compile(r"^(?:参考文献|references|bibliography)\s*$", re.I)
_FIELD_TAGS = ".//w:fldChar|.//w:fldSimple|.//w:instrText"


def _unknown(rule, reason, paragraph=None, unreadable=False):
    if unreadable:
        reason = UNREADABLE + "；" + reason
    return unknown(rule, _title(rule), reason, paragraph)


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else TITLES[code]


def _original_text(p):
    """Keep exact code points; no strip/normalization is applied to evidence."""

    return visible_text(p._p, page_breaks_as_formfeed=True)


def _excluded(p):
    names, style = [], p.style
    while style is not None and style.style_id not in names:
        names.append(style.style_id)
        if _SPECIAL_STYLE.search(style.name or ""):
            return True
        style = style.base_style
    text = _original_text(p)
    return bool(_LINK.search(text) or p._p.xpath(".//w:hyperlink|.//m:oMath|.//w:object|.//w:drawing|.//w:pict"))


def _body(ctx, mapping, side):
    return [r for r in _body_records(ctx, mapping, side) if not _excluded(r["paragraph"])]


def _kind(ch):
    if re.fullmatch("[" + _CJK + "]", ch):
        return "中文"
    if re.fullmatch("[" + _LETTERS + "]", ch):
        return "英文"
    if re.fullmatch("[" + _DIGITS + "]", ch):
        return "数字"
    return None


def _width_samples(text):
    result = defaultdict(list)
    for i, char in enumerate(text):
        if "A" <= char <= "Z" or "Ａ" <= char <= "Ｚ":
            key = "大写字母"
        elif "a" <= char <= "z" or "ａ" <= char <= "ｚ":
            key = "小写字母"
        elif "0" <= char <= "9" or "０" <= char <= "９":
            key = "数字"
        elif char in _PUNCT:
            left = next((c for c in reversed(text[:i]) if c not in _BLANKS), "")
            right = next((c for c in text[i + 1:] if c not in _BLANKS), "")
            # Only matching linguistic contexts; numeric colon is time/ratio,
            # parentheses around numbers may be citations or equation labels.
            if _kind(left) == _kind(right) == "中文":
                context = "中文语境"
            elif _kind(left) == _kind(right) == "英文":
                context = "英文语境"
            else:
                continue
            key = context + _PUNCT[char]
        else:
            continue
        result[key].append(("全角" if 0xFF01 <= ord(char) <= 0xFF5E else "半角", i, char))
    return result


def _boundary_samples(text):
    result = defaultdict(list)
    pattern = re.compile("([" + _CJK + _LETTERS + _DIGITS + "])([" + re.escape(_BLANKS) + "]*)(?=[" + _CJK + _LETTERS + _DIGITS + "])")
    for match in pattern.finditer(text):
        end = match.end()
        first, second = _kind(match.group(1)), _kind(text[end])
        if {first, second} not in ({"中文", "英文"}, {"中文", "数字"}):
            continue
        result[first + "→" + second].append((match.group(2), match.start(), text[match.start():end + 1]))
    return result


#: 中西文/数字交界处 → 中文说法
_BOUNDARY_LABELS = {"中文→英文": "中文与英文之间", "中文→数字": "中文与数字之间",
                    "英文→中文": "英文与中文之间", "数字→中文": "数字与中文之间",
                    "英文→数字": "英文与数字之间", "数字→英文": "数字与英文之间"}


def _boundary_label(key: str) -> str:
    return _BOUNDARY_LABELS.get(key, key.replace("→", "与") + "之间")


def _space_description(value) -> str:
    """空格的中文说明：只写“几个半角/全角空格”，不写 U+ 码位。"""

    if not value:
        return "没有空格"
    parts = []
    for char, name in ((" ", "半角空格"), ("\u3000", "全角空格"), ("\t", "制表符")):
        count = sum(1 for item in value if item == char)
        if count:
            parts.append(f"{count} 个{name}")
    return "有 " + "＋".join(parts) if parts else f"有 {len(value)} 个空白字符"


def _auto_spacing_text(flags: dict) -> str:
    """Word 自动间距设置的中文说明（不写字段名与 True/False）。"""

    def state(key, label):
        value = flags.get(key)
        if value is None:
            return label + "无法读取"
        return label + ("已开启" if value else "已关闭")

    return ("Word 自动间距：" + state("autoSpaceDE", "中文与西文间距调整")
            + "；" + state("autoSpaceDN", "中文与数字间距调整"))


def _auto_spacing(p):
    result = {}
    props = [p._p.pPr]
    seen, style = set(), p.style
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        props.append(style.element.pPr)
        style = style.base_style
    props.extend(p.part.document.styles.element.xpath("./w:docDefaults/w:pPrDefault/w:pPr"))
    for name in ("autoSpaceDE", "autoSpaceDN"):
        result[name] = next((n.get(qn("w:val"), "1") not in {"0", "false", "off"}
                            for prop in props if prop is not None
                            for n in prop.findall(qn("w:" + name))), True)
    return result


def _instruction(p):
    return " ".join([n.text or "" for n in p._p.iter(qn("w:instrText"))] +
                    [n.get(qn("w:instr"), "") for n in p._p.iter(qn("w:fldSimple"))])


def _field_records(ctx):
    """Recognize explicit local labels, never infer dates from free decimals."""
    result, in_references = [], False
    for p in ctx.paragraphs:
        text = _original_text(p)
        if _BACK_HEAD.fullmatch(text.strip()):
            in_references = True
            continue
        if in_references:
            if _heading_level(p) and re.fullmatch(r"(?:致谢|附录(?:\s*[A-Za-zＡ-Ｚ0-9一二三四五六七八九十]+)?|acknowledg(?:e)?ments|appendix(?:\s+[A-Za-z0-9]+)?)", text.strip(), re.I):
                in_references = False
            else:
                continue
        if _excluded(p):
            continue
        match = _DATE_LABEL.fullmatch(text)
        if match:
            label, value = match.groups()
            if not value:
                continue
            result.append({"paragraph": p, "key": (label.strip(), p.style.name), "label": label.strip(),
                           "value": value, "field": bool(p._p.xpath(_FIELD_TAGS))})
        elif re.search(r"\b(?:DATE|TIME|CREATEDATE|SAVEDATE|PRINTDATE)\b", _instruction(p), re.I):
            result.append({"paragraph": p, "key": None, "label": "动态日期时间域", "value": text, "field": True})
    return result


def _value_pattern(label, value):
    digits = "[0-9０-９]"
    if label.endswith("日期"):
        allowed = (rf"{digits}{{4}}[.．/／\-－]{digits}{{1,2}}(?:[.．/／\-－]{digits}{{1,2}})?",
                   rf"{digits}{{4}}[ \u00a0\u3000]*年[ \u00a0\u3000]*{digits}{{1,2}}[ \u00a0\u3000]*月(?:[ \u00a0\u3000]*{digits}{{1,2}}[ \u00a0\u3000]*日)?")
    elif label.endswith("时间"):
        allowed = (rf"{digits}{{1,2}}[:：]{digits}{{2}}(?:[:：]{digits}{{2}})?",
                   rf"{digits}{{1,2}}[ \u00a0\u3000]*时[ \u00a0\u3000]*{digits}{{1,2}}[ \u00a0\u3000]*分(?:[ \u00a0\u3000]*{digits}{{1,2}}[ \u00a0\u3000]*秒)?")
    elif label.endswith(("年份", "年度")):
        allowed = (rf"{digits}{{4}}(?:[ \u00a0\u3000]*年)?",)
    elif label.endswith("编号"):
        allowed = (r"[A-Za-zＡ-Ｚａ-ｚ0-9０-９]+(?:[.．/／\-－][A-Za-zＡ-Ｚａ-ｚ0-9０-９]+)*",)
    else:
        return None
    if not any(re.fullmatch(pattern, value) for pattern in allowed):
        return None
    # Keep literal separators/spaces. Values are masked only in a separate
    # comparison signature; the source string is never transformed or written.
    tokens = re.split(r"([0-9０-９]+|[A-Za-zＡ-Ｚａ-ｚ]+)", value)
    signature = []
    for token in tokens:
        if not token:
            continue
        if re.fullmatch(r"[0-9０-９]+", token):
            widths = tuple(sorted({"全角" if ord(c) > 127 else "半角" for c in token}))
            signature.append(("数字", widths))
        elif label.endswith("编号") and re.fullmatch(r"[A-Za-zＡ-Ｚａ-ｚ]+", token):
            signature.append(("字母", tuple(sorted({("全角" if ord(c) > 127 else "半角",
                                                     "大写" if c.isupper() else "小写") for c in token}))))
        else:
            signature.append(("原样", token))
    lengths = tuple(len(token) for token in tokens if re.fullmatch(r"[0-9０-９]+", token or ""))
    numeric = tuple(token for token in tokens if re.fullmatch(r"[0-9０-９]+", token or ""))
    return tuple(signature), lengths, numeric


def _empty(p):
    # Table end-paragraphs, fields, bookmarks, drawings, explicit page/column
    # breaks and section properties are structural, not empty layout filler.
    if any(a.tag == qn("w:tc") for a in p._p.iterancestors()):
        return False
    if _original_text(p).strip(_BLANKS + "\n"):
        return False
    return not bool(p._p.xpath(".//w:drawing|.//w:pict|.//w:object|.//m:oMath|.//w:fldChar|.//w:fldSimple|.//w:instrText|.//w:bookmarkStart|.//w:sectPr|.//w:br[@w:type='page' or @w:type='column']|.//w:lastRenderedPageBreak|.//w:footnoteReference|.//w:endnoteReference")) and not effective_paragraph(p)["page_break_before"]


def _gap_records(ctx, mapping, side):
    records = _body(ctx, mapping, side)
    roles = {r["paragraph"]._p: r["role"] for r in records}
    # Actual native body paragraphs only: never cross a table or a content
    # control, whose empty end paragraphs have independent layout semantics.
    paragraphs = {p._p: p for p in ctx.paragraphs}
    previous, blanks, result = None, [], []
    for element in ctx.document.element.body:
        p = paragraphs.get(element)
        if p is None:
            previous, blanks = None, []
            continue
        if _empty(p):
            blanks.append(p)
            continue
        current = roles.get(element)
        if current and previous:
            prior, prior_role = previous
            result.append({"key": (prior_role, current), "count": len(blanks), "paragraph": p,
                           "blanks": list(blanks), "prior": prior,
                           "break_before": bool(effective_paragraph(p)["page_break_before"])})
        previous = (p, current) if current else None
        blanks = []
    return result


def _line_groups(text):
    return [len(match.group()) - 1 for match in re.finditer(r"\n{2,}", text)]


# ---------------------------------------------------------------------------
# 126/127/130/131/132：文档自身的字符卫生（不需要模板基准，规则本身就是标准）
# ---------------------------------------------------------------------------

#: 需要与数字之间留空格的计量单位（SI 写法）。
_SPACED_UNITS = ("kg", "mg", "g", "km", "cm", "mm", "nm", "m", "ms", "s", "min", "h",
                 "Hz", "kHz", "MHz", "GHz", "mol", "mmol", "mA", "A", "kV", "mV", "V",
                 "kW", "MW", "W", "kJ", "J", "kPa", "MPa", "Pa", "kN", "N", "K", "mL", "L",
                 "rpm", "dB")
_SPACED_UNIT_PATTERN = "|".join(sorted((re.escape(unit) for unit in _SPACED_UNITS), key=len, reverse=True))
#: 常见大小写错误（只在紧跟数字时判定，避免把 KG=knowledge graph 之类的缩写误报）。
_UNIT_CASE_FIXES = {"KG": "kg", "KM": "km", "CM": "cm", "MM": "mm", "ML": "mL",
                    "KW": "kW", "HZ": "Hz", "KHZ": "kHz", "MHZ": "MHz", "GHZ": "GHz",
                    "MPA": "MPa", "KPA": "kPa"}
_ZERO_WIDTH = re.compile(r"[\u200b-\u200f\u202a-\u202e\ufeff\u2060]")
_CONTROL_CHARS = re.compile(r"[\u0000-\u0008\u000e-\u001f\u007f]")
_BROKEN_CHARS = re.compile(r"[\ufffd\ufffe\uffff]")
_LINE_SEPARATORS = re.compile(r"[\u2028\u2029\u000b\u000c]")
_CJK_PUNCT = "，。；：、！？）】》」”’"
_CJK_PUNCT_OPEN = "（【《「“‘"


def _first_position(text, pattern):
    match = pattern.search(text)
    return None if match is None else match.start()


def _hygiene_findings(record):
    """逐段扫描字符卫生问题：一段一条，各自对应一个细粒度规则。"""

    paragraph = record["paragraph"]
    text = _original_text(paragraph)
    findings = []

    def report(code, expected, actual, position=0):
        finding = make_finding(code, _title(code), expected, actual, paragraph)
        finding.start = position
        finding.end = position + 1
        finding.excerpt = text[position:position + 12]
        findings.append(finding)

    position = _first_position(text, _BROKEN_CHARS)
    if position is not None:
        report("132.1", "不应出现乱码字符",
               f"第 {position + 1} 个字符是替换符或非法码位（U+{ord(text[position]):04X}）", position)
    position = _first_position(text, _CONTROL_CHARS) or _first_position(text, _ZERO_WIDTH)
    if position is not None:
        report("132.2", "不应包含不可见控制字符或零宽字符",
               f"第 {position + 1} 个字符为不可见字符（U+{ord(text[position]):04X}）", position)
    position = _first_position(text, _LINE_SEPARATORS)
    if position is not None:
        report("132.3", "换行应使用段落或标准换行符",
               f"第 {position + 1} 个字符是异常换行符（U+{ord(text[position]):04X}）", position)

    stripped = text.rstrip(" \u3000\t")
    if stripped != text and text.strip():
        position = len(stripped)
        report("126.1", "段首段尾不应有多余空格", f"段尾有 {len(text) - len(stripped)} 个多余空格", position)
    body = text.lstrip(" \u3000\t")
    offset = len(text) - len(body)
    match = re.search(r"[ \u3000]{2,}", body)
    if match:
        position = offset + match.start()
        report("126.3", "词语之间只保留一个半角空格",
               f"第 {position + 1} 个字符起有连续 {len(match.group())} 个空格", position)
    match = re.search(r"[ \u3000]+[" + re.escape(_CJK_PUNCT) + r"]", body)
    if match:
        position = offset + match.start()
        report("126.2", "中文标点前后不应有多余空格",
               f"第 {position + 1} 个字符起标点前有空格", position)
    match = re.search(r"[" + re.escape(_CJK_PUNCT) + r"][ \u3000]{2,}", body)
    if match:
        position = offset + match.start()
        report("126.2", "中文标点前后不应有多余空格",
               f"第 {position + 1} 个字符起标点后有连续空格", position)

    match = re.search(r"[0-9０-９](" + _SPACED_UNIT_PATTERN + r")(?![A-Za-z0-9])", body)
    if match:
        position = offset + match.start()
        report("127.2", f"数字与单位 {match.group(1)} 之间应有空格",
               f"“{match.group()}”数字与单位之间没有空格", position)
    match = re.search(r"[0-9０-９][ \u3000]+(?:%|℃)", body)
    if match:
        position = offset + match.start()
        report("127.2", "百分号/摄氏度符号前不加空格", f"“{match.group()}”符号前多了空格", position)

    match = re.search(r"[0-9０-９][ \u3000]*(" + "|".join(_UNIT_CASE_FIXES) + r")(?![A-Za-z])", body)
    if match:
        wrong = match.group(1)
        position = offset + match.end() - len(wrong)
        report("130.1", f"计量单位应写作 {_UNIT_CASE_FIXES[wrong]}",
               f"“{wrong}”大小写不符合计量单位写法", position)
    for run in paragraph.runs:
        if run.font.italic and re.fullmatch(_SPACED_UNIT_PATTERN, run.text.strip()):
            position = text.find(run.text)
            report("130.2", "计量单位应使用正体",
                   f"单位“{run.text.strip()}”被设置为斜体", max(position, 0))
            break

    match = re.search(r"[\uff05]|[0-9０-９][ \u3000]+(?:°C|℃|度)", body)
    if match:
        position = offset + match.start()
        report("131.1", "百分号写作 %，摄氏度写作 ℃（不留空格）",
               f"“{match.group()}”写法不规范", position)
    match = re.search(r"--+|[—–]{2,}", body)
    if match:
        position = offset + match.start()
        report("131.2", "连接号/破折号应使用标准符号“—”",
               f"“{match.group()}”不是规范连接号", position)
    match = re.search(r"\.{3,}|。{3,}|(?<!…)…(?!…)", body)
    if match:
        position = offset + match.start()
        report("131.3", "省略号应使用中文省略号“……”",
               f"“{match.group()}”写法不规范", position)
    return findings


def _check_hygiene(ctx, mapping=None):
    findings = []
    for record in _body(ctx, mapping, "document"):
        try:
            findings.extend(_hygiene_findings(record))
        except Exception:
            continue  # 单段读取失败不影响其他段落
    return findings


def extract_character_baseline(ctx, mapping=None):
    data = {"width": defaultdict(list), "spaces": defaultdict(list), "dates": defaultdict(list),
            "gaps": defaultdict(list), "lines": defaultdict(list), "issues": defaultdict(list)}
    for record in _body(ctx, mapping, "template"):
        p, role = record["paragraph"], record["role"]
        try:
            if _has_revision(p):
                data["issues"][role].append("模板正文含修订，原始字符范围无法确认")
                continue
            text = _original_text(p)
            for key, values in _width_samples(text).items():
                data["width"][(role, key)].append(tuple(v[0] for v in values))
            for key, values in _boundary_samples(text).items():
                data["spaces"][(role, key)].append(tuple(v[0] for v in values))
            data["lines"][role].append(tuple(_line_groups(text)))
        except Exception as exc:
            data["issues"][role].append(type(exc).__name__)
    for record in _field_records(ctx):
        key = record["key"]
        if key is not None:
            value = None if record["field"] or _has_revision(record["paragraph"]) else _value_pattern(record["label"], record["value"])
            data["dates"][key].append((value, record["value"]))
    for record in _gap_records(ctx, mapping, "template"):
        revised = any(_has_revision(p) for p in [record["prior"], record["paragraph"], *record["blanks"]])
        data["gaps"][record["key"]].append(None if revised else (record["count"], record["break_before"]))
    return {"characters": {key: dict(value) for key, value in data.items()}}


def _consensus(groups):
    # Two independent examples, unanimity, no document majority voting.
    if len(groups) < 2:
        return None
    values = {value for group in groups for value in group}
    return next(iter(values)) if len(values) == 1 else None


def _check_text(data, record):
    p, role = record["paragraph"], record["role"]
    if _has_revision(p):
        return [_unknown(rule, "正文含修订，字符有效范围不确定", p, True) for rule in ("124", "125")]
    text, findings = _original_text(p), []
    for rule, name, extractor in (("124", "width", _width_samples), ("125", "spaces", _boundary_samples)):
        bad, uncertain = [], []
        for key, values in extractor(text).items():
            expected = _consensus(data.get(name, {}).get((role, key), []))
            if expected is None or data.get("issues", {}).get(role):
                uncertain.append(key)
                continue
            for actual, _, evidence in values:
                if actual != expected:
                    bad.append((key, expected, actual, evidence))
        if bad:
            for key, exp, cur, evidence in bad:
                if rule == "124":
                    code = "124.1" if key in {"中文语境", "英文语境"} else "124.2"
                    expected_text = key + "=" + str(exp)
                    actual_text = "“" + evidence + "”" + key + "=" + str(cur)
                else:
                    # 空格比模板多 → 多余空格（125.2）；否则为缺少空格（125.1）
                    code = "125.2" if len(cur) > len(exp) else "125.1"
                    label = _boundary_label(key)
                    expected_text = "模板要求：" + label + _space_description(exp)
                    actual_text = ("“" + evidence + "”处" + label + _space_description(cur)
                                   + "；" + _auto_spacing_text(_auto_spacing(p))
                                   + "（自动间距与实际空格分开判断）")
                findings.append(make_finding(code, _title(code), expected_text, actual_text, p))
        if uncertain:
            findings.append(_unknown(rule, "模板对" + "、".join(uncertain) + "缺少至少两处同用途一致样例或样例冲突；不套用其他语境", p))
    return findings


def _check_dates(data, ctx):
    findings = []
    for record in _field_records(ctx):
        p, label = record["paragraph"], record["label"]
        try:
            if record["field"] or _has_revision(p):
                findings.append(_unknown("129", label + "含动态域或修订，缓存文字不作为已更新的日期时间结果", p, True))
                continue
            expected = data.get("dates", {}).get(record["key"], [])
            actual = _value_pattern(label, record["value"])
            if len(expected) < 2 or any(value is None for value, _ in expected):
                findings.append(_unknown("129", label + "缺少至少两处可确认的同用途模板示例或模板字段不可解析", p))
                continue
            signatures = {value[0] for value, _ in expected}
            if len(signatures) != 1 or actual is None:
                findings.append(_unknown("129", label + "模板形式冲突或当前字段不是可确认的日期/时间/年份/编号字符形式", p))
                continue
            current_signature, current_lengths, current_numbers = actual
            wrong = current_signature != next(iter(signatures))
            # Fixed numeric padding requires at least two distinct small
            # values showing that padding; one sample cannot imply it.
            if not wrong:
                for i, count in enumerate(current_lengths):
                    samples = [value for value, _ in expected]
                    lengths = {value[1][i] for value in samples}
                    starts_zero = {value[2][i] for value in samples if value[2][i].startswith(("0", "０"))}
                    if len(lengths) == 1 and len(starts_zero) >= 2 and count != next(iter(lengths)):
                        wrong = True
            if wrong:
                # 日期/时间/年份类 → 129.1；编号类 → 129.2
                code = "129.2" if "编号" in label or "编号" in str(record.get("key", "")) else "129.1"
                findings.append(make_finding(code, _title(code), label + "同用途字符形式参照：" + "；".join(value for _, value in expected[:3]),
                                             label + "：“" + record["value"] + "”（仅比对字符形式，不校验内容真实性）", p))
        except Exception as exc:
            findings.append(_unknown("129", type(exc).__name__, p, True))
    return findings


def _check_gaps(data, ctx, mapping):
    findings = []
    for record in _gap_records(ctx, mapping, "document"):
        if record["count"] < 2:
            continue  # Existing 075 retains responsibility for single blanks.
        p = record["paragraph"]
        expected = data.get("gaps", {}).get(record["key"], [])
        if _has_revision(record["prior"]) or _has_revision(p) or any(_has_revision(b) for b in record["blanks"]):
            findings.append(_unknown("128", "空段及相邻段落含修订，当前有效段落边界不可确认", p, True))
            continue
        if len(expected) < 2 or None in expected or len(set(expected)) != 1 or any(data.get("issues", {}).get(role) for role in record["key"]):
            findings.append(_unknown("128", "该正文边界的模板空段/分页示例不足或冲突；不按固定空行阈值判断", p))
            continue
        count, page_before = expected[0]
        if record["count"] > count:
            finding = make_finding("128", TITLES["128"], f"对应正文边界空段数={count}；后段段前分页={page_before}",
                                   f"连续空段数={record['count']}；后段段前分页={record['break_before']}；物理分页效果未经排版确认", p)
            findings.append(finding)
    for record in _body(ctx, mapping, "document"):
        p, role = record["paragraph"], record["role"]
        groups = _line_groups(_original_text(p))
        if not groups:
            continue
        values = data.get("lines", {}).get(role, [])
        if _has_revision(p) or data.get("issues", {}).get(role) or len(values) < 2 or len(set(values)) != 1:
            findings.append(_unknown("128", "同用途段内换行样例不足、冲突或含修订，连续空行不能确认为错误", p))
        elif tuple(groups) != values[0] and (not values[0] or max(groups) > max(values[0])):
            findings.append(make_finding("128", TITLES["128"], "同用途模板连续手动换行形成的空行=" + str(values[0]),
                                         "当前连续手动换行形成的空行=" + str(tuple(groups)) + "；不将其当成真实分页结果", p))
    return findings


def check_character_errors(baseline, ctx, mapping=None):
    data = baseline.get("characters", {})
    if data.get("unreadable"):
        return [_unknown(rule, data["unreadable"], unreadable=True) for rule in TITLES]
    findings = []
    for record in _body(ctx, mapping, "document"):
        try:
            findings.extend(_check_text(data, record))
        except Exception as exc:
            findings.extend(_unknown(rule, type(exc).__name__, record["paragraph"], True) for rule in ("124", "125"))
    for rule, checker in (("129", lambda: _check_dates(data, ctx)), ("128", lambda: _check_gaps(data, ctx, mapping))):
        try:
            findings.extend(checker())
        except Exception as exc:
            findings.append(_unknown(rule, type(exc).__name__, unreadable=True))
    try:
        findings.extend(_check_hygiene(ctx, mapping))
    except Exception as exc:
        findings.append(_unknown("126", type(exc).__name__, unreadable=True))
    return findings
