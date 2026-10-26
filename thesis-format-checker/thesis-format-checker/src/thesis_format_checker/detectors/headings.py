"""Read-only heading/list checks for rules 048--055.

Rules 045--047 remain in the original COM detectors.  XML list identifiers
are document-local: comparisons below use resolved definitions, never numId
or styleId equality between documents.  No rendered-page facts are invented.
"""
from __future__ import annotations

import re
from collections import defaultdict

from thesis_format_checker.core.xmlutil import qn
from docx.text.paragraph import Paragraph

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, effective_paragraph, effective_runs, make_finding, unknown, memo
from thesis_format_checker.core.scope import front_matter_end


_TITLES = {
    "048": "标题层级跳级", "049": "同级标题格式或编号不一致",
    "050": "章节编号不连续", "051": "标题编号格式不符",
    "052": "标题编号与文字分隔不符", "053": "标题编号缺失或多余",
    "054": "标题与后续内容同页状态无法确认", "055": "标题使用手工编号",
}
_SPECIAL = {
    "摘要": "abstract-cn", "中文摘要": "abstract-cn", "abstract": "abstract-en",
    "英文摘要": "abstract-en", "目录": "contents", "contents": "contents",
    "致谢": "acknowledgements", "acknowledgements": "acknowledgements",
    "acknowledgments": "acknowledgements", "附录": "appendix", "appendix": "appendix",
    "原创性声明": "declaration", "独创性声明": "declaration", "授权使用声明": "authorization",
    "攻读学位期间成果": "achievements", "参考文献": "references", "references": "references",
}


def _normalize(text):
    return re.sub(r"\s+", "", text).casefold()


def _styles(paragraph):
    seen = set()
    style = paragraph.style
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


def _val(parent, name, default=None):
    node = parent.find(qn("w:" + name)) if parent is not None else None
    return node.get(qn("w:val"), default) if node is not None else default


# ---------------------------------------------------------------------------
# 标题层级树：按文档自身的编号体系识别标题，避免把误用标题样式的正文当标题
# ---------------------------------------------------------------------------

#: 标题文本长度上限：超过即视为正文（正文段落常被误设为“标题 2”样式）。
_MAX_HEADING_CHARS = 40
#: 句末标点：出现即不是标题。
_SENTENCE_END = re.compile(r"[。；！？!?;]")
#: 以冒号/逗号结尾的引导句（“本文采用以下方法：”）是被误设样式的正文，不是标题。
_TRAILING_LIST_PUNCT = re.compile(r"[：:，,、]$")
_CHAPTER_PREFIX = re.compile(r"^第\s*([0-9〇零一二三四五六七八九十百千两]{1,6})\s*[章篇]")
_SECTION_PREFIX = re.compile(r"^第\s*([0-9〇零一二三四五六七八九十百千两]{1,6})\s*节")
#: 多级编号前缀：1 / 1.1 / 1.1.1（是否算标题交给 _numbering_level 继续判断）。
_NUMBER_PREFIX = re.compile(r"^(\d{1,3}(?:[.．]\d{1,3})*)")
#: 编号后的标题文字至少要有一个汉字或字母，用来排除“31.5%”这类数据。
_TITLE_CHARS = re.compile(r"[A-Za-z\u3400-\u9fff]")
_TOC_STYLE = re.compile(r"^(?:toc|目录)\s*([1-9])$", re.I)
_STYLE_NAME_LEVEL = re.compile(r"(?:heading|标题)\s*([1-9])", re.I)
#: 文档标题（封面题目）样式：属于封面信息，不是章节标题。
_DOCUMENT_TITLE_STYLES = {"title", "题目", "论文题目", "副标题", "subtitle"}
#: 图表题注（图2-1/表3-2/Fig.1/Tab.1…）：即使被误用标题样式，也不是章节标题。
_CAPTION_LIKE = re.compile(r"^(?:图|表|Figure|Table|Fig\.?|Tab\.?)\s*[0-9０-９]", re.I)


def _style_heading_level(paragraph):
    """仅按样式/大纲级别判断（旧逻辑；作为标题树无法建立时的兜底）。"""

    styles = list(_styles(paragraph))
    # A TOC paragraph can inherit a heading-like custom style.  Its own TOC
    # role wins: cached directory entries are not a second body heading tree.
    if styles and re.match(r"^(toc|目录)\s*\d*$", styles[0].name or "", re.I):
        return None
    for props in _properties(paragraph):
        value = _val(props, "outlineLvl")
        if value is not None:
            value = int(value)
            return value + 1 if 0 <= value <= 8 else None
    for style in styles:
        match = re.fullmatch(r"(?:heading|标题)\s*([1-9])", style.name or "", re.I)
        if match:
            return int(match.group(1))
    return None


def _first_style_name(paragraph) -> str:
    styles = list(_styles(paragraph))
    return (styles[0].name or "") if styles else ""


def _toc_entries(paragraphs) -> dict[str, int]:
    """目录域缓存（toc 1/2/3）→ {规范化标题: 级别}：文档自己声明的层级。"""

    entries: dict[str, int] = {}
    for paragraph in paragraphs:
        match = _TOC_STYLE.match(_first_style_name(paragraph))
        if not match:
            continue
        # 去掉页码：阿拉伯数字或罗马数字（前置部分常用 I/II 或 Ⅰ/Ⅱ）
        text = re.sub(r"[\s\u3000]*[0-9０-９Ⅰ-ⅿIVXivx]+\s*$", "", paragraph.text.strip())
        key = _normalize(text)
        if key:
            entries.setdefault(key, int(match.group(1)))
    return entries


def _looks_like_heading(text: str) -> bool:
    """短、无句末标点才算“像标题”。"""

    stripped = text.strip()
    if not stripped or len(stripped) > _MAX_HEADING_CHARS:
        return False
    if _TRAILING_LIST_PUNCT.search(stripped):
        return False
    return not _SENTENCE_END.search(stripped)


def _numbering_level(text: str, chapter_prefix: bool):
    """按正文编号推断层级：第X章 / N / N.N / N.N.N；无编号返回 None。

    返回值 ``(级别, 是否强编号)``：``第X章`` 与多级编号（N.N）本身就是结构证据；
    单独的 ``N`` 可能只是正文里的列表序号，需要样式佐证才当标题。
    """

    if _CHAPTER_PREFIX.match(text):
        return 1, True
    if _SECTION_PREFIX.match(text):
        return (2 if chapter_prefix else 1), True
    match = _NUMBER_PREFIX.match(text)
    if not match:
        return None, False
    rest = re.sub(r"^[ \t\u3000]*[.．、:：]?", "", text[match.end():]).strip()
    parts = re.split(r"[.．]", match.group(1))
    if not rest or int(parts[0]) > 99 or int(parts[0]) == 0:
        return None, False  # 纯数字开头更像正文数据（年份、数量）
    if rest[0].isdigit() or rest[0] in ".．":
        return None, False  # 数字还没结束，例如 0.5320
    if not _TITLE_CHARS.search(rest):
        return None, False  # 只有百分号等符号，不是标题
    return len(parts) + (1 if chapter_prefix else 0), len(parts) > 1


def _build_heading_levels(document) -> dict:
    """重建标题层级：目录缓存 > 正文编号 > 样式，且必须通过“像标题”检查。

    摘要页之前（封面、题名页、声明页等前置区）不参与标题层级判断：那里的
    “原创性声明”一类标题、封面题目只由封面模块按模板同位置比对，否则会被
    当作章节标题参与 048–055 等通用检测，产生整片误报。
    """

    paragraphs = [Paragraph(element, document._body)
                  for element in document.element.body.iter(qn("w:p"))
                  if not any(ancestor.tag == qn("w:txbxContent") for ancestor in element.iterancestors())]
    front_end = front_matter_end([paragraph.text for paragraph in paragraphs])
    entries = _toc_entries(paragraphs)
    chapter_prefix = any(_CHAPTER_PREFIX.match(paragraph.text.strip())
                         for index, paragraph in enumerate(paragraphs) if index >= front_end)
    levels: dict = {}
    poisoned: set = set()
    seen: set[int] = set()
    for index, paragraph in enumerate(paragraphs):
        if 0 < front_end and index < front_end:
            continue  # 封面/声明等前置区：只由封面模块按模板同位置比对
        text = paragraph.text.strip()
        if not text or _TOC_STYLE.match(_first_style_name(paragraph)):
            continue
        if _first_style_name(paragraph).strip().casefold() in _DOCUMENT_TITLE_STYLES:
            continue  # 封面题目不是章节标题
        if _CAPTION_LIKE.match(text):
            continue  # 图表题注不是章节标题
        if any(ancestor.tag == qn("w:tc") for ancestor in paragraph._p.iterancestors()):
            continue  # 表格单元格里的编号/数值不是文档标题
        try:
            style_level = _style_heading_level(paragraph)
        except Exception:
            # 大纲级别不可读：保留原行为，由调用方报告“无法判定”。
            poisoned.add(paragraph._p)
            continue
        toc_level = entries.get(_normalize(text))
        if not _looks_like_heading(text) and toc_level is None:
            continue  # 长句正文：即使样式被设成标题，也不算标题
        numbered, strong_number = _numbering_level(text, chapter_prefix)
        if numbered is not None and numbered > 1 and (numbered - 1) not in seen:
            numbered = None  # 父级标题没出现过，说明这不是真实标题编号
        if numbered is not None and not strong_number and style_level is None:
            numbered = None  # 单独的“N 标题”需要样式佐证，避免把列表序号当标题
        # 目录缓存是文档自己声明的层级，优先于编号与样式。
        level = toc_level or numbered or style_level
        if not level:
            continue
        levels[paragraph._p] = level
        seen.add(level)
    return levels, poisoned


def _heading_levels_cached(paragraph):
    """按文档缓存标题层级表；无法建立时返回 None（退回样式判断）。

    缓存挂在 ``part`` 上，按“一次解析、一次检测”的口径复用；同一
    Document 修改段落之后再检测需要重新打开文档，否则会读到旧的标题树。
    """

    try:
        # 注意：``part.document`` 每次都会新建一个包装对象，缓存必须挂在 part 上。
        part = paragraph.part
        document = part.document
    except Exception:
        return None
    cached = getattr(part, "_thesis_heading_levels", None)
    if cached is not None:
        return cached
    try:
        cached = _build_heading_levels(document)
    except Exception:
        return None
    try:
        part._thesis_heading_levels = cached
    except Exception:
        pass
    return cached


def _heading_level(paragraph):
    """段落标题级别：优先用标题层级树，建树失败时退回样式/大纲级别。"""

    cached = _heading_levels_cached(paragraph)
    if cached is None:
        return _style_heading_level(paragraph)
    levels, poisoned = cached
    if paragraph._p in poisoned:
        return _style_heading_level(paragraph)  # 让调用方拿到“大纲级别不可读”
    return levels.get(paragraph._p)


def _roman(value):
    if not 0 < value < 4000:
        return None
    result = []
    for number, token in ((1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                          (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                          (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")):
        while value >= number:
            result.append(token)
            value -= number
    return "".join(result)


def _chinese_number(token):
    digits = dict(zip("零一二三四五六七八九", range(10)))
    digits.update({"〇": 0, "两": 2})
    if token in digits:
        return digits[token]
    if token and all(ch in digits for ch in token):
        return int("".join(str(digits[ch]) for ch in token))
    total, unit_value = 0, 0
    for ch in token:
        if ch in digits:
            unit_value = digits[ch]
        elif ch in {"十", "百", "千"}:
            total += (unit_value or 1) * {"十": 10, "百": 100, "千": 1000}[ch]
            unit_value = 0
        else:
            return None
    return total + unit_value if token else None


def _token_value(token):
    if token.isdecimal():
        return int(token), "decimal"
    if re.fullmatch(r"[零〇一二两三四五六七八九十百千]+", token):
        return _chinese_number(token), "chineseCounting"
    if re.fullmatch(r"[IVXLCDMivxlcdm]+", token):
        upper = token.upper()
        values = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100, "D": 500, "M": 1000}
        value = sum(-values[ch] if i + 1 < len(upper) and values[ch] < values[upper[i + 1]]
                    else values[ch] for i, ch in enumerate(upper))
        if _roman(value) == upper:
            return value, "upperRoman" if token.isupper() else "lowerRoman"
    if re.fullmatch(r"[A-Za-z]", token):
        return ord(token.upper()) - ord("A") + 1, "upperLetter" if token.isupper() else "lowerLetter"
    return None, None


def _manual_number(text):
    """Recognize lexical prefixes only; a template must authorize any check.

    A bare year attached to title words, e.g. '2026年研究', is intentionally
    not a number.  The parser is not an inference that every prefix is required.
    """
    token = r"[0-9０-９零〇一二两三四五六七八九十百千]+"
    match = re.match(rf"^(第)({token})(章|节|篇|部分)([ \t\u3000]*)(.+)$", text)
    if match:
        value, fmt = _token_value(match[2])
        return {"number": match[1] + match[2] + match[3], "number_values": [value],
                "number_pattern": "第%1" + match[3], "number_formats": (fmt,),
                "separator": match[4], "title": match[5]}
    match = re.match(r"^([（(])([0-9０-９零〇一二两三四五六七八九十百千IVXLCDMivxlcdmA-Za-z]+)([）)])([ \t\u3000]*)(.+)$", text)
    if match and (match[1], match[3]) in (("(", ")"), ("（", "）")):
        value, fmt = _token_value(match[2])
        if value is not None:
            return {"number": match[1] + match[2] + match[3], "number_values": [value],
                    "number_pattern": match[1] + "%1" + match[3], "number_formats": (fmt,),
                    "separator": match[4], "title": match[5]}
    match = re.match(r"^([0-9０-９]+(?:[.．\-][0-9０-９]+)*)([.．、)]?)([ \t\u3000]+)(.+)$", text)
    if not match:
        match = re.match(r"^([0-9０-９]+)([.．、)])([ \t\u3000]*)([^0-9０-９].*)$", text)
    if not match:
        # A dotted multi-component number has an unambiguous internal boundary.
        match = re.match(r"^([0-9０-９]+(?:[.．\-][0-9０-９]+)+)([.．、)]?)([ \t\u3000]*)([^0-9０-９.．\-].*)$", text)
    if match:
        numbers = re.findall(r"[0-9０-９]+", match[1])
        iterator = iter(range(1, len(numbers) + 1))
        pattern = re.sub(r"[0-9０-９]+", lambda _: "%" + str(next(iterator)), match[1]) + match[2]
        return {"number": match[1] + match[2], "number_values": list(map(int, numbers)),
                "number_pattern": pattern, "number_formats": ("decimal",) * len(numbers),
                "separator": match[3], "title": match[4]}
    match = re.match(r"^([零〇一二两三四五六七八九十百千]+)([、.．])([ \t\u3000]*)(.+)$", text)
    if not match:
        match = re.match(r"^([IVXLCDMivxlcdmA-Za-z]+)([.)]?)([ \t\u3000]+)(.+)$", text)
    if match:
        value, fmt = _token_value(match[1])
        if value is not None:
            return {"number": match[1] + match[2], "number_values": [value],
                    "number_pattern": "%1" + match[2], "number_formats": (fmt,),
                    "separator": match[3], "title": match[4]}
    return None


def _num_pr(paragraph):
    values = {}
    for props in _properties(paragraph):
        num = props.find(qn("w:numPr"))
        if num is None:
            continue
        for key in ("numId", "ilvl"):
            if key not in values and _val(num, key) is not None:
                values[key] = int(_val(num, key))
        if values.get("numId") == 0:
            return None
    if "numId" not in values:
        return None
    return values


def _list_definition(document, paragraph, num_pr):
    root = document.part.numbering_part.element
    nums = [node for node in root.findall(qn("w:num"))
            if node.get(qn("w:numId")) == str(num_pr["numId"])]
    if len(nums) != 1:
        raise ValueError("列表实例不存在或重复")
    num = nums[0]
    abstract_id = _val(num, "abstractNumId")
    abstracts = [node for node in root.findall(qn("w:abstractNum"))
                 if node.get(qn("w:abstractNumId")) == abstract_id]
    if len(abstracts) != 1:
        raise ValueError("抽象多级列表不存在或重复")
    abstract = abstracts[0]
    if abstract.find(qn("w:numStyleLink")) is not None:
        raise ValueError("编号样式链接尚不能可靠展开")
    if abstract.get("{http://schemas.microsoft.com/office/word/2012/wordml}restartNumberingAfterBreak") in ("1", "true"):
        raise ValueError("列表按节重新编号的扩展依赖Word排版")
    levels = {}
    for level in abstract.findall(qn("w:lvl")):
        key = int(level.get(qn("w:ilvl")))
        if key in levels:
            raise ValueError("列表层级定义重复")
        levels[key] = level
    starts = {}
    for override in num.findall(qn("w:lvlOverride")):
        key = int(override.get(qn("w:ilvl")))
        replacement = override.find(qn("w:lvl"))
        if replacement is not None:
            levels[key] = replacement
        if _val(override, "startOverride") is not None:
            starts[key] = int(_val(override, "startOverride"))
    ilvl = num_pr.get("ilvl")
    if ilvl is None:
        ids = {style.style_id for style in _styles(paragraph)}
        matches = [key for key, level in levels.items() if _val(level, "pStyle") in ids]
        if len(matches) > 1:
            raise ValueError("段落样式映射到多个编号层级")
        ilvl = matches[0] if matches else 0  # OOXML absence semantics.
    if ilvl not in levels:
        raise ValueError("所用编号层级不存在")
    definitions = {}
    for key, level in levels.items():
        start = starts.get(key)
        if start is None and _val(level, "start") is not None:
            start = int(_val(level, "start"))
        restart_raw = _val(level, "lvlRestart")
        restart = int(restart_raw) - 1 if restart_raw is not None else key - 1
        if restart_raw == "0":
            restart = None
        if restart is not None and restart >= key:
            raise ValueError("重新编号层级关系无效")
        tabs = []
        for tab in level.findall("./" + qn("w:pPr") + "/" + qn("w:tabs") + "/" + qn("w:tab")):
            tabs.append((tab.get(qn("w:val")), tab.get(qn("w:pos")), tab.get(qn("w:leader"), "none")))
        indentation = level.find("./" + qn("w:pPr") + "/" + qn("w:ind"))
        definitions[key] = {"pattern": _val(level, "lvlText"), "format": _val(level, "numFmt"),
                            "start": start, "restart": restart, "suffix": _val(level, "suff", "tab"),
                            "legal": level.find(qn("w:isLgl")) is not None,
                            "tabs": tuple(sorted(tabs)), "ilvl": key,
                            "indent": tuple(sorted(indentation.attrib.items())) if indentation is not None else (),
                            "picture": level.find(qn("w:lvlPicBulletId")) is not None}
        if definitions[key]["suffix"] not in ("tab", "space", "nothing"):
            raise ValueError("列表编号后跟随字符设置无效")
    return {"num_id": num_pr["numId"], "ilvl": ilvl, "levels": definitions,
            "multilevel": _val(abstract, "multiLevelType") != "singleLevel"}


def _format_number(value, fmt):
    if fmt == "decimal":
        return str(value)
    if fmt in ("upperRoman", "lowerRoman"):
        result = _roman(value)
        return result.lower() if result and fmt == "lowerRoman" else result
    if fmt in ("upperLetter", "lowerLetter") and 1 <= value <= 26:
        result = chr(ord("A") + value - 1)
        return result.lower() if fmt == "lowerLetter" else result
    # East-Asian formats, legal numbering and >26 letter behavior are not
    # approximated; their raw definitions remain available for comparison.
    return None


def _list_label(definition, states):
    num_id, ilvl, levels = definition["num_id"], definition["ilvl"], definition["levels"]
    current = levels[ilvl]
    counters = states.setdefault(num_id, {})
    if current["start"] is None:
        raise ValueError("列表未提供可确认的起始值")
    counters[ilvl] = counters.get(ilvl, current["start"] - 1) + 1
    # Descendant reset policies are OOXML list semantics, not school rules.
    for key in list(counters):
        if key > ilvl and levels.get(key, {}).get("restart") is not None:
            if ilvl <= levels[key]["restart"]:
                counters.pop(key, None)
    pattern = current["pattern"]
    if not pattern or current["picture"]:
        raise ValueError("图片或缺失文本的编号无法取得显示值")
    if current["legal"]:
        raise ValueError("法律式编号显示尚不能可靠计算")
    references = [int(value) - 1 for value in re.findall(r"%([1-9])", pattern)]
    values, formats, replacements = [], [], {}
    for key in references:
        if key not in levels or key not in counters:
            raise ValueError("编号引用的上级计数不可确认")
        fmt = levels[key]["format"]
        rendered = _format_number(counters[key], fmt)
        if rendered is None:
            raise ValueError("编号数字格式的实际显示需Word计算")
        values.append(counters[key])
        formats.append(fmt)
        replacements[key + 1] = rendered
    if not references:
        raise ValueError("项目符号或固定编号不是可计算的章节序列")
    label = re.sub(r"%([1-9])", lambda m: replacements[int(m[1])], pattern)
    return label, values, tuple(formats)


def __toc_paragraph_ids_impl(ctx):
    excluded = set()
    active = []
    for paragraph in ctx.paragraphs:
        inside = any(item["toc"] for item in active)
        for node in paragraph._p.iter():
            if node.tag == qn("w:fldSimple"):
                if re.match(r"\s*TOC\b", node.get(qn("w:instr"), ""), re.I):
                    inside = True
            elif node.tag == qn("w:fldChar"):
                kind = node.get(qn("w:fldCharType"))
                if kind == "begin":
                    active.append({"toc": False, "instruction": ""})
                elif kind == "end" and active:
                    inside = inside or active[-1]["toc"]
                    active.pop()
            elif node.tag == qn("w:instrText") and active:
                active[-1]["instruction"] += node.text or ""
                active[-1]["toc"] = bool(re.match(r"\s*TOC\b", active[-1]["instruction"], re.I))
                inside = inside or active[-1]["toc"]
        if inside:
            excluded.add(paragraph._p)
    return excluded


def _collect_heading_records_impl(ctx, mapping=None):
    """Collect body heading records without updating fields or mutating XML.

    `number_known` distinguishes a known unnumbered heading from an unknown
    automatic label.  Consumers must not compare a None display_text as text.
    `mapping` is reserved for the existing document mapping contract.
    """
    del mapping
    records, counters = [], {}
    excluded = _toc_paragraph_ids(ctx)
    poisoned_sequences = set()
    for paragraph in ctx.paragraphs:
        if paragraph._p in excluded:
            continue
        if any(ancestor.tag in (qn("w:del"), qn("w:moveFrom")) for ancestor in paragraph._p.iterancestors()):
            continue
        text = paragraph.text
        level_error = None
        try:
            level = _heading_level(paragraph)
        except Exception as exc:
            level = None
            level_error = "标题大纲级别读取异常：" + str(exc)
            for style in _styles(paragraph):
                match = re.fullmatch(r"(?:heading|标题)\s*([1-9])", style.name or "", re.I)
                if match:
                    level = int(match[1])
                    break
        definition, num_pr, label, values, formats, error = None, None, None, [], (), None
        definition_error = None
        try:
            num_pr = _num_pr(paragraph)
            if num_pr:
                definition = _list_definition(ctx.document, paragraph, num_pr)
                if num_pr["numId"] in poisoned_sequences:
                    raise ValueError("本列表此前存在无法可靠读取的编号节点")
                label, values, formats = _list_label(definition, counters)
        except Exception as exc:
            error = str(exc)
            if definition is None:
                definition_error = error
            if num_pr and definition is None:
                poisoned_sequences.add(num_pr["numId"])
        if level is None:
            continue
        record = {"paragraph": paragraph, "level": level, "text": text, "title": text,
                  "style_name": paragraph.style.name if paragraph.style is not None else "",
                  "number": label, "number_values": values, "number_formats": formats,
                  "automatic": bool(num_pr), "numbering_unknown": error,
                  "definition_error": definition_error,
                  "numbering": definition, "number_pattern": None, "separator": "",
                  "numbering_signature": None}
        if level_error:
            record["numbering_unknown"] = level_error
            record["definition_error"] = level_error
            record["level_unknown"] = True
        if (paragraph._p.xpath(".//w:pPrChange|.//w:numberingChange|.//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo") or
                any(ancestor.tag in (qn("w:ins"), qn("w:moveTo")) for ancestor in paragraph._p.iterancestors())):
            record["numbering_unknown"] = "标题含未接受的修订或编号历史，需先明确修订视图"
            record["revision_affected"] = True
        fields = paragraph._p.xpath(".//w:fldSimple|.//w:fldChar|.//w:instrText")
        if fields:
            record["numbering_unknown"] = "标题含动态域，编号与标题显示边界需Word计算"
            record["field_numbering"] = True
        if definition:
            spec = definition["levels"][definition["ilvl"]]
            record["number_pattern"] = spec["pattern"]
            record["separator"] = {"tab": "\t", "space": " ", "nothing": ""}.get(spec["suffix"])
            record["numbering_signature"] = (spec["pattern"], spec["format"], spec["suffix"])
        elif not num_pr and not record["numbering_unknown"]:
            manual = _manual_number(text)
            if manual:
                record.update(manual)
                record["numbering_signature"] = (record["number_pattern"], record["number_formats"], record["separator"])
        record["role"] = _SPECIAL.get(_normalize(record["title"]), "body")
        record["number_known"] = record["numbering_unknown"] is None
        record["num_label"] = record["number"]
        if not record["number_known"]:
            record["display_text"] = None
        elif record["automatic"]:
            record["display_text"] = (record["number"] or "") + (record["separator"] or "") + text
        else:
            record["display_text"] = text
        records.append(record)
    return records


def _format_signature(record):
    paragraph = record["paragraph"]
    runs = effective_runs(paragraph)
    result = {}
    for key, matches in (("east_asia_font", r"[\u3400-\u9fff]"), ("ascii_font", r"[A-Za-z0-9]"),
                         ("size_pt", r"\S")):
        values = {props[key] for text, props in runs if re.search(matches, text) and props[key] is not None}
        result[key] = tuple(sorted(values, key=str)) if values else None
    props = effective_paragraph(paragraph)
    for key in ("first_line_indent", "left_indent", "right_indent"):
        result[key] = props.get(key)
    if record.get("numbering"):
        spec = record["numbering"]["levels"][record["numbering"]["ilvl"]]
        if spec["indent"]:
            for key in ("first_line_indent", "left_indent", "right_indent"):
                result[key] = None
    result["numbering"] = _shape(record)
    return result


def _shape(record):
    pattern = record.get("number_pattern")
    if record.get("numbering"):
        spec = record["numbering"]["levels"][record["numbering"]["ilvl"]]
        references = [int(value) - 1 for value in re.findall(r"%([1-9])", pattern or "")]
        formats = tuple(record["numbering"]["levels"].get(key, {}).get("format") for key in references)
        # Rebase placeholder indices only for comparing visible shape.  A
        # standalone level-2 '%2' has the same shape as manual '1'.
        order = {value: i + 1 for i, value in enumerate(dict.fromkeys(references))}
        normalized = re.sub(r"%([1-9])", lambda m: "%" + str(order[int(m[1]) - 1]), pattern or "")
        return normalized, formats
    if record.get("number"):
        return pattern, tuple(record.get("number_formats", ()))
    return None


def _candidates(data, record):
    candidates = [item for item in data["headings"]
                  if item["level"] == record["level"] and item["role"] == record["role"]]
    if not candidates:
        return []
    exact = [item for item in candidates if _normalize(item["title"]) == _normalize(record["title"])]
    modes = {_mode(item) for item in candidates}
    if len(modes) > 1:
        same_style = [item for item in candidates if item["style_name"] == record["style_name"]]
        if exact:
            return exact
        if same_style and len({_mode(item) for item in same_style}) == 1:
            return same_style
    return candidates


def _mode(record):
    if record["automatic"]:
        return "automatic"
    if record.get("numbering_unknown"):
        return "unknown"
    return "manual" if record.get("number") else "none"


def extract_heading_numbering_baseline(ctx, mapping=None):
    headings = collect_heading_records(ctx, mapping)
    for record in headings:
        try:
            record["format_signature"] = _format_signature(record)
            record["pagination"] = effective_paragraph(record["paragraph"])
        except Exception as exc:
            record["format_signature"] = None
            record["pagination"] = {}
            record["format_error"] = str(exc)
    transitions, levels = defaultdict(set), defaultdict(set)
    previous = {}
    for record in headings:
        role, level = record["role"], record["level"]
        levels[role].add(level)
        if role in previous and level > previous[role]:
            transitions[role].add((previous[role], level))
        previous[role] = level
    return {"headings": {"headings": headings, "transitions": dict(transitions), "levels": dict(levels)}}


def _unavailable(rule, reason, record=None):
    return unknown(rule, _title(rule), UNREADABLE + "；" + reason,
                   record["paragraph"] if record else None)


#: 同级标题属性 → 细粒度规则编号
_SAME_LEVEL_CODES = {
    "east_asia_font": "049.1", "ascii_font": "049.1",
    "size_pt": "049.2",
    "first_line_indent": "049.3", "left_indent": "049.3", "right_indent": "049.3",
    "numbering": "049.4",
}


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else _TITLES[code]


def _separator_code(record: dict, extra_wrong: bool) -> str:
    """编号与文字的分隔问题 → 052.1 空格 / 052.2 制表符 / 052.3 标点。"""

    if extra_wrong:
        return "052.1"
    description = _separator_description(record["separator"])
    if "制表" in description:
        return "052.2"
    if "空格" in description:
        return "052.1"
    return "052.3"


def _expected_sequence(candidates):
    automatic = [record for record in candidates if record.get("numbering")]
    if automatic:
        specs = [record["numbering"]["levels"][record["numbering"]["ilvl"]] for record in automatic]
        keys = {(item["start"], item["restart"], item["ilvl"]) for item in specs}
        if len(keys) == 1 and specs[0]["start"] is not None:
            displayed = defaultdict(list)
            for record in automatic:
                if record["number_known"] and record["number_values"]:
                    displayed[(record["numbering"]["num_id"], tuple(record["number_values"][:-1]))].append(record["number_values"][-1])
            steps = [b - a for values in displayed.values() for a, b in zip(values, values[1:])]
            return {"start": specs[0]["start"], "step": 1, "restart": specs[0]["restart"],
                    "ilvl": specs[0]["ilvl"], "heading_continuous": bool(steps) and all(step == 1 for step in steps)}
        return None
    manual = [record for record in candidates if record.get("number_values")]
    if len(manual) < 2:
        return None
    groups = defaultdict(list)
    for record in manual:
        groups[tuple(record["number_values"][:-1])].append(record["number_values"][-1])
    starts = {items[0] for items in groups.values()}
    steps = {b - a for items in groups.values() for a, b in zip(items, items[1:])}
    if len(starts) == 1 and len(steps) == 1 and next(iter(steps)) > 0:
        level = manual[0]["level"]
        return {"start": next(iter(starts)), "step": next(iter(steps)), "restart": level - 2,
                "ilvl": level - 1}
    return None


def _separator_description(separator):
    if separator is None:
        return "无法读取"
    if not separator:
        return "不附加字符"
    return "、".join("制表符" if ch == "\t" else "全角空格" if ch == "\u3000" else "半角空格"
                    if ch == " " else repr(ch) for ch in separator)


def _shape_description(shape):
    if shape is None:
        return "无编号"
    pattern, formats = shape
    names = {"decimal": "阿拉伯数字", "upperRoman": "大写罗马数字", "lowerRoman": "小写罗马数字",
             "upperLetter": "大写字母", "lowerLetter": "小写字母", "chineseCounting": "中文数字",
             "chineseCountingThousand": "中文千位计数", "decimalZero": "前补零数字"}
    return re.sub(r"%([1-9])", lambda match: "〈序号" + match[1] + "〉", pattern or "无法读取") + \
        "（" + "、".join(names.get(value, str(value)) for value in formats) + "）"


def _restart_description(value):
    if value is None:
        return "不随上级重启"
    return "无上级" if value < 0 else f"遇第{value + 1}级及其上级标题重启"


def check_heading_numbering_errors(baseline, ctx, mapping=None):
    data = baseline.get("headings", {})
    if not data.get("headings"):
        if collect_heading_records(ctx, mapping):
            return [unknown(rule, title, "模板没有可可靠匹配的标题层级或编号示例")
                    for rule, title in _TITLES.items()]
        return []
    records = collect_heading_records(ctx, mapping)
    findings, unknown_keys = [], set()

    def uncertain(rule, reason, record=None):
        key = (rule, record["role"] if record else "", record["level"] if record else None, reason)
        if key not in unknown_keys:
            unknown_keys.add(key)
            findings.append(_unavailable(rule, reason, record))

    previous_levels, actual_groups, sequence_state, automatic_state, ancestors = {}, defaultdict(list), {}, {}, {}
    broken_scopes: set = set()

    def chapter_scope(role, level):
        """当前标题所属“章”的标识（最近一个一级标题的段落身份）。

        论文每章的子级编号通常从 1 重新开始（第二章又是 1.1、1.2…），
        因此编号序列必须**按章分组**：只按“级”分组会把每章重新开头的
        小节编号当成全局跳号，整篇刷屏误报。
        """

        if level <= 1:
            return None
        anchor = ancestors.get((role, 1))
        return anchor[0] if anchor else None

    def scope_key(role, level):
        return (role, level, chapter_scope(role, level))

    def note_number(role, level, record):
        """登记一条标题的“上一处编号”，只记账、不做判定。

        模板缺少该层级基准、模板编号不可读等情况只影响“能不能判定”，不影响
        “上一条实际编号是多少”。漏记会造成序列状态空洞，让下一条同父标题的
        编号被当成“应为 1”的跳号（论文实测中 2.2 误报的直接原因）。
        """

        values = record.get("number_values")
        if not values:
            return
        sequence_state[sequence_key(role, level, values)] = values[-1]

    def sequence_key(role, level, values):
        """编号序列的分组键：**章 + 本级编号路径**。

        论文每章的子级编号通常从 1 重新开始（第二章又是 1.1、1.2…），所以序列
        必须按章分组；同时用“父编号路径”而不是“父标题对象”作组键，中间夹入
        无编号标题时才不会把同一组编号误拆成两段。
        """

        return (role, level, chapter_scope(role, level), tuple(values[:-1]))

    for record in records:
        role, level, paragraph = record["role"], record["level"], record["paragraph"]
        # 标题树位置必须无条件登记：下面任何一个 continue 都不能让“父标题/章”
        # 身份过期，否则同一条编号会被拿去和上一章（或占位父级）比较。
        for key in [key for key in ancestors if key[0] == role and key[1] > level]:
            del ancestors[key]
        ancestors[(role, level)] = (id(paragraph._p), record.get("number_values"))
        if role == "references":
            # 待开会讨论，本版本暂不开发：参考文献相关业务。
            continue
        candidates = _candidates(data, record)
        if record.get("level_unknown"):
            uncertain("048", record["numbering_unknown"], record)
            for rule in ("049", "050", "051", "052", "053", "054", "055"):
                uncertain(rule, record["numbering_unknown"], record)
            continue
        template_levels = data["levels"].get(role, set())
        previous_level = previous_levels.get(role)
        start_level = min(template_levels) if template_levels else None
        missing = sorted(value for value in template_levels if
                         (previous_level if previous_level is not None else (start_level or level) - 1) < value < level)
        if missing and (previous_level is None or level > previous_level):
            if (previous_level, level) not in data["transitions"].get(role, set()):
                findings.append(make_finding("048", _TITLES["048"],
                                             "依模板已有层级顺序经过" + "、".join(str(x) + "级" for x in missing),
                                             f"前一同用途标题层级为{previous_level or '无'}，当前为{level}级", paragraph))
        previous_levels[role] = level
        actual_groups[(role, level)].append(record)
        if not candidates:
            for rule in ("050", "051", "052", "053", "055"):
                uncertain(rule, "模板中没有相应用途和层级的编号基准", record)
            note_number(role, level, record)
            continue
        if record.get("definition_error") or record.get("field_numbering") or record.get("revision_affected"):
            for rule in ("050", "051", "052", "053", "055"):
                uncertain(rule, record["numbering_unknown"], record)
            # 本条编号读不出来：该作用域的编号链断裂，后续标题不再判定连续性，
            # 避免把“读不到的那一条”造成的序号差误报成跳号。
            broken_scopes.add(scope_key(role, level))
            continue
        if any(item.get("definition_error") or item.get("field_numbering") or item.get("revision_affected") for item in candidates):
            for rule in ("050", "051", "052", "053", "055"):
                uncertain(rule, "模板相应标题编号不能可靠读取", record)
            note_number(role, level, record)
            continue
        if record["numbering_unknown"]:
            uncertain("050", record["numbering_unknown"], record)
        if any(item["numbering_unknown"] for item in candidates):
            uncertain("050", "模板编号的最终显示不能可靠计算，仍核对可读取的编号定义", record)
        modes = {_mode(item) for item in candidates}
        expected_numbered = {mode != "none" for mode in modes}
        numbered = bool(record["number"] or record["automatic"])
        if len(expected_numbered) != 1:
            uncertain("053", "模板同用途同级标题同时存在编号与无编号变体，不能唯一匹配", record)
            note_number(role, level, record)
            continue
        required = next(iter(expected_numbered))
        if required != numbered:
            code = "053.2" if required else "053.1"
            findings.append(make_finding(code, _title(code), "应有编号" if required else "不应编号",
                                         "未识别到编号" if not numbered else f"存在编号{record['number']}", paragraph))
            if numbered:
                # 本不该编号却出现了编号：编号含义无法确认，后续序列不据此判定。
                broken_scopes.add(scope_key(role, level))
            continue
        if not required:
            continue
        shapes = {_shape(item) for item in candidates}
        shape = _shape(record)
        shape_ok = shape in shapes
        if not shape_ok:
            findings.append(make_finding("051", _TITLES["051"],
                                         "；".join(_shape_description(item) for item in sorted(shapes, key=str)),
                                         f"编号{record['number'] or '显示值未能计算'}；组成方式{_shape_description(shape)}", paragraph))
        separators = {item["separator"] for item in candidates}
        extra = re.match(r"^[ \t\u3000]+", record["text"]) if record["automatic"] else None
        template_extras = {match[0] if match else "" for item in candidates
                           if item["automatic"] for match in [re.match(r"^[ \t\u3000]+", item["text"]) ]}
        extra_wrong = bool(extra and (extra[0] not in template_extras))
        if record["separator"] not in separators or extra_wrong:
            code = _separator_code(record, extra_wrong)
            findings.append(make_finding(code, _title(code),
                                         "或".join(_separator_description(item) for item in sorted(separators, key=str)),
                                         _separator_description(record["separator"]) +
                                         ("；标题文字前另输入" + _separator_description(extra[0]) if extra_wrong else ""), paragraph))
        if record["automatic"]:
            spec = record["numbering"]["levels"][record["numbering"]["ilvl"]]
            expected_tabs = {item["numbering"]["levels"][item["numbering"]["ilvl"]]["tabs"]
                             for item in candidates if item.get("numbering")}
            direct_tabs = any(props.find(qn("w:tabs")) is not None for props in _properties(paragraph))
            if direct_tabs:
                uncertain("052", "标题段落或样式含制表位覆盖，无法完整合并列表制表位", record)
            elif len(expected_tabs) == 1 and next(iter(expected_tabs)) and spec["tabs"] not in expected_tabs:
                findings.append(make_finding("052.2", _title("052.2"), f"编号跟随制表位{next(iter(expected_tabs))}",
                                             f"当前列表制表位{spec['tabs']}", paragraph))
            if spec["indent"] or any(item.get("numbering") and item["numbering"]["levels"][item["numbering"]["ilvl"]]["indent"] for item in candidates):
                uncertain("052", "列表缩进与段落/样式缩进的最终叠加位置尚无法完整确认", record)
        elif modes == {"automatic"} and all(item["numbering"]["multilevel"] for item in candidates):
            findings.append(make_finding("055", _TITLES["055"], "使用模板对应标题的多级列表自动编号",
                                         f"编号{record['number']}为段落中手工输入文字，未关联多级列表", paragraph))
        sequence = _expected_sequence(candidates)
        if record["automatic"]:
            if sequence is None:
                uncertain("050", "模板没有可唯一确认的起始、递增或重启规则", record)
                continue
            spec = record["numbering"]["levels"][record["numbering"]["ilvl"]]
            if spec["start"] != sequence["start"] or spec["restart"] != sequence["restart"]:
                findings.append(make_finding("050.3", _title("050.3"),
                                             f"起始值{sequence['start']}；{_restart_description(sequence['restart'])}",
                                             f"起始值{spec['start']}；{_restart_description(spec['restart'])}", paragraph))
            if record["number_known"] and shape_ok and record["number_values"]:
                list_key = (role, level, record["numbering"]["num_id"], chapter_scope(role, level))
                old = automatic_state.get(list_key)
                restart_level = sequence["restart"] + 1 if sequence["restart"] is not None else None
                parent = ancestors.get((role, restart_level)) if restart_level is not None else None
                expected = old[0] + sequence["step"] if old and old[1] == parent else sequence["start"]
                # Only report a displayed gap inside the same actual list.
                # The start-definition mismatch above already covers wrong starts.
                if old and old[1] == parent and record["number_values"][-1] != expected and sequence.get("heading_continuous"):
                    findings.append(make_finding("050.1", _title("050.1"),
                                                 f"同一列表同级标题的下一序号为{expected}",
                                                 f"当前显示编号{record['number']}，列表可能被非标题段落占用", paragraph))
                automatic_state[list_key] = (record["number_values"][-1], parent)
            # Word calculates consecutive values within a valid list instance.
            # Different numIds may intentionally be independent sequences.
        elif shape_ok and record["number_values"]:
            # 手工编号先登记“上一处编号”，再判定间隔：登记与判定分离，模板缺少
            # 该层级唯一基准（sequence 为 None）时只记账不判定，否则中间一条标题
            # 会让状态空洞，把下一条同父标题的编号误报成跳号。
            current_values = record["number_values"]
            state_key = sequence_key(role, level, current_values)
            old = sequence_state.get(state_key)
            sequence_state[state_key] = current_values[-1]
            # 模板列表层级可能和标题层级错位（如“第X章 / N / N.M”体系），序列基准
            # 由规则目录给出，这里只要求模板能唯一确认起始与步长。
            if sequence is None or scope_key(role, level) in broken_scopes:
                if sequence is None:
                    uncertain("050", "模板没有可唯一确认的起始、递增或重启规则", record)
                continue
            expected = old + sequence["step"] if old is not None else sequence["start"]
            problems = []
            if old is not None and current_values[-1] == old:
                # 050.2：同一层级再次出现相同编号（重号），与“跳号”分开报。
                findings.append(make_finding("050.2", _title("050.2"),
                                             f"同一层级章节编号应唯一（上一处为 {old}）",
                                             f"当前编号 {record['number']} 与上一处重复", paragraph))
            elif current_values[-1] != expected:
                problems.append(f"当前末级序号{current_values[-1]}，应为{expected}")
            # 编号组件同样按标题树对齐：末级组件是本级序号，其余组件从本级向上逐级对应，
            # “第X章 + 2.1”与“2.1.1”这类编号才能找到真正的父标题。
            first_component_level = level - (len(current_values) - 1)
            for offset, actual_parent in enumerate(current_values[:-1]):
                mapped_level = first_component_level + offset
                parent_record = ancestors.get((role, mapped_level)) if mapped_level >= 1 else None
                if parent_record and parent_record[1] and parent_record[1][-1] != actual_parent:
                    problems.append(f"第{mapped_level}级父编号为{actual_parent}，当前父标题为{parent_record[1][-1]}")
            if problems:
                findings.append(make_finding("050.1", _title("050.1"),
                                             f"从{sequence['start']}起按{sequence['step']}递增，并按模板父级重启",
                                             "；".join(problems), paragraph))
        elif sequence is None:
            uncertain("050", "模板没有可唯一确认的起始、递增或重启规则", record)
            continue

    # 049 supplements within-document inconsistency.  A uniformly wrong group
    # belongs to the unchanged 045--047 detectors and is not duplicated here.
    labels = {"east_asia_font": "中文字体", "ascii_font": "西文字体", "size_pt": "字号",
              "first_line_indent": "首行缩进", "left_indent": "左缩进", "right_indent": "右缩进",
              "numbering": "编号形式"}
    for (role, level), group in actual_groups.items():
        if len(group) < 2:
            continue
        signatures = []
        for record in group:
            try:
                signatures.append((record, _format_signature(record)))
            except Exception as exc:
                uncertain("049", str(exc), record)
        for record, signature in signatures:
            candidates = _candidates(data, record)
            specs = [item.get("format_signature") for item in candidates if item.get("format_signature")]
            if not specs:
                uncertain("049", "模板对应同级格式无法读取", record)
                continue
            if record.get("numbering") and record["numbering"]["levels"][record["numbering"]["ilvl"]]["indent"]:
                uncertain("049", "列表缩进与段落/样式缩进的最终叠加尚不能完整读取", record)
            wrong = []
            for key in labels:
                actual_values = {str(item[key]) for _, item in signatures if item[key] is not None}
                expected_values = {str(item[key]) for item in specs if item[key] is not None}
                if len(actual_values) > 1 and expected_values and signature[key] is not None:
                    if str(signature[key]) not in expected_values:
                        if key == "numbering":
                            expected_text = "或".join(sorted({_shape_description(item[key]) for item in specs}))
                            actual_text = _shape_description(signature[key])
                        else:
                            expected_text = "或".join(sorted(expected_values))
                            actual_text = str(signature[key])
                        wrong.append((key, f"{labels[key]}模板允许{expected_text}，当前{actual_text}"))
            # 字体 / 字号 / 缩进 / 编号形式 各自成条
            for key, text in wrong:
                code = _SAME_LEVEL_CODES.get(key, "049")
                findings.append(make_finding(code, _title(code), "同用途同级标题应采用模板允许的格式",
                                             text, record["paragraph"]))
    # Separate list instances can be intentional, but cannot silently count as
    # proven chapter continuity where the template models a single sequence.
    for (role, level), group in actual_groups.items():
        template = _candidates(data, group[0])
        expected_ids = {item["numbering"]["num_id"] for item in template if item.get("numbering")}
        actual_ids = {item["numbering"]["num_id"] for item in group if item.get("numbering")}
        if len(expected_ids) == 1 and len(actual_ids) > 1:
            uncertain("050", "模板同用途同级标题使用一个列表实例，当前使用多个实例；独立序列或意外重启无法唯一判断", group[0])
    # No python-docx API paginates a document.  Keep-with-next is evidence of
    # an intended constraint, never proof of actual separation or success.
    for (role, level), group in actual_groups.items():
        template = _candidates(data, group[0])
        if template:
            wanted = {item.get("pagination", {}).get("keep_with_next") for item in template}
            try:
                actual = effective_paragraph(group[0]["paragraph"]).get("keep_with_next")
            except Exception:
                actual = "读取异常"
            uncertain("054", f"需Word实时分页确认标题和后续内容；模板与下段同页设置{wanted}，"
                      f"当前首个同类标题设置{actual}，仅此开关不能证明是否分离（本组{len(group)}个标题）", group[0])
    return findings

def collect_heading_records(ctx, mapping=None):
    """带缓存的入口：同一份文档只扫描一次（见 core.docx_context.memo）。"""

    return memo(ctx, "heading_records", lambda: _collect_heading_records_impl(ctx, mapping))

def _toc_paragraph_ids(ctx):
    """带缓存的入口：同一份文档只扫描一次（见 core.docx_context.memo）。"""

    return memo(ctx, "toc_paragraph_ids", lambda: __toc_paragraph_ids_impl(ctx))
