"""结构完整性检测（001–011）的取证实现。

003、004、007、011：本条依赖大模型语义，本阶段暂不实现。
008 及所有参考文献引用判定：【待开会讨论，本版本暂不开发】。
The component vocabulary identifies objects; it is not a school requirement.
Only explicit template instructions establish mandatory components or counts.
"""

from __future__ import annotations

import re

from thesis_format_checker.core.xmlutil import qn

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import make_finding, unknown


_COMPONENTS = {
    "cover": ("封面",),
    "title_page": ("题名页", "扉页", "封二"),
    # 部分学校模板将原创性声明页简写成独立标题“声 明”。这里只在
    # _heading_components 的精确标题路径使用，正文里的普通“声明”不会命中。
    "originality": ("原创性声明", "独创性声明", "声明"),
    "authorization": ("授权使用声明", "使用授权声明", "论文使用授权声明", "学位论文使用授权声明", "学位论文版权使用授权书", "版权使用授权书"),
    "abstract_zh": ("中文摘要", "摘要"),
    "abstract_en": ("英文摘要", "abstract"),
    "toc": ("目录",),
    "references": ("参考文献", "references", "bibliography"),
    "acknowledgements": ("致谢",),
    "appendix": ("附录",),
    "achievements": ("攻读学位期间成果", "攻读学位期间取得的成果", "攻读学位期间的研究成果", "攻读学位期间发表的论文", "攻读硕士学位期间取得的研究成果", "攻读博士学位期间取得的研究成果"),
}
#: 组成部分 → 细粒度规则编号（标题统一从规则目录取，避免两处维护文案）。
#: “原创性声明/独创性声明”是同一份声明的两种叫法，因此合为一件事（002.1）。
_COMPONENT_CODES = {
    "cover": "001.1",
    "title_page": "001.2",
    "originality": "002.1",
    "authorization": "002.2",
    "acknowledgements": "009.1",
    "appendix": "009.2",
    "achievements": "009.3",
}
#: 声明类组件：标题存在但后面没有正文，单独作为一条规则
_DECLARATION_COMPONENTS = ("originality", "authorization")
_DECLARATION_CONTENT_RULE = "002.3"
#: 关键词：中/英文各自分开的“缺少”与“数量不符”
_KEYWORD_CODES = {
    "zh": {"missing": "005.1", "count": "005.2"},
    "en": {"missing": "005.3", "count": "005.4"},
}
_KEYWORD_CONTENT_RULE = "005.5"
#: 目录：缺少 / 手工目录 / 生成来源不符
_TOC_CODES = {"missing": "006.1", "manual": "006.2", "source": "006.3"}
#: 组成部分排列顺序
_ORDER_RULE = "010.1"
_CHAPTER_CODES = {"empty": "011.1", "duplicate": "011.2"}
_TITLES = {
    "001": "封面或题名页缺失",
    "002": "声明缺失或只有标题",
    "005": "关键词缺失或数量不符",
    "006": "目录缺失或生成方式不符",
    "009": "必需组成部分缺失",
    "010": "组成部分排列顺序不符",
    "011": "存在无内容的章节、重复章节或明显遗漏的章节",
}
_MANDATORY = re.compile(r"必须(?:包含|包括|设置|提供|有|填写|保留)|(?:应|须|需)(?:当)?(?:包含|包括|设置|提供|有|填写|保留)|不得(?:缺少|省略)|不可(?:缺少|省略)|(?:必需|必备|必填|必选)(?:部分|项)?")
_OPTIONAL = re.compile(r"(?:可以|可)(?:以)?(?:省略|不设|不填|选)|(?:非|不)(?:必需|必备|必填|必选)|(?:不|无)(?:须|需)(?:要)?|视(?:情况|需要)|如有|若有|必要时|有.{0,12}时")
_ORDER = re.compile(r"(?:组成部分|论文结构|前置部分|论文各部分|各部分|结构|排列|章节|组成)?(?:排列)?顺序|依次(?:为|包括|包含)|按.{0,8}顺序")
#: 后置部分的固定次序（用于把“缺少附录/致谢/成果”的批注放在它应处的位置）
_BACK_MATTER_ORDER = ("references", "appendix", "acknowledgements", "achievements")
_KEYWORD_LABEL = re.compile(r"^\s*(中文关键词|英文关键词|关键词|key\s*words?)\s*[:：]\s*(.*)$", re.I)
_KEYWORD_COUNT = re.compile(r"(?:数量\s*[:：为是]?\s*|(?:应|须|需)(?:为|有|包含|包括|填写|设置)|(?:为|包含|包括))\s*(\d+)\s*(?:[-—–~～至到]\s*(\d+)\s*)?(?:个|项)")
_READ_LIMIT = "该对象docx库无法读取，本条暂无法执行检测"
_TOC_PAGE_SUFFIX = re.compile(r"[\t\u3000 ]+[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩivxlcdmIVXLCDM0-9Xx]{1,6}\s*$")


def _compact(text):
    return re.sub(r"\s+", "", text).lower().strip("：:。.")


def _title(code: str) -> str:
    """细粒度规则的标题统一取自规则目录（``checks/rule_specs.py``）。"""

    return FINE_RULES[code].title


def _mentioned(text):
    """Mention matching is only used inside an explicit instruction."""
    text = _compact(text)
    found = []
    for key, aliases in _COMPONENTS.items():
        matches = [(text.find(_compact(alias)), alias) for alias in aliases if _compact(alias) in text]
        # A generic Chinese 摘要 must not also match 英文摘要.
        if key == "abstract_zh" and "英文摘要" in text and "中文摘要" not in text:
            text_without_english = text.replace("英文摘要", "")
            matches = [(text.find("摘要"), "摘要")] if "摘要" in text_without_english else []
        if matches:
            found.append((min(item[0] for item in matches), key))
    return [key for _, key in sorted(found)]


def _heading_components(text):
    text = _compact(text)
    if not text or _MANDATORY.search(text) or _ORDER.search(text):
        return []
    result = []
    for key, aliases in _COMPONENTS.items():
        if any(text == _compact(alias) for alias in aliases):
            result.append(key)
        elif key in ("originality", "authorization") and len(text) < 65:
            # A combined declaration heading is permitted; content is checked below.
            if any(_compact(alias) in text for alias in aliases) and not re.search(r"[:：，,；;。]", text):
                result.append(key)
        elif key == "appendix" and re.fullmatch(r"附录(?:[a-z0-9一二三四五六七八九十]+)(?:[、.．:：].*)?", text):
            result.append(key)
        elif key == "title_page" and "封二" in text and len(text) < 40:
            result.append(key)
        elif key == "cover" and "封面格式" in text and len(text) < 40:
            result.append(key)
    return result


def _toc_level(paragraph):
    """目录缓存段落层级；目录文字只能用于模板顺序，不能充当论文实际组成部分。"""

    try:
        style = paragraph.style.name or ""
    except Exception:
        return None
    match = re.match(r"^(?:toc|目录)\s*([1-9])$", style.strip(), re.I)
    return int(match.group(1)) if match else None


def _toc_component_order(ctx):
    """读取模板目录缓存中的组成部分顺序，例如“致谢→参考文献→附录”。"""

    order = []
    for paragraph in ctx.paragraphs:
        if _toc_level(paragraph) is None:
            continue
        text = _TOC_PAGE_SUFFIX.sub("", paragraph.text or "")
        for key in _heading_components(text):
            if key not in order:
                order.append(key)
    return order


def _component_positions(ctx):
    positions = {}
    for index, paragraph in enumerate(ctx.paragraphs):
        if _toc_level(paragraph) is not None:
            continue
        text = paragraph.text
        for key in _heading_components(text):
            positions.setdefault(key, index)
    # Plain covers frequently use table cells, with no printed heading "封面".
    # Require independent label classes, not a particular number of paragraphs.
    stop = min((i for key, i in positions.items() if key != "cover"), default=len(ctx.texts))
    candidates = ctx.texts[:stop]
    title_indices = [i for i, t in enumerate(candidates) if re.match(r"\s*(?:论文题目|论文题名|题目|题名)\s*[:：]", t)]
    author_indices = [i for i, t in enumerate(candidates) if re.match(r"\s*(?:作者姓名|研究生姓名|学生姓名|姓名|作者|学号)\s*[:：]", t)]
    if title_indices and author_indices:
        positions.setdefault("cover", min(title_indices + author_indices))
    # 封一与封二必须分开取证。封一用分类号/密级/UDC 组合识别；封二用
    # “作者：+指导教师：”组合识别，避免只剩一页时仍同时判两者存在。
    front = [(i, re.sub(r"\s+", "", text)) for i, text in enumerate(ctx.texts[:140])]
    cover_indices = [i for i, text in front
                     if re.search(r"分类号.{0,20}密级", text) or "UDC" in text.upper()]
    author_page = [i for i, text in front if re.search(r"作者[:：]", text)]
    supervisor_page = [i for i, text in front if re.search(r"指导教师[:：]", text)]
    if cover_indices:
        positions.setdefault("cover", min(cover_indices))
    if author_page and supervisor_page:
        positions.setdefault("title_page", min(author_page + supervisor_page))
    return positions


def _field_instructions(ctx):
    """Inspect field XML without requesting an update or changing its result."""
    fields, stack = [], []
    malformed = False
    for node in ctx.document.element.body.iter():
        if node.tag == qn("w:fldSimple"):
            fields.append(node.get(qn("w:instr"), ""))
        elif node.tag == qn("w:fldChar"):
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                stack.append({"parts": [], "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                fields.append("".join(stack.pop()["parts"]))
            elif kind in ("separate", "end"):
                malformed = True
        elif node.tag == qn("w:instrText"):
            if stack and not stack[-1]["separated"]:
                stack[-1]["parts"].append(node.text or "")
            elif not stack:
                malformed = True
    malformed = malformed or bool(stack)
    return fields, malformed


def _has_order_cycle(pairs):
    outgoing = {}
    for before, after in pairs:
        outgoing.setdefault(before, set()).add(after)
    visiting, visited = set(), set()

    def visit(key):
        if key in visiting:
            return True
        if key in visited:
            return False
        visiting.add(key)
        if any(visit(other) for other in outgoing.get(key, ())):
            return True
        visiting.remove(key)
        visited.add(key)
        return False

    return any(visit(key) for key in outgoing)


def _toc_modes(instructions):
    modes = []
    for instruction in instructions:
        if not re.match(r"^\s*TOC(?:\s|$)", instruction, re.I):
            continue
        if re.search(r"\\[out](?:\s|$)", instruction, re.I):
            modes.append("headings")
        elif re.search(r"\\[fc](?:\s|$)", instruction, re.I):
            modes.append("other")
        else:
            # TOC without source switches uses Word's built-in headings.
            modes.append("headings")
    return modes


def _keyword_language(label):
    return "en" if label.lower().startswith(("英文", "key")) else "zh"


def extract_structure_baseline(ctx):
    """Return new namespaced evidence; never replace legacy template keys."""
    data = {"required": {}, "optional": set(), "conditional": set(), "samples": {},
            "order": [], "order_conflict": False, "keyword": {}, "toc_modes": [],
            "toc_heading_required": False, "toc_unreadable": False, "errors": []}
    try:
        data["samples"] = _component_positions(ctx)
        orders = []
        for paragraph in ctx.texts:
            # Sentences and semicolons separate optional/mandatory declarations.
            for clause in re.split(r"[。；;\n]", paragraph):
                mentions = _mentioned(clause)
                if _OPTIONAL.search(clause):
                    destination = "conditional" if re.search(r"如有|若有|必要时|有.{0,12}时", clause) else "optional"
                    data[destination].update(mentions)
                elif _MANDATORY.search(clause):
                    for key in mentions:
                        data["required"].setdefault(key, []).append(clause.strip())
                if _ORDER.search(clause) and len(mentions) > 1:
                    if not _OPTIONAL.search(clause):
                        orders.append(mentions)
                if re.search(r"关键词|key\s*words?", clause, re.I):
                    # Sample keyword values are never interpreted as counts.
                    count_match = _KEYWORD_COUNT.search(clause)
                    scoped = []
                    if "中文" in clause:
                        scoped.append("zh")
                    if "英文" in clause or re.search(r"key\s*words?", clause, re.I):
                        scoped.append("en")
                    if not scoped:
                        scoped = ["zh", "en"]
                    required = bool(_MANDATORY.search(clause) or count_match) and not _OPTIONAL.search(clause)
                    if required:
                        for lang in scoped:
                            entry = data["keyword"].setdefault(lang, {"required": False, "counts": [], "sources": []})
                            entry["required"] = True
                            entry["sources"].append(clause.strip())
                            if count_match:
                                low = int(count_match.group(1))
                                high = int(count_match.group(2) or low)
                                entry["counts"].append((low, high))
                if "目录" in clause and re.search(r"(?:自动生成|自动目录|标题样式|大纲级别)", clause):
                    if _MANDATORY.search(clause) or re.search(r"目录\s*(?:须|应|需|必须)", clause):
                        data["required"].setdefault("toc", []).append(clause.strip())
                        data["toc_heading_required"] = bool(re.search(r"标题样式|大纲级别", clause))
        if orders:
            # Multiple disjoint order requirements remain separate constraints.
            data["order"] = orders
            pairs = {(a, b) for seq in orders for i, a in enumerate(seq) for b in seq[i + 1:]}
            data["order_conflict"] = _has_order_cycle(pairs)
        elif len(data["samples"]) > 1:
            # 规范模板自身的实际排列就是默认顺序；无需再要求模板额外写一句
            # “依次为……”。模板是撰写规范时，后置部分的正文示例可能按附件顺序
            # 摆放，因此优先采用自动目录缓存里的后置部分顺序；实际组成部分扫描
            # 同时排除 TOC 条目，防止“目录里的附录”冒充正文附录。
            physical = [key for key, _ in sorted(
                data["samples"].items(), key=lambda item: item[1])]
            toc_order = _toc_component_order(ctx)
            if toc_order:
                front_order = [key for key in physical if key not in _BACK_MATTER_ORDER]
                data["order"] = [sequence for sequence in (front_order, toc_order)
                                 if len(sequence) > 1]
            else:
                data["order"] = [physical]
        instructions, malformed = _field_instructions(ctx)
        data["toc_modes"] = _toc_modes(instructions)
        data["toc_unreadable"] = malformed and ("toc" in data["samples"] or "toc" in data["required"])
        for index, paragraph in enumerate(ctx.texts):
            match = _KEYWORD_LABEL.match(paragraph)
            if match:
                data["keyword"].setdefault(_keyword_language(match.group(1)), {"required": False, "counts": [], "sources": [], "sample": index})
    except Exception as exc:
        data["errors"].append(f"模板结构读取异常（{type(exc).__name__}）；{_READ_LIMIT}")
    data["optional"] = sorted(data["optional"])
    data["conditional"] = sorted(data["conditional"])
    return {"structure": data}


def _anchor(ctx, position=None):
    if not ctx.paragraphs:
        return None
    if position is None:
        return ctx.paragraphs[0]
    return ctx.paragraphs[min(max(position, 0), len(ctx.paragraphs) - 1)]


def _missing_anchor(ctx, data, positions, key):
    """缺失组成部分的批注锚点：落在“它应该出现的位置”，不落在封面。

    优先按模板声明的前后顺序找相邻组成部分；推不出来时按后置部分的固定次序
    （参考文献 → 附录 → 致谢 → 攻读成果）在论文已有部分里找参照；再不行就落在
    正文最后一段——绝不能因为找不到位置就把这类结论锚到第 1 段（封面）。
    """

    for sequence in data["order"]:
        if key in sequence:
            offset = sequence.index(key)
            for next_key in sequence[offset + 1:]:
                if next_key in positions:
                    return _anchor(ctx, positions[next_key])
            for previous_key in reversed(sequence[:offset]):
                if previous_key in positions:
                    return _anchor(ctx, positions[previous_key])
    # 按后置部分次序推算应处位置：缺失项的前一个/后一个已有部分就是它的邻居。
    if key in _BACK_MATTER_ORDER:
        offset = _BACK_MATTER_ORDER.index(key)
        following = [name for name in _BACK_MATTER_ORDER[offset + 1:] if name in positions]
        if following:
            return _anchor(ctx, positions[following[0]])
        preceding = [name for name in reversed(_BACK_MATTER_ORDER[:offset]) if name in positions]
        if preceding:
            return _anchor(ctx, positions[preceding[0]])
    return _last_body_anchor(ctx)


def _last_body_anchor(ctx):
    """正文最后一段（后置部分缺失时用它作为“应该在文末”的位置）。"""

    for paragraph in reversed(ctx.paragraphs):
        if paragraph.text.strip():
            return paragraph
    return _anchor(ctx)


def _content_after(ctx, positions, key):
    start = positions[key]
    end = min((value for value in positions.values() if value > start), default=len(ctx.paragraphs))
    for paragraph in ctx.paragraphs[start + 1:end]:
        try:
            style_name = (paragraph.style.name or "").lower()
            if re.match(r"(?:heading\s*\d+|标题\s*\d+)$", style_name):
                return False
        except (AttributeError, KeyError):
            pass
        if paragraph.text.strip() or paragraph._p.xpath(".//w:drawing | .//w:pict | .//m:oMath"):
            return True
    return False


def _check_components(data, ctx, positions):
    """必需组成部分检查：一个组成部分一条结论。"""

    findings = []
    unresolved = []
    for key, code in _COMPONENT_CODES.items():
        name = _COMPONENTS[key][0]
        title = _title(code)
        required = key in data["required"]
        conflict = required and key in data["optional"]
        if conflict or key in data["conditional"]:
            unresolved.append((code, f"{name}的模板要求冲突或附带尚未确认的适用条件"))
            continue
        if key not in positions:
            if required:
                anchor = _missing_anchor(ctx, data, positions, key)
                if ctx.has_textboxes:
                    findings.append(unknown(code, title, f"未在正文或表格识别到{name}，文档含文本框，可能存在于其中；{_READ_LIMIT}", paragraph=anchor))
                else:
                    findings.append(make_finding(code, title, f"必须包含{name}（{'；'.join(data['required'][key])}）", f"未识别到{name}", paragraph=anchor))
            elif key in data["samples"] and key not in data["optional"]:
                unresolved.append((code, f"模板有{name}示例，但未明确其必需性，不能据此判定缺失"))
        elif key in _DECLARATION_COMPONENTS and not _content_after(ctx, positions, key):
            if required or key in data["samples"]:
                content_title = _title(_DECLARATION_CONTENT_RULE)
                if ctx.has_textboxes:
                    findings.append(unknown(_DECLARATION_CONTENT_RULE, content_title, f"仅识别到{name}标题，文档含文本框，无法确认其正文；{_READ_LIMIT}", paragraph=_anchor(ctx, positions[key])))
                else:
                    findings.append(make_finding(_DECLARATION_CONTENT_RULE, content_title, f"{name}应包含声明正文", "仅识别到声明标题，未识别到后续声明内容", paragraph=_anchor(ctx, positions[key])))
    for code, reason in unresolved:
        findings.append(unknown(code, _title(code), reason, paragraph=_anchor(ctx)))
    return findings


def _check_keywords(data, ctx):
    """关键词检查：中/英文分开，“缺少”与“数量不符”各自成条。"""

    findings = []
    actual = {}
    for index, text in enumerate(ctx.texts):
        match = _KEYWORD_LABEL.match(text)
        if match:
            actual.setdefault(_keyword_language(match.group(1)), []).append((index, match.group(2).strip()))
    for lang, requirements in data["keyword"].items():
        rules = _KEYWORD_CODES[lang]
        label = "英文关键词" if lang == "en" else "中文关键词"
        entries = actual.get(lang, [])
        counts = set(requirements["counts"])
        if len(counts) > 1 or any(low > high for low, high in counts):
            findings.append(unknown(rules["count"], _title(rules["count"]),
                                    f"模板对{label}的数量要求冲突，无法选定基准", paragraph=_anchor(ctx)))
            continue
        if len(entries) > 1:
            findings.append(unknown(rules["missing"], _title(rules["missing"]),
                                    f"识别到多组{label}，无法确定适用的摘要区域", paragraph=_anchor(ctx, entries[0][0])))
            continue
        if not entries:
            code = rules["missing"]
            if requirements["required"]:
                if ctx.has_textboxes:
                    findings.append(unknown(code, _title(code), f"未识别到{label}，文档含文本框；{_READ_LIMIT}", paragraph=_anchor(ctx)))
                else:
                    findings.append(make_finding(code, _title(code), f"包含{label}；{'；'.join(requirements['sources'])}", f"未识别到{label}", paragraph=_anchor(ctx)))
            elif "sample" in requirements:
                # 模板对照系统以模板中实际出现的中/英文关键词区为结构基准。
                # “样例值不能推导数量”仍保持不变，但整块关键词缺失可以确定判断；
                # 否则一份同时给出中英文摘要与关键词的规范模板反而无法发现缺项。
                findings.append(make_finding(
                    code, _title(code), f"包含{label}（模板存在对应关键词区）",
                    f"未识别到{label}", paragraph=_anchor(ctx)))
            else:
                findings.append(unknown(code, _title(code), f"模板未提供可确认的{label}要求", paragraph=_anchor(ctx)))
            continue
        index, value = entries[0]
        if not value:
            code = _KEYWORD_CONTENT_RULE
            in_table = bool(ctx.paragraphs[index]._p.xpath("ancestor::w:tc"))
            if ctx.has_textboxes or in_table:
                findings.append(unknown(code, _title(code), f"{label}标签后的内容为空，可能使用跨单元格或文本框填写，无法可靠确认关键词区域；{_READ_LIMIT}", paragraph=_anchor(ctx, index)))
            else:
                findings.append(make_finding(code, _title(code), f"{label}应填写有效内容", "关键词标签后为空", paragraph=_anchor(ctx, index)))
            continue
        if not counts:
            continue
        # English spaces stay inside a phrase. Punctuation separates entries.
        if not re.search(r"[;；,，、]", value) and re.search(r"\s{2,}|\t", value):
            code = rules["count"]
            findings.append(unknown(code, _title(code), f"{label}使用空白分隔，无法可靠区分词组与关键词边界", paragraph=_anchor(ctx, index)))
            continue
        words = [part.strip().strip("。.") for part in re.split(r"[;；,，、]", value) if part.strip().strip("。.")]
        low, high = next(iter(counts))
        if not low <= len(words) <= high:
            code = rules["count"]
            expected = f"{label}数量为{low}个" if low == high else f"{label}数量为{low}—{high}个"
            findings.append(make_finding(code, _title(code), expected, f"识别到{len(words)}个：{value}", paragraph=_anchor(ctx, index)))
    return findings


def _check_toc(data, ctx, positions):
    """目录检查：缺少目录 / 手工目录 / 生成来源不符，各自成条。"""

    relevant = "toc" in data["samples"] or "toc" in data["required"] or data["toc_modes"]
    if not relevant:
        return []
    anchor = _anchor(ctx, positions.get("toc"))
    missing_code, manual_code, source_code = _TOC_CODES["missing"], _TOC_CODES["manual"], _TOC_CODES["source"]
    if data["toc_unreadable"]:
        return [unknown(manual_code, _title(manual_code), f"模板目录域结构不完整，无法确认目录生成方式；{_READ_LIMIT}", paragraph=anchor)]
    instructions, malformed = _field_instructions(ctx)
    modes = _toc_modes(instructions)
    if malformed:
        return [unknown(manual_code, _title(manual_code), f"文档包含不完整的域，无法确认目录生成结构；{_READ_LIMIT}", paragraph=anchor)]
    if "toc" not in positions and not modes:
        if ctx.has_textboxes:
            return [unknown(missing_code, _title(missing_code), f"未识别到目录，文档含文本框；{_READ_LIMIT}", paragraph=anchor)]
        if "toc" in data["required"] and "toc" not in data["optional"] and "toc" not in data["conditional"]:
            return [make_finding(missing_code, _title(missing_code), "模板明确要求目录", "未识别到目录标题或目录域", paragraph=anchor)]
        return [unknown(missing_code, _title(missing_code), "模板有目录示例，但未明确目录的必需性，不能据此判定缺失", paragraph=anchor)]
    expected_auto = bool(data["toc_modes"] or data["toc_heading_required"] or any("自动" in s for s in data["required"].get("toc", [])))
    if expected_auto and not modes:
        return [make_finding(manual_code, _title(manual_code), "模板要求自动生成的目录域", "识别到目录文字，但没有TOC目录域", paragraph=anchor)]
    expected_heading = data["toc_heading_required"] or (bool(data["toc_modes"]) and set(data["toc_modes"]) == {"headings"})
    if expected_heading and modes and "headings" not in modes:
        return [make_finding(source_code, _title(source_code), "由标题样式或大纲级别生成目录", "目录域使用其他来源，未识别到标题或大纲来源", paragraph=anchor)]
    return []


def _check_order(data, ctx, positions):
    if data["order_conflict"]:
        return [unknown(_ORDER_RULE, _title(_ORDER_RULE), "模板中的组成部分先后要求互相冲突，无法确定顺序基准")]
    if not data["order"]:
        shared = [key for key in positions if key in data["samples"]]
        if len(shared) > 1:
            return [unknown(_ORDER_RULE, _title(_ORDER_RULE), "模板提供组成部分示例，但未明确排列顺序要求，本条未作通过判定")]
        return []
    findings = []
    emitted = set()
    for sequence in data["order"]:
        present = [key for key in sequence if key in positions]
        for index, key in enumerate(present):
            for before in present[:index]:
                if positions[before] > positions[key] and (before, key) not in emitted:
                    emitted.add((before, key))
                    first, second = _COMPONENTS[before][0], _COMPONENTS[key][0]
                    findings.append(make_finding(_ORDER_RULE, _title(_ORDER_RULE), f"{first}应位于{second}之前", f"{first}位于{second}之后", paragraph=_anchor(ctx, positions[key])))
    return findings


def _chapter_records(ctx):
    """按标题层级列出章节，并统计每个章节自己的正文内容量。

    标题级别来自标题层级树（见 ``detectors/headings.py``），因此被误用标题样式的
    正文段落不会被当成章节。
    """

    from thesis_format_checker.detectors.headings import _heading_level

    records = []
    for index, paragraph in enumerate(ctx.paragraphs):
        level = _heading_level(paragraph)
        if not level:
            continue
        records.append({"index": index, "level": level, "paragraph": paragraph,
                        "text": paragraph.text.strip(), "content": 0})
    for position, record in enumerate(records):
        end = records[position + 1]["index"] if position + 1 < len(records) else len(ctx.paragraphs)
        for paragraph in ctx.paragraphs[record["index"] + 1:end]:
            if paragraph.text.strip() or paragraph._p.xpath(".//w:drawing | .//w:pict | .//w:object"):
                record["content"] += 1
    return records


def _check_chapters(ctx):
    """011.1 无内容章节 / 011.2 同层级重复章节（只比对论文自身结构）。"""

    records = _chapter_records(ctx)
    if not records:
        return [unknown("011", _title("011"), "文档中没有可识别的章节标题", _anchor(ctx))]
    findings = []
    # 子树内容 = 自己的正文 + 所有下级章节的正文；只有整棵子树都空才算空章节。
    for position in range(len(records) - 1, -1, -1):
        record = records[position]
        total = record["content"]
        for following in records[position + 1:]:
            if following["level"] <= record["level"]:
                break
            total += following["content"]
        record["subtree"] = total
    for record in records:
        if record["subtree"] == 0:
            title = _title(_CHAPTER_CODES["empty"])
            findings.append(make_finding(_CHAPTER_CODES["empty"], title,
                                         "章节标题下应有正文内容",
                                         f"“{record['text'][:30]}”及其下级章节均没有正文内容",
                                         record["paragraph"]))
    # “引言”“本章小结”“实验结果与分析”等标题允许在不同章下重复；只有
    # 同一父级路径下出现同层级同名标题，或一级章标题重复，才是真正的重复章节。
    seen: dict[tuple[int, tuple[str, ...], str], int] = {}
    ancestors: dict[int, str] = {}
    for record in records:
        level = record["level"]
        ancestors = {item_level: text for item_level, text in ancestors.items()
                     if item_level < level}
        parent_path = tuple(ancestors[item_level] for item_level in sorted(ancestors))
        key = (level, parent_path, _compact(record["text"]))
        if key in seen:
            title = _title(_CHAPTER_CODES["duplicate"])
            findings.append(make_finding(_CHAPTER_CODES["duplicate"], title,
                                         "同一层级不应重复出现同名章节标题",
                                         f"“{record['text'][:30]}”与第 {seen[key] + 1} 段的标题重复",
                                         record["paragraph"]))
        else:
            seen[key] = record["index"]
        ancestors[level] = _compact(record["text"])
    return findings


def check_structure_errors(baseline, ctx):
    """Independent rule failures degrade to comments, never abort compare."""
    data = baseline.get("structure", {})
    if not data:
        return []
    if data.get("errors"):
        return [unknown("001", "结构基准读取受限", "；".join(data["errors"]), paragraph=_anchor(ctx))]
    try:
        positions = _component_positions(ctx)
    except Exception as exc:
        return [unknown("001", "结构对象读取受限", f"{type(exc).__name__}；{_READ_LIMIT}", paragraph=_anchor(ctx))]
    findings = []
    jobs = (
        ("001", lambda: _check_components(data, ctx, positions)),
        ("005", lambda: _check_keywords(data, ctx)),
        ("006", lambda: _check_toc(data, ctx, positions)),
        ("010", lambda: _check_order(data, ctx, positions)),
        ("011", lambda: _check_chapters(ctx)),
    )
    for rule, check in jobs:
        try:
            findings.extend(check())
        except Exception as exc:
            findings.append(unknown(rule, _TITLES[rule], f"读取失败（{type(exc).__name__}）；{_READ_LIMIT}", paragraph=_anchor(ctx)))
    return findings
