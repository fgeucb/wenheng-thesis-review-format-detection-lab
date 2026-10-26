"""结构完整性检测（001–011 的“组成部分/章节是否齐全”部分）。

设计思路：**模板就是标准**。
================================

真实学校模板是一份排版好的样张，里面往往没有“必须包含：封面、题名页……”这类
说明文字（实测某校模板 170 段里 0 条），因此旧实现只能对每个对象报“无法判定”。

本模块改为**从模板自身提取它的组成部分清单**，三种来源按优先级合并：

1. **目录域缓存**（``toc 1`` / ``toc 2`` 样式段落）——模板自己的章节清单，
   信息最完整（实测模板 25 条），是判断“需要哪些章节”的主要依据；
2. **一级标题**（``Heading 1`` / ``标题 1`` 样式，或大纲级别 0–2 的段落）；
3. **组成部分标签**——封面、题名页、声明、摘要、ABSTRACT、目录、参考文献、
   附录、成果、致谢等固定部件的标题文字。

然后把论文按同样的规则扫一遍，逐个比对：模板里有、论文里没有 → 报“缺少”。

与旧模块的分工
==============

* 本模块只负责“**有没有**”（组成部分是否存在），以及“模板列出的章节是否齐全”；
* 旧 ``structure.py`` 继续负责它擅长的部分：模板里**明确写出**要求时的判定、
  关键词数量、声明正文是否为空、组成部分排列顺序。
* 为避免重复报错：如果旧模块已经从模板的“必须包含”文字里判定过同一个组成部分，
  本模块就不再重复输出。
"""

from __future__ import annotations

import re
from typing import Any

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import make_finding
from thesis_format_checker.core.xmlutil import qn

#: 本模块负责的规则族（供规则目录自检识别）
RULE_TITLES = {
    "001": "缺少封面或题名页",
    "002": "缺少声明",
    "003": "缺少中文摘要",
    "004": "缺少英文摘要",
    "006": "缺少目录",
    "007": "正文缺少模板列出的章节",
    "008": "缺少参考文献",
    "009": "缺少致谢、附录或成果",
}

#: 组成部分 → 细粒度规则编号
_COMPONENT_CODES = {
    "cover": "001.1",
    "title_page": "001.2",
    "declaration": "002.1",
    "authorization": "002.2",
    "abstract_zh": "003.1",
    "abstract_en": "004.1",
    "toc": "006.1",
    "references": "008.1",
    "acknowledgements": "009.1",
    "appendix": "009.2",
    "achievements": "009.3",
}

#: 组成部分标签（同一部件的多种写法）；顺序即匹配优先级
_COMPONENT_LABELS: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("achievements", ("攻读学位期间", "学术成果", "研究成果", "发表的论文", "成果目录")),
    ("acknowledgements", ("致谢", "acknowledgement", "acknowledgment")),
    ("appendix", ("附录", "appendix")),
    ("references", ("参考文献", "references", "bibliography")),
    ("abstract_en", ("abstract", "英文摘要")),
    ("abstract_zh", ("中文摘要", "摘要")),
    ("declaration", ("原创性声明", "独创性声明")),
    ("authorization", ("授权使用声明", "使用授权声明", "版权使用授权", "授权书")),
    ("title_page", ("题名页", "扉页", "封二")),
    ("cover", ("封面",)),
)

#: 目录标题的精确写法（避免把“学术成果目录”当成目录）
_TOC_TEXTS = {"目录", "目次", "contents", "tableofcontents"}
#: 章节标题：第X章 / 第X篇
_CHAPTER = re.compile(r"^第\s*([0-9０-９一二三四五六七八九十百零〇Xx]+)\s*[章篇](?P<tail>.*)$")
#: 自动目录常见的“1<TAB>绪论<TAB>9”；只取一级目录项用于正文缺章核对。
_NUMBERED_TOC_HEADING = re.compile(
    r"^\s*[0-9０-９]+(?:[.．][0-9０-９]+)*[\t\u3000 ]+(?P<tail>.+?)\s*$"
)
#: 目录条目里的页码（制表符/空格 + 罗马数字或数字/XX）
_PAGE_SUFFIX = re.compile(r"[\t\u3000 ]+[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩivxlcdmIVXLCDM0-9Xx]{1,6}\s*$")
#: 模板提示语：全角括号里写着“三号黑体”“居中”等排版说明
_TEMPLATE_HINT = re.compile(r"[（(][^（()）]{0,30}(?:黑体|宋体|楷体|仿宋|字号|磅|居中|左对齐|右对齐|两端对齐)[^（()）]{0,30}[）)]")


def _normalize(text: str) -> str:
    """标准化标题文字：去空白/提示语/页码，统一大小写，便于比对。"""

    value = _PAGE_SUFFIX.sub("", text or "")
    value = _TEMPLATE_HINT.sub("", value)
    value = re.sub(r"[\s\u3000]+", "", value)
    return value.casefold().strip("：:。.、")


def _placeholder(text: str) -> bool:
    """模板占位标题（第X章 XXX、1 XXX 之类）不参与文字比对。"""

    return bool(re.fullmatch(r"[xX×]+|[xX\s×]{2,}", text)) or text in {"xxx", "xx", "x", ""}


def _title(code: str) -> str:
    item = FINE_RULES.get(code)
    return item.title if item else code


def _display_label(text: str) -> str:
    """清单条目在批注里的显示文字：去掉页码与排版提示语。"""

    return re.sub(r"[\s\u3000]+", " ", _PAGE_SUFFIX.sub("", text or "")).strip() or "（未命名）"


def _style_level(paragraph) -> int | None:
    """段落标题层级：样式名 Heading 1/标题 2 → 1/2；否则看大纲级别。"""

    style = ""
    try:
        style = paragraph.style.name or ""
    except Exception:
        style = ""
    match = re.match(r"^(?:heading|标题|标题)\s*([1-9])", style, re.I)
    if match:
        return int(match.group(1))
    props = getattr(paragraph._p, "pPr", None)
    if props is not None and props.outlineLvl is not None:
        try:
            return int(props.outlineLvl.get(qn("w:val"))) + 1
        except (TypeError, ValueError):
            return None
    return None


def _toc_level(paragraph) -> int | None:
    """目录条目层级：``toc 1`` / ``目录 2`` → 1/2。"""

    try:
        style = paragraph.style.name or ""
    except Exception:
        return None
    match = re.match(r"^(?:toc|目录)\s*([1-9])$", style.strip(), re.I)
    return int(match.group(1)) if match else None


def _component_of(text: str) -> str | None:
    """判断一段标题文字属于哪个固定组成部分。"""

    value = _normalize(text)
    if not value:
        return None
    if value in _TOC_TEXTS:
        return "toc"
    if value == "声明":
        return "declaration"
    for key, labels in _COMPONENT_LABELS:
        if key == "toc":
            continue
        if any(label in value for label in labels):
            return key
    return None


def _has_cover_block(ctx) -> bool:
    """模板前部是否存在“封面/题名页”式字段区（无标题也能识别）。"""

    field = re.compile(r"(论文题目|题目|题名|作者姓名|姓名|学号|指导教师|学院|专业|班级)")
    for paragraph in ctx.paragraphs[:40]:
        text = paragraph.text.strip()
        if text and field.search(text):
            return True
        level = _style_level(paragraph)
        if level == 1 and paragraph.text.strip() and _component_of(paragraph.text) is None:
            continue
        try:
            style = (paragraph.style.name or "").casefold()
        except Exception:
            style = ""
        if style in {"title", "标题"} and text:
            return True
    return False


def _front_matter_parts(ctx) -> set[str]:
    """识别封一和封二，避免“还剩一张封面”把两项都判为存在。

    封一通常含分类号/密级/UDC 和分字排版的“硕 士 学 位 论 文”；封二通常含
    “作者：”“指导教师：”以及中英文题名页。这里只使用字段组合，不依赖学校名称。
    """

    texts = [re.sub(r"\s+", "", paragraph.text) for paragraph in ctx.paragraphs[:90]]
    joined = "\n".join(texts)
    parts: set[str] = set()
    if (re.search(r"分类号.{0,20}密级", joined) or "UDC" in joined.upper()) and "学位论文" in joined:
        parts.add("cover")
    if (re.search(r"作者[:：]", joined) and re.search(r"指导教师[:：]", joined)) or (
            "MasterDissertation".casefold() in joined.casefold() and "Supervisedby".casefold() in joined.casefold()):
        parts.add("title_page")
    return parts


def extract_outline_baseline(ctx) -> dict[str, Any]:
    """从模板提取“必需组成部分/章节”清单（不修改模板）。"""

    components: list[dict[str, Any]] = []
    seen: set[str] = set()

    def add(key: str, label: str, tail: str = "", source: str = "") -> None:
        token = f"{key}:{tail}" if key == "chapter" else key
        if token in seen:
            return
        seen.add(token)
        components.append({"key": key, "label": label, "tail": tail, "source": source,
                           "order": len(components)})

    # 1) 目录域缓存 = 模板自己的章节清单（信息最完整）
    for paragraph in ctx.paragraphs:
        level = _toc_level(paragraph)
        if level is None:
            continue
        text = paragraph.text.strip()
        if not text:
            continue
        normalized = _normalize(text)
        match = _CHAPTER.match(normalized)
        if match and not _placeholder(match.group("tail")):
            add("chapter", text, match.group("tail"), "toc")
            continue
        key = _component_of(text)
        if key:
            add(key, text, source="toc")

    # 2) 一级 / 二级标题
    for paragraph in ctx.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        level = _style_level(paragraph)
        if level is None or level > 2:
            continue
        normalized = _normalize(text)
        match = _CHAPTER.match(normalized)
        if match and not _placeholder(match.group("tail")):
            add("chapter", text, match.group("tail"), "heading")
            continue
        key = _component_of(text)
        if key:
            add(key, text, source="heading")

    # 3) 组成部分标签（含未被标题样式标记的封面/声明等）
    for paragraph in ctx.paragraphs:
        text = paragraph.text.strip()
        if not text or len(text) > 40:
            continue
        key = _component_of(text)
        if key:
            add(key, text, source="label")

    if _has_cover_block(ctx) and "cover" not in seen:
        add("cover", "封面（模板前部字段区）", source="field")
    has_toc_field = any("TOC" in str(node) or "TOC" in (node.text or "")
                        for node in ctx.document.element.body.iter()
                        if node.tag in (qn("w:instrText"), qn("w:fldSimple")))
    if has_toc_field and "toc" not in seen:
        add("toc", "目录（TOC 域）", source="field")

    return {"outline": {"components": components, "has_cover_block": _has_cover_block(ctx)}}


def _paper_index(ctx) -> dict[str, Any]:
    """把论文里可用来匹配的标题/标签整理成索引。"""

    headings: list[tuple[int, str, str]] = []   # (段落号, 标准化全文, 去掉章节号后的尾文字)
    toc_chapters: list[tuple[int, str, str]] = []  # (段落号, 目录显示文字, 标题尾文字)
    labels: set[str] = set()
    positions: dict[str, int] = {}              # 组成部分 → 段落号（1 起）
    for index, paragraph in enumerate(ctx.paragraphs, 1):
        text = paragraph.text.strip()
        if not text:
            continue
        normalized = _normalize(text)
        toc_level = _toc_level(paragraph)
        level = _style_level(paragraph)
        if toc_level == 1:
            value = _PAGE_SUFFIX.sub("", text).strip()
            chapter = _CHAPTER.match(_normalize(value))
            numbered = _NUMBERED_TOC_HEADING.match(value)
            tail = chapter.group("tail") if chapter else (
                _normalize(numbered.group("tail")) if numbered else ""
            )
            if tail and _component_of(value) is None:
                toc_chapters.append((index, text, tail))
        if toc_level is None and level is not None and level <= 3:
            match = _CHAPTER.match(normalized)
            tail = match.group("tail") if match else normalized
            headings.append((index, normalized, tail))
        key = _component_of(text)
        if key and toc_level is None:
            labels.add(key)
            positions.setdefault(key, index)
    # 论文里的目录域
    has_toc_field = any("TOC" in (node.text or "") for node in ctx.document.element.body.iter()
                        if node.tag == qn("w:instrText"))
    front_parts = _front_matter_parts(ctx)
    labels.update(front_parts)
    return {"headings": headings, "toc_chapters": toc_chapters,
            "labels": labels, "positions": positions,
            "front_parts": front_parts, "toc_field": has_toc_field}


def _chapter_present(tail: str, index: dict[str, Any]) -> bool:
    """章节是否在论文中出现：按“章节号之后的标题文字”做包含匹配。"""

    tokens = [token for token in (tail, tail[:3], tail[:2]) if len(token) >= 2]
    for _, _, paper_tail in index["headings"]:
        if any(token in paper_tail or paper_tail in token for token in tokens):
            return True
    return False


def check_outline_errors(baseline: dict, ctx) -> list:
    """逐个比对模板清单与论文，缺失的组成部分各出一条结论。"""

    data = baseline.get("outline") or {}
    components = data.get("components") or []
    if not components:
        return []
    index = _paper_index(ctx)
    # 旧模块已从模板“必须包含”文字判定过的组成部分，这里不重复报
    already = set((baseline.get("structure") or {}).get("required") or {})
    declared = {"cover": "cover", "title_page": "title_page",
                "declaration": "originality", "authorization": "authorization",
                "abstract_zh": "abstract_zh", "abstract_en": "abstract_en",
                "toc": "toc", "references": "references",
                "acknowledgements": "acknowledgements", "appendix": "appendix",
                "achievements": "achievements"}

    findings = []
    reported_chapter_tails = set()
    for component in components:
        key = component["key"]
        if key in described_skip(already, declared):
            continue
        if key == "chapter":
            if _chapter_present(component["tail"], index):
                continue
            code = "007.1"
            reported_chapter_tails.add(component["tail"])
            # 章节用的是模板里那一章的标题文字，便于人工核对
            expected = f"模板列出的章节「{_display_label(component['label'])}」"
            actual = "论文中没有对应章节标题"
        else:
            code = _COMPONENT_CODES.get(key)
            if code is None:
                continue
            if key == "toc" and index["toc_field"]:
                continue
            if key in index["labels"]:
                continue
            # 封面/题名页：模板有字段区时，论文只要也有同样的字段区即算通过
            # （封面通常没有“封面”二字，不能只按标签文字判断）
            if key in {"cover", "title_page"}:
                if key in index.get("front_parts", set()):
                    continue
            expected = f"模板存在组成部分「{_display_label(component['label'])}」"
            actual = "论文中未识别到该组成部分"
        finding = make_finding(code, _title(code), expected, actual,
                               _missing_component_anchor(ctx, index, key))
        findings.append(finding)
    # 规范模板可能只给“第 X 章”占位样例，无法预先列出论文应有的具体章名。
    # 但论文自己的自动目录若仍列出某一级章，而正文已经不存在该标题，这是可确定的
    # 结构缺失。批注落在对应目录项，不回退到封面或文首。
    for paragraph_number, label, tail in index.get("toc_chapters", []):
        if tail in reported_chapter_tails or _chapter_present(tail, index):
            continue
        code = "007.1"
        findings.append(make_finding(
            code,
            _title(code),
            f"目录列出的章节「{_display_label(label)}」应在正文中存在对应一级标题",
            "正文中未识别到该章节标题",
            ctx.paragraphs[paragraph_number - 1],
        ))
    return findings


#: 论文组成部分的固定次序：缺失项的批注落在“它应处的位置”，不落在封面。
_COMPONENT_ORDER = ("cover", "title_page", "declaration", "authorization", "abstract_zh",
                    "abstract_en", "toc", "references", "appendix", "acknowledgements",
                    "achievements")


def _last_body_paragraph(ctx):
    for paragraph in reversed(ctx.paragraphs):
        if paragraph.text.strip():
            return paragraph
    return ctx.paragraphs[0] if ctx.paragraphs else None


def _missing_component_anchor(ctx, index: dict[str, Any], key: str):
    """缺失组成部分的批注锚点：优先落在它后面第一个已有部分（例如附录 → 致谢），
    其次落在它前面最后一个已有部分，最后落在正文末段。"""

    positions = index.get("positions") or {}
    if key in _COMPONENT_ORDER:
        offset = _COMPONENT_ORDER.index(key)
        following = [name for name in _COMPONENT_ORDER[offset + 1:] if name in positions]
        if following:
            return ctx.paragraphs[positions[following[0]] - 1]
        preceding = [name for name in reversed(_COMPONENT_ORDER[:offset]) if name in positions]
        if preceding:
            return ctx.paragraphs[positions[preceding[0]] - 1]
    return _last_body_paragraph(ctx)


def described_skip(already: set, declared: dict[str, str]) -> set[str]:
    """旧模块已判定必需性的组成部分集合（避免重复报）。"""

    return {key for key, alias in declared.items() if alias in already}
