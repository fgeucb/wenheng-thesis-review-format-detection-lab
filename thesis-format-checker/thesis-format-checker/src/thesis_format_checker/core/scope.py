from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Any

from thesis_format_checker.core.effective_format import clean_word_text

EXCLUDED_STYLE_TOKENS = (
    "标题",
    "heading",
    "title",
    "目录",
    "toc",
    "摘要",
    "abstract",
    "关键词",
    "keyword",
    "参考文献",
    "bibliography",
    "caption",
    "题注",
    "页眉",
    "页脚",
    "header",
    "footer",
)
PREFERRED_BODY_STYLES = ("正文", "normal", "body text", "bodytext", "普通")

#: 图表题注（图2-1/表3-1/Fig.2-1/Tab.3-1…）：属于题注，由题注规则负责比对，
#: 不能作为正文样例——题注常用五号字，当成正文就会报“正文字号应为小四”。
_CAPTION_LIKE = re.compile(r"^(?:图|表|figure\b|table\b|fig\.|tab\.)\s*\d", re.I)
#: 编号标题：``第X章``/``第X节`` 或多级编号 ``N.N``（编号本身就是结构证据）。
_NUMBERED_HEADING = re.compile(
    r"^(?:第\s*[0-9〇零一二三四五六七八九十百千两]{1,6}\s*[章节篇部分]|\d{1,3}(?:[.．]\d{1,3})+)")
#: 无编号前缀的“N 标题”：仅在文档使用“第X章”体系时才按标题处理（与标题树同口径）。
_BARE_NUMBER_HEADING = re.compile(r"^\d{1,3}(?![0-9])[ \t\u3000]+\S")
#: 标题的“像标题”判据：短、无句末标点、不以冒号收尾（引导句是正文）。
_HEADING_MAX_CHARS = 40
_SENTENCE_END = re.compile(r"[。；！？!?;]")


def _looks_like_numbered_heading(text: str, chapter_prefix: bool) -> bool:
    """段落是否像章节标题：用于把标题排除出正文样例（Word 侧正文范围）。"""

    stripped = text.strip()
    if not stripped:
        return False
    if _NUMBERED_HEADING.match(stripped):
        return True
    if not chapter_prefix or not _BARE_NUMBER_HEADING.match(stripped):
        return False
    return (len(stripped) <= _HEADING_MAX_CHARS and not _SENTENCE_END.search(stripped)
            and not stripped.endswith(("：", ":")))

#: 前置/后置部分的标志性标题：它们之间的内容才算正文（与 docx 侧口径一致）。
_FRONT_LABELS = ("摘要", "摘 要", "abstract", "目录", "contents", "封面", "题名页",
                 "原创性声明", "独创性声明", "授权使用声明", "使用授权声明")
_BACK_LABELS = ("参考文献", "references", "bibliography", "致谢", "acknowledg",
                "附录", "appendix", "攻读学位期间")


def _region_label(text: str) -> str | None:
    """标题属于前置(front)/后置(back)/正文(body)；返回 None 表示不像标志性标题。"""

    compact = re.sub(r"\s+", "", text or "").casefold()
    if not compact or len(compact) > 20:
        return None
    if any(compact.startswith(label.casefold().replace(" ", "")) for label in _FRONT_LABELS):
        return "front"
    if any(compact.startswith(label.casefold()) for label in _BACK_LABELS):
        return "back"
    return None


#: 前置区（封面/题名页/声明）结束的标志性标题：摘要或目录。
_FRONT_END_LABELS = ("摘要", "abstract", "目录", "contents")
#: 前置区证据：摘要之前出现这些字样，才说明确实存在封面/声明页需要单独处理。
#: 封面页通常**不出现“封面”二字**，而是学校名、毕业设计/毕业论文、题目、姓名、
#: 学号、学院、专业、指导教师、职称等要素，所以这里按“封面要素 + 声明页”取证。
_FRONT_EVIDENCE_LABELS = (
    # 声明/授权页
    "封面", "题名页", "声明", "授权",
    # 封面要素（中文）
    "毕业设计", "毕业论文", "学位论文", "题目", "题名", "姓名", "学号", "学院",
    "专业", "班级", "指导教师", "导师", "职称", "密级", "分类号", "udc",
    # 封面要素（英文封面常见写法，已去掉空格后比较）
    "bachelor", "thesis", "dissertation", "studentid", "supervisor", "college", "major",
)


def front_matter_end(texts) -> int:
    """返回“封面/声明等前置区”结束的段落下标（第一条摘要或目录标志）。

    该下标及之后的段落属于摘要/正文等常规部分；找不到标志时返回 0（不做裁剪）。
    前置区只应由封面模块按模板同位置比对，不参与标题层级、正文等通用判定。

    只有“摘要之前确实出现过封面/声明类内容”时才裁剪：否则（例如正文里恰好
    有一个叫“摘要”的小节）把前面全部当封面会误伤正文，此时一律返回 0。

    返回 0 同时表示“整篇都从第 1 段开始算”，调用方必须写成
    ``if 0 < front_end and index < front_end`` 才能避免把首段误伤。
    """

    evidence = False
    for index, text in enumerate(texts):
        compact = re.sub(r"[\s\u3000]+", "", str(text or "")).casefold()
        if not compact:
            continue
        if any(label in compact for label in _FRONT_EVIDENCE_LABELS):
            evidence = True
            continue
        if len(compact) > 10:
            continue
        if any(compact.startswith(label) for label in _FRONT_END_LABELS):
            return index if evidence else 0
    return 0


def front_end_of(ctx) -> int:
    """取文档上下文的前置区结束下标；0 表示没有摘要/目录标志，不做裁剪。

    优先复用 :class:`~thesis_format_checker.core.docx_context.DocxContext`
    上的缓存属性（避免每次调用都重新扫描段落），普通对象则退回文本扫描。
    """

    value = getattr(ctx, "front_matter_end", None)
    if isinstance(value, int):
        return value
    return front_matter_end([getattr(p, "text", "") or "" for p in getattr(ctx, "paragraphs", [])])


@dataclass(frozen=True)
class ParagraphInfo:
    index: int
    paragraph: Any
    text: str
    style_name: str


def style_name_of(paragraph) -> str:
    try:
        style = paragraph.Range.Style
        for attr in ("NameLocal", "Name"):
            value = getattr(style, attr, None)
            if value:
                return str(value)
        return str(style)
    except Exception:
        return ""


def style_is_excluded(name: str) -> bool:
    lowered = name.casefold().replace(" ", "")
    return any(token.casefold().replace(" ", "") in lowered for token in EXCLUDED_STYLE_TOKENS)


def _mapping_side(mapping: dict[str, Any] | None, side: str) -> dict[str, Any]:
    if not mapping:
        return {}
    value = mapping.get(side, {})
    return value if isinstance(value, dict) else {}


def collect_body_paragraphs(document, mapping: dict[str, Any] | None, side: str,
                            *, facts=None) -> list[ParagraphInfo]:
    """采集正文段落。

    ``facts``：可选的 :class:`~thesis_format_checker.core.word_facts.WordFacts`
    事实表。传入时所有段落属性（文本/样式/大纲级别/是否在表格）都从事实表读取，
    同一份文档只读一遍；不传时保持旧的逐段读取（单测与外部调用）。
    两种路径读到的是同一个 COM 接口的同一个值，因此结果一致。
    """

    settings = _mapping_side(mapping, side)
    configured_styles = {str(v).casefold() for v in settings.get("body_styles", [])}
    start_marker = settings.get("start_text")
    end_marker = settings.get("end_text")
    start_index = settings.get("start_paragraph")
    end_index = settings.get("end_paragraph")
    explicit_range = bool(start_marker or end_marker or start_index or end_index)

    def collect(skip_outer_regions: bool, drop_front_matter: bool = True
                ) -> tuple[list[ParagraphInfo], bool]:
        """采集正文段；``skip_outer_regions`` 时跳过封面/声明（前置）与后置部分。

        ``drop_front_matter`` 时再按“摘要/目录标志”裁掉封面、题名页、声明页：
        这部分只由封面模块按模板同位置比对，不进入正文格式基准。
        """

        items: list[ParagraphInfo] = []
        texts: list[str] = []
        active = start_marker is None
        region: str | None = None
        saw_body_heading = False
        total = facts.count if facts is not None else int(document.Paragraphs.Count)
        for index in range(1, total + 1):
            if facts is not None:
                paragraph = facts.paragraph(index)
                text = facts.clean_text(index)
            else:
                paragraph = document.Paragraphs(index)
                text = clean_word_text(paragraph.Range.Text)
            texts.append(text)
            if start_marker and str(start_marker) in text:
                active = True
                continue
            if end_marker and str(end_marker) in text:
                break
            if not active or not text.strip():
                continue
            if start_index and index < int(start_index):
                continue
            if end_index and index > int(end_index):
                continue
            if facts is not None:
                in_table = facts.in_table(index)
                outline = facts.outline_level(index)
            else:
                try:
                    in_table = bool(paragraph.Range.Information(12))  # wdWithInTable
                except Exception:
                    in_table = False
                try:
                    outline = int(paragraph.OutlineLevel)
                except Exception:
                    outline = None
            if not in_table and outline is not None and outline != 10 and text.strip():
                label = _region_label(text)
                if label == "back":
                    break
                if label == "front":
                    region = "front"
                else:
                    region = "body"
                    saw_body_heading = True
                continue
            if in_table or (outline is not None and outline != 10):
                continue
            if skip_outer_regions and region in (None, "front"):
                continue  # 封面、声明、摘要等前置内容不作为正文样例
            style_name = facts.style_name(index) if facts is not None else style_name_of(paragraph)
            if configured_styles:
                if style_name.casefold() not in configured_styles:
                    continue
            elif style_is_excluded(style_name):
                continue
            items.append(ParagraphInfo(index, paragraph, text, style_name))
        if drop_front_matter:
            front_end = front_matter_end(texts)
            if 0 < front_end:
                items = [item for item in items if item.index > front_end]
        # 题注与章节标题不属于正文范围：题注按题注规则比对、标题按标题规则比对，
        # 否则“图2-2 模型整体流程图”（五号）会被当成正文并报“正文字号应为小四”。
        chapter_prefix = any(_NUMBERED_HEADING.match(str(text).strip())
                             and str(text).strip().startswith("第") for text in texts)
        items = [item for item in items
                 if not _CAPTION_LIKE.match(item.text.strip())
                 and not _looks_like_numbered_heading(item.text, chapter_prefix)]
        return items, saw_body_heading

    if explicit_range:
        # 显式区间由调用方通过映射文件指定：此时不按前置区裁剪，避免越权改变配置含义。
        candidates, _ = collect(False, drop_front_matter=False)
    else:
        candidates, saw_body_heading = collect(True)
        if not candidates and not saw_body_heading:
            candidates, _ = collect(False)  # 文档没有标题结构时按原逻辑采集

    if configured_styles or not candidates:
        return candidates

    # 无映射时，优先选择 Word 的普通正文样式；否则采用总正文字符数最大的非标题样式。
    preferred = [
        p
        for p in candidates
        if any(token in p.style_name.casefold() for token in PREFERRED_BODY_STYLES)
    ]
    if preferred:
        return preferred
    lengths: dict[str, int] = defaultdict(int)
    for item in candidates:
        lengths[item.style_name] += len(item.text)
    chosen_style = max(lengths, key=lengths.get)
    return [p for p in candidates if p.style_name == chosen_style]
