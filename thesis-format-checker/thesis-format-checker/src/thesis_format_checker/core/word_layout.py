"""Word 排版快照：真实分页、每页所属节与页眉页脚变体、版心首末正文字符位置。

用途：页眉页脚问题的批注落点。Word 不允许批注写进页眉页脚；
本模块读取本页正文真实首末边界，供写入零长度批注。只读、不保存输入。

要点：

* XML 段落号与 Word 段落号会因文本框等段落产生偏移，用“文字完全一致的唯一段落”
  做映射；重复文字须由两侧唯一段落及完全相同的段落序列夹定；
* 页眉页脚变体用 Word 的首页/奇偶页真实归属读取（SeekView + HeaderFooter.Index），
  不按页码奇偶猜测；
* 只有“可安全定位”的段落（无域、无修订、无嵌入对象）才作为批注落点。
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from thesis_format_checker.core.point_comments import safe_paragraph
from thesis_format_checker.core.word_session import WordSession


@dataclass(frozen=True)
class BodyPoint:
    """版心内的一个可用批注落点（正文字符）。"""

    paragraph: int      # XML（python-docx）段落号，0 起
    offset: int         # 段内字符偏移
    position: int       # Word 文档字符位置
    y: float | None = None
    edge: str | None = None


@dataclass
class Page:
    number: int                                  # 文档实际页序，1 起（不采用可能重启的印刷页码）
    section: int | None
    header_variant: str | None = None            # default / first / even
    footer_variant: str | None = None
    top: BodyPoint | None = None
    bottom: BodyPoint | None = None
    reason: str = ""
    chapter_paragraph: int | None = None
    adjusted_number: int | None = None


@dataclass
class Pagination:
    pages: list[Page] = field(default_factory=list)
    abstract_page: int | None = None
    notes: list[str] = field(default_factory=list)
    boundaries_verified: bool = False

    def page_of_section(self, section_index: int) -> list[Page]:
        """第 ``section_index`` 节（0 起）的所有页。"""

        return [page for page in self.pages if page.section == section_index]


def normalized(text: str) -> str:
    return text.replace("\x0b", "\n").replace("\x0c", "\n").replace("\r", "\n")


def paragraph_matches(xml_texts, word_texts) -> dict[int, int]:
    """Word 段落号 → XML 段落号；唯一文字和严格夹定的相同序列可匹配。"""

    left = [normalized(text) for text in xml_texts]
    right = [normalized(text) for text in word_texts]
    if left == right:
        return {index: index for index in range(len(left))}
    counts_left, counts_right = Counter(left), Counter(right)
    unique = {text: index for index, text in enumerate(left)
              if text.strip() and counts_left[text] == counts_right[text] == 1}
    matches = {index: unique[text] for index, text in enumerate(right) if text in unique}
    # 重复空段只能在两侧唯一段落已对齐、整段序列逐项相同时映射。
    # 不使用“最近的空段”或比例推算，保留 Word/XML 多段偏移时的安全性。
    boundaries = [(-1, -1), *sorted(matches.items()), (len(right), len(left))]
    for (wa, xa), (wb, xb) in zip(boundaries, boundaries[1:]):
        if wb > wa and xb > xa and right[wa + 1:wb] == left[xa + 1:xb]:
            matches.update({wi: xa + wi - wa for wi in range(wa + 1, wb)})
    return matches


def utf16_offsets(text: str) -> list[int]:
    """字符下标 → UTF-16 偏移（Word 的字符位置按 UTF-16 计）。"""

    result = [0]
    for char in text:
        result.append(result[-1] + len(char.encode("utf-16-le")) // 2)
    return result


def select_body_points(fragments, page_height: float):
    """从本页各段的首末字符里选出“最靠上/最靠下”的两个落点。"""

    if not fragments:
        return None, None
    tops = [pair[0] for pair in fragments]
    bottoms = [pair[1] for pair in fragments]
    top = (min(tops, key=lambda point: (point.y, point.position))
           if all(point.y is not None for point in tops)
           else min(tops, key=lambda point: point.position))
    bottom = (max(bottoms, key=lambda point: (point.y, point.position))
              if all(point.y is not None for point in bottoms)
              else max(bottoms, key=lambda point: point.position))
    return top, bottom


def select_safe_page_boundaries(on_page, boundary, records, start: int, stop: int):
    """Return the nearest safe top/bottom points without leaving the page."""

    from bisect import bisect_left

    if not on_page:
        return None, None
    top = next((point for index in on_page
                if (point := boundary(index, False)) is not None), None)
    bottom_candidates = list(reversed(on_page))
    first, _last, text = records[bottom_candidates[0]]
    fragment_end = bisect_left(utf16_offsets(text), stop - first)
    if (first >= start and text[:fragment_end] == "\f"
            and len(bottom_candidates) > 1):
        bottom_candidates = bottom_candidates[1:]
    bottom = next((point for index in bottom_candidates
                   if (point := boundary(index, True)) is not None), None)
    return top, bottom


def _page_at(document, position: int) -> int:
    end = int(document.Content.End)
    return int(document.Range(position, min(position + 1, end)).Information(3))


def _page_start(document, page_number: int, end: int) -> int:
    """该页的起始字符位置：GoTo 结果要核对物理页序，编号重启时用二分定位。"""

    try:
        start = int(document.GoTo(What=1, Which=1, Count=page_number).Start)
        if _page_at(document, start) == page_number:
            return start
    except Exception:
        pass
    low, high = 0, max(0, end - 1)
    while low < high:
        middle = (low + high) // 2
        if _page_at(document, middle) < page_number:
            low = middle + 1
        else:
            high = middle
    return low


def _variant(window, document, position: int, kind: str) -> str | None:
    """读取该字符位置所在页实际使用的页眉/页脚变体（首页/奇偶页由 Word 判定）。"""

    try:
        window.View.SeekView = 0
        document.Range(position, position).Select()
        window.View.SeekView = 9 if kind == "header" else 10
        return {1: "default", 2: "first", 3: "even"}.get(
            int(window.Selection.HeaderFooter.Index))
    except Exception:
        return None
    finally:
        window.View.SeekView = 0


def read_pagination(source, ctx) -> Pagination:
    """读取按页的排版事实；``ctx`` 为待检测文档的 DocxContext。"""

    # 与排版几何路径共用真实页边界读取，不能复用“最先找到的安全正文段落”。
    with WordSession() as session:
        document = session.open_readonly(Path(source).resolve())
        session.resume_pagination()
        document.Windows(1).View.Type = 3
        document.Windows(1).View.ShowRevisionsAndComments = False
        records = _paragraph_records(document)
        matches = paragraph_matches([p.text for p in ctx.paragraphs],
                                    [record[2] for record in records])
        end = int(document.Content.End)
        count = int(document.ComputeStatistics(2))
        starts = [_page_start(document, n, end) for n in range(1, count + 1)] + [end]
        return _read_boundary_pagination(document, ctx, records, matches, starts)


def _read_boundary_pagination(document, ctx, records, matches, starts) -> Pagination:
    """取本页最靠近页顶/页底的安全正文边界。

    页末长段落按 UTF-16 偏移截到本页，页末段落标记或末字符前的空范围不会跨页。
    若页首或页末段落含图片、域等复杂对象，只在**同一页**继续寻找最近的安全段落；
    不跨页、不退回文首，也不改选封面或其他页面的无关正文。整页都没有安全段落时，
    边界保持为空，由写入层记为仅报告。
    """
    from bisect import bisect_left
    from thesis_format_checker.core.point_comments import boundary_text

    snapshot = Pagination(boundaries_verified=True)
    matches = paragraph_matches([boundary_text(p) for p in ctx.paragraphs],
                                [r[2] for r in records])
    window = document.Windows(1)
    for number, (start, stop) in enumerate(zip(starts, starts[1:]), 1):
        if start >= stop or _page_at(document, start) != number:
            snapshot.pages.append(Page(number, None, reason="本页无可定位的正文范围"))
            continue
        ran = document.Range(start, max(start, stop - 1))
        if int(ran.Sections.Count) != 1:
            snapshot.pages.append(Page(number, None, reason="本页跨节，无法唯一确认页眉页脚归属"))
            continue
        page = Page(number, int(ran.Information(2)) - 1)
        page.adjusted_number = int(document.Range(start, min(start + 1, stop)).Information(1))
        page.header_variant = _variant(window, document, start, "header")
        page.footer_variant = _variant(window, document, start, "footer")
        on_page = [i for i, (a, b, _) in enumerate(records) if a < stop and b > start]

        def boundary(index, bottom):
            xml_index = matches.get(index)
            if xml_index is None:
                return None
            paragraph = ctx.paragraphs[xml_index]
            first, last, text = records[index]
            offsets = utf16_offsets(text)
            if int(document.Range(first, last).Font.Hidden) != 0:
                return None
            if not safe_paragraph(paragraph):
                from thesis_format_checker.core.point_comments import safe_field_edge
                position = last - 1 if bottom else first
                if (safe_field_edge(paragraph) and start <= position < stop
                        and _page_at(document, position) == number):
                    return BodyPoint(xml_index, len(paragraph.text) if bottom else 0,
                                     position, edge="end" if bottom else "start")
                return None
            if last - first not in {offsets[-1] + 1, offsets[-1] + 2}:
                return None
            if bottom:
                # end-of-paragraph is safe; a soft page boundary belongs to the NEXT page,
                # so use the zero-width point immediately before this page's last character.
                offset = min(len(text), bisect_left(offsets, stop - first) - 1)
            else:
                offset = bisect_left(offsets, max(0, start - first))
            if not 0 <= offset <= len(text):
                return None
            position = first + offsets[offset]
            if not start <= position < stop or _page_at(document, position) != number:
                return None
            point_range = document.Range(position, position)
            try:
                window.ScrollIntoView(point_range, True)
                y = _number(point_range.Information(6))
                if y is not None and y < 0:
                    y = None
            except Exception:
                y = None
            return BodyPoint(xml_index, offset, position, y)

        if on_page:
            page.top, page.bottom = select_safe_page_boundaries(
                on_page, boundary, records, start, stop)
        if page.top is None or page.bottom is None:
            page.reason = "实际页边界含复杂对象、映射不唯一或不可读；不改选无关正文"
        snapshot.pages.append(page)
    return snapshot


# ---------------------------------------------------------------------------
# 排版几何：段落框、图形框、表格分页（供 078 题注位置、087 跨页表格、033/098 越界使用）
# ---------------------------------------------------------------------------

def _interesting_indices(ctx) -> set[int]:
    """需要几何坐标的段落：图表题注、图片、公式、标题、表格与目录条目。"""

    wanted: set[int] = set()
    try:
        from thesis_format_checker.detectors.figures import collect_caption_records

        wanted |= {ctx.index(record["paragraph"]) for record in collect_caption_records(ctx)}
    except Exception:
        pass
    try:
        from thesis_format_checker.detectors.equations import _read_records

        records, _ = _read_records(ctx)
        wanted |= {ctx.index(record["paragraph"]) for record in records}
    except Exception:
        pass
    try:
        from thesis_format_checker.detectors.headings import collect_heading_records

        wanted |= {ctx.index(record["paragraph"]) for record in collect_heading_records(ctx)}
    except Exception:
        pass
    try:
        from thesis_format_checker.detectors.toc import _toc_level

        wanted |= {index for index, paragraph in enumerate(ctx.paragraphs)
                   if _toc_level(paragraph) is not None}
    except Exception:
        pass
    for index, paragraph in enumerate(ctx.paragraphs):
        if paragraph._p.xpath(".//w:drawing|.//w:pict|.//w:object|.//w:tbl"):
            wanted.add(index)
    return {index for index in wanted if 0 <= index < len(ctx.paragraphs)}

def _number(value):
    """Word 对“混合/未定义”返回 9999999，视为不可读。"""

    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    return None if number >= 9999999 else number


def _point(document, position: int):
    """字符位置 → ``{page, x, y}``；坐标不可读时返回 None。"""

    end = int(document.Content.End)
    if position < 0 or position > end:
        return None
    ran = document.Range(position, min(position + 1, end))
    page = int(ran.Information(3))
    x, y = _number(ran.Information(5)), _number(ran.Information(6))
    if x is None or y is None or x < 0 or y < 0:
        return None
    return {"page": page, "x": x, "y": y}


def range_box(document, ran, height: float = 0.0):
    """字符范围 → 矩形：``{page, end_page, x, y, bottom}``。"""

    first = _point(document, int(ran.Start))
    last = _point(document, max(int(ran.Start), int(ran.End) - 1))
    if first is None or last is None:
        return None
    return {"page": first["page"], "end_page": last["page"], "x": first["x"],
            "right": last["x"], "y": first["y"],
            "bottom": last["y"] + max(0.0, float(height or 0.0))}


def shape_box(document, shape):
    """浮动图形的矩形；旋转或相对定位无法换算时返回 None（不猜位置）。"""

    anchor = _point(document, int(shape.Anchor.Start))
    if anchor is None or abs(_number(shape.Rotation) or 0.0) > 0.01:
        return None
    top, left = _number(shape.Top), _number(shape.Left)
    if top is None or left is None:
        return None
    vertical = int(shape.RelativeVerticalPosition)
    horizontal = int(shape.RelativeHorizontalPosition)
    setup = shape.Anchor.Sections(1).PageSetup
    if vertical == 0:
        top += float(setup.TopMargin)
    elif vertical in (2, 3):
        top += anchor["y"]
    elif vertical != 1:
        return None
    if horizontal == 0:
        left += float(setup.LeftMargin)
    elif horizontal == 2:
        if int(setup.TextColumns.Count) != 1:   # 分栏时列起点不可知
            return None
        left += float(setup.LeftMargin)
    elif horizontal == 3:
        left += anchor["x"]
    elif horizontal != 1:
        return None
    height, width = _number(shape.Height) or 0.0, _number(shape.Width) or 0.0
    if top < 0 or top + height > float(setup.PageHeight):
        return None
    return {"page": anchor["page"], "end_page": anchor["page"], "x": left, "y": top,
            "bottom": top + height, "width": width, "height": height}


def _paragraph_format(word_paragraph) -> dict:
    """段落的有效格式（Word 解析）：对齐、行距、段间距、字符单位缩进。"""

    fmt = word_paragraph.Format
    result = {}
    for key, name in (("alignment", "Alignment"), ("space_before", "SpaceBefore"),
                      ("space_after", "SpaceAfter"), ("line_spacing", "LineSpacing"),
                      ("line_spacing_rule", "LineSpacingRule")):
        result[key] = _number(getattr(fmt, name))
    for key, points, chars in (("first_line_indent", "FirstLineIndent", "CharacterUnitFirstLineIndent"),
                               ("left_indent", "LeftIndent", "CharacterUnitLeftIndent"),
                               ("right_indent", "RightIndent", "CharacterUnitRightIndent")):
        value, char_value = _number(getattr(fmt, points)), _number(getattr(fmt, chars))
        if char_value is not None and (char_value != 0 or value == 0):
            result[key] = ("字符", char_value)
        elif value is not None:
            result[key] = ("磅", value)
    return result


def _paragraph_records(document, facts=None) -> list[tuple[int, int, str]]:
    """每段的 ``(起点, 终点, 不含段落标记的文字)``。

    给出事实表（旧版 Word 流程里已经读过每段文字与位置）时直接复用缓存；
    没有事实表或位置读取失败时退回逐段读取，保证与旧实现逐项一致。
    """

    if facts is not None:
        records: list[tuple[int, int, str]] = []
        for index in facts.indices():
            start = facts.start(index)
            text = facts.text(index)
            if start is None:
                records = []
                break
            records.append((start, start + len(text.encode("utf-16-le")) // 2, text.rstrip("\r\x07")))
        else:
            if records:
                return records
    records = []
    for index in range(1, document.Paragraphs.Count + 1):
        paragraph = document.Paragraphs(index)
        ran = paragraph.Range
        records.append((int(ran.Start), int(ran.End), ran.Text.rstrip("\r\x07")))
    return records


def read_layout(source, ctx, *, paragraph_indices=None, want_figures: bool = True,
                want_tables: bool = True) -> dict:
    """读取排版几何（按 **XML 段落号** 索引），供题注位置、跨页表格、越界检测使用。

    本函数会**自己打开一次 Word**；如果调用方已经打开同一份文档（旧版 Word 流程
    合并遍历的场景），请改用 :func:`read_layout_in_session`，省掉一次打开与重排。

    返回::

        {
          "paragraphs": {段落号: {"box": {...}, "format": {...}}},
          "figures":    {段落号: {"boxes": [...]}},   # 该段内图片/浮动图形矩形
          "tables":     {宿主段落号: {...}},           # 表格分页、标题行、手工分页、浮动
          "notes":      [...],                        # 读取受限的说明
        }

    细节见 :func:`read_layout_in_session`。
    """

    with WordSession() as session:
        document = session.open_readonly(Path(source).resolve())
        return read_layout_in_session(document, ctx, paragraph_indices=paragraph_indices,
                                      want_figures=want_figures, want_tables=want_tables)


def read_layout_in_session(document, ctx, *, paragraph_indices=None, want_figures: bool = True,
                           want_tables: bool = True, facts=None) -> dict:
    """在**已经打开的 Word 文档**上读取排版几何（不负责开关 Word）。

    与同门一致：坐标读不到就返回 None/不记录，绝不用存储顺序冒充视觉位置。

    性能：页码/节号/打印页码**按页查询**（每页 3 次 COM），段落的页归属用字符位置
    在页面边界里二分换算；几何只对“需要的段落”（题注/图片/公式/表格/标题/目录条目
    与每页首末正文）读取，不再对整篇 1000+ 段逐段读 `Information()`。

    ``facts``：可选的 :class:`~thesis_format_checker.core.word_facts.WordFacts`。
    给出时"每段文字 + 起止位置"直接复用事实表缓存（旧版 Word 流程里已经读过一遍），
    不再对整篇一千多个段落重复读 `Range.Text/Start/End`（实测省下数千次 COM 调用）。
    """

    from thesis_format_checker.core.xmlutil import qn

    layout = {"paragraphs": {}, "figures": {}, "tables": {}, "notes": []}
    interesting = _interesting_indices(ctx) if paragraph_indices is None else set(paragraph_indices)
    document.Windows(1).View.Type = 3
    document.Windows(1).View.ShowRevisionsAndComments = False
    document.Repaginate()
    records = _paragraph_records(document, facts)
    matches = paragraph_matches([paragraph.text for paragraph in ctx.paragraphs],
                                [record[2] for record in records])
    word_paragraph = ((lambda index: facts.paragraph(index)) if facts is not None
                      else (lambda index: document.Paragraphs(index)))
    word_paragraphs = {xml_index: word_paragraph(word_index + 1)
                       for word_index, xml_index in matches.items()}
    # 页边界 + 每页的节号/打印页码：COM 调用数与“页数”成正比，而不是“段数”。
    end = int(document.Content.End)
    count = int(document.ComputeStatistics(2))
    starts = [_page_start(document, number, end) for number in range(1, count + 1)] + [end]
    page_info: dict[int, dict] = {}
    for number in range(1, count + 1):
        ran = document.Range(starts[number - 1], min(starts[number - 1] + 1, end))
        printed = _number(ran.Information(1))
        page_info[number] = {"section": int(ran.Information(2)) - 1,
                             "printed": int(printed) if printed else None}
    layout["page_info"] = page_info

    def page_of(position: int) -> int | None:
        if position < 0 or position >= end:
            return None
        import bisect

        return max(1, bisect.bisect_right(starts, position))

    # 只对“需要的段落”读几何：题注/图片/公式/表格/标题/目录条目。
    for xml_index in sorted(interesting):
        paragraph = word_paragraphs.get(xml_index)
        if paragraph is None:
            continue
        item = {}
        try:
            item["format"] = _paragraph_format(paragraph)
            start = int(paragraph.Range.Start)
            page = page_of(start)
            item["page"] = page
            if page:
                item["section"] = page_info[page]["section"]
                item["printed_number"] = page_info[page]["printed"]
            height = (item["format"].get("line_spacing")
                      or _number(paragraph.Range.Font.Size) or 12.0)
            item["box"] = range_box(document, paragraph.Range, height)
        except Exception as exc:
            item["error"] = type(exc).__name__ + ": " + str(exc)
        layout["paragraphs"][xml_index] = item
    if want_figures:
        for xml_index, paragraph in word_paragraphs.items():
            element = ctx.paragraphs[xml_index]._p
            if not element.xpath(".//w:drawing|.//w:pict|.//w:object"):
                continue
            try:
                boxes = []
                for inline in paragraph.Range.InlineShapes:
                    box = range_box(document, inline.Range, float(inline.Height))
                    if box is not None:
                        box["width"] = float(inline.Width)
                        box["height"] = float(inline.Height)
                    boxes.append(box)
                for shape in document.Shapes:
                    if int(paragraph.Range.Start) <= int(shape.Anchor.Start) < int(paragraph.Range.End):
                        boxes.append(shape_box(document, shape))
                layout["figures"][xml_index] = {"boxes": boxes}
            except Exception as exc:
                layout["notes"].append(
                    f"第 {xml_index + 1} 段图形坐标读取失败：{type(exc).__name__}")
    if want_tables:
        table_paragraphs = []
        for xml_index, paragraph in enumerate(ctx.paragraphs):
            element = paragraph._p.getparent()
            has_table = element is not None and any(
                node.tag == qn("w:tbl") for node in element)
            if not has_table:
                continue      # 只有紧邻表格的段落才需要问 Word，避免整篇扫描
            try:
                count = int(word_paragraphs[xml_index].Range.Tables.Count) if xml_index in word_paragraphs else 0
            except Exception:
                count = 0
            if count == 1:
                table_paragraphs.append(xml_index)
        for xml_index in table_paragraphs:
            paragraph = word_paragraphs[xml_index]
            try:
                table = paragraph.Range.Tables(1)
                start, end = int(table.Range.Start), int(table.Range.End) - 1
                pages = (int(document.Range(start, start + 1).Information(3)),
                         int(document.Range(end, end + 1).Information(3)))
                rows = [None if _number(row.HeadingFormat) is None
                        else bool(_number(row.HeadingFormat)) for row in table.Rows]
                element = next((node for node in ctx.paragraphs[xml_index]._p.getparent()
                                if node.tag == qn("w:tbl")), None)
                layout["tables"][xml_index] = {
                    "pages": pages, "headers": rows,
                    "floating": bool(element is not None
                                     and element.xpath("./w:tblPr/w:tblpPr")),
                    "box": range_box(document, table.Range, 0.0),
                }
            except Exception as exc:
                layout["notes"].append(
                    f"第 {xml_index + 1} 段表格分页读取失败：{type(exc).__name__}")
    # 顺带产出“按页锚点”：页眉/页脚批注需要的版心首末正文字符，
    # 这样写批注时不必再单独打开一次 Word（省一次 3–4 分钟的排版遍历）。
    layout["pagination"] = _read_boundary_pagination(document, ctx, records, matches, starts)
    return layout


def _pagination_from_layout(ctx, layout: dict) -> Pagination:
    """从逐段排版事实里组装按页锚点（版心顶部/底部正文字符）。"""

    from thesis_format_checker.core.scope import front_matter_end

    candidates: dict[int, list[BodyPoint]] = {}
    for page, points in (layout.get("page_points") or {}).items():
        candidates[page] = list(points)
    info = layout.get("page_info") or {}
    snapshot = Pagination()
    for page_number in sorted(candidates):
        points = candidates[page_number]
        top = (min(points, key=lambda point: (point.y, point.paragraph))
               if all(point.y is not None for point in points)
               else min(points, key=lambda point: point.paragraph))
        bottom = (max(points, key=lambda point: (point.y, point.paragraph))
                  if all(point.y is not None for point in points)
                  else max(points, key=lambda point: point.paragraph))
        page = info.get(page_number) or {}
        snapshot.pages.append(Page(page_number, page.get("section"), top=top,
                                   bottom=bottom, adjusted_number=page.get("printed")))
    front_end = front_matter_end(ctx.texts)
    if front_end:
        for page in snapshot.pages:
            if any(point.paragraph >= front_end for point in (page.top, page.bottom) if point):
                snapshot.abstract_page = page.number
                break
    return snapshot
