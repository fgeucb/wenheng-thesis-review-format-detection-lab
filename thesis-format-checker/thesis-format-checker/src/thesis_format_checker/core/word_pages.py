"""Word 分页事实：段落所在页、节页范围与总页数（只读，不写回文档）。

有些规则（标题与后续内容是否同页、章节是否另起页、页码缺失/重复/跳号）在
OOXML 里无法判定，只能由 Word 排版后给出。本模块把排版结果整理成事实，
供后续规则使用；调用方必须已经打开文档，且从不保存该文档。

性能（优化 Tier 0）
==================

旧实现是**逐段**读 ``Range.Information(3)`` 与 ``Information(2)``：样例论文约
1000 段 → 3000+ 次 COM 调用，而且每次都要 Word 重新解释版面。现在流程内部使用
:class:`PageFacts`：

* **页边界按页读**：``GoTo(wdGoToPage, wdGoToAbsolute, n)`` 每页一次；
* **段落页码/节号按需读**：只有真正会被问到的段落（标题、标题后的首段、目录条目…）
  才读，第一次读到后缓存；
* **空白页按页判定**：看该页范围内有没有文字或图形，比旧实现"该页有没有段落起始"
  更准确（跨页的长段落不会再把整页误判成空白页）。

所有值仍然用与旧实现相同的 COM 接口读取，只是不再重复读，因此结论一致。
:func:`read_page_facts` 保留为**参考实现**（逐段读取），用于对照校验与单测。
"""
from __future__ import annotations

#: ``Range.Information`` 取值（与 ``core/word_facts.py`` 保持一致）。
WD_ACTIVE_END_SECTION_NUMBER = 2
WD_ACTIVE_END_PAGE_NUMBER = 3

#: 判"这一页有没有内容"时要忽略的排版标记：段落标记、单元格标记、图形占位符等。
_PAGE_MARKERS = "\r\x07\x01\x02\x0b\x0c \t\u3000\u00a0"


def read_page_facts(document) -> dict:
    """逐段读取分页事实（**参考实现**）：``{total_pages, pages, sections, section_of}``。

    与优化前的实现逐项一致，用于单测与 ``tools/verify_word_facts.py`` 的对照；
    正式流程请用 :class:`PageFacts`（按页读取 + 按需查询）。

    说明：``Document.Repaginate()`` 只影响内存中的分页视图，检测流程不会保存文档，
    因此不会改动用户的文件；页码取“物理页”，显示用的编号由各节设置换算。
    """

    try:
        document.Repaginate()
    except Exception:
        pass
    pages: dict[int, int | None] = {}
    start_pages: dict[int, int | None] = {}
    section_of: dict[int, int | None] = {}
    total = None
    try:
        total = int(document.ComputeStatistics(2))  # wdStatisticPages
    except Exception:
        total = None
    try:
        count = int(document.Paragraphs.Count)
    except Exception:
        count = 0
    for index in range(1, count + 1):
        try:
            range_ = document.Paragraphs(index).Range
            pages[index] = int(range_.Information(WD_ACTIVE_END_PAGE_NUMBER))
            try:
                start = int(range_.Start)
                end = int(document.Content.End)
                start_pages[index] = int(
                    document.Range(start, min(start + 1, end))
                    .Information(WD_ACTIVE_END_PAGE_NUMBER))
            except Exception:
                start_pages[index] = pages[index]
            section_of[index] = int(range_.Information(WD_ACTIVE_END_SECTION_NUMBER))
        except Exception:
            pages[index] = None
            start_pages[index] = None
            section_of[index] = None
    sections = []
    try:
        for position in range(1, int(document.Sections.Count) + 1):
            numbers = [pages[index] for index, value in section_of.items()
                       if value == position and pages.get(index)]
            sections.append({"index": position,
                             "first_page": min(numbers) if numbers else None,
                             "last_page": max(numbers) if numbers else None})
    except Exception:
        sections = []
    if total is None:
        known = [page for page in pages.values() if page]
        total = max(known) if known else None
    return {"total_pages": total, "pages": pages, "start_pages": start_pages,
            "sections": sections,
            "section_of": section_of}


class PageFacts:
    """按页读取、按需查询的分页事实。

    实现要点（都是实测踩出来的）：

    * Word 的 ``GoTo(wdGoToPage, n)`` **定位不到完全空白的页**（返回 0），因此页边界
      只对"能定位的页"建立，空白页通过"下一张能定位的页之前的那个字符"探针确认；
    * 段落页码/节号仍然用 ``Information(3)/(2)`` 精确读取（与旧实现同一个接口），
      只是第一次被问到才读；
    * 字符位置 → 页码同样用 ``Information(3)``（页边界二分在幻影页上会算错），
      结果缓存，供目录条目这类少量定位使用。
    """

    def __init__(self, document, facts=None) -> None:
        #: ``facts``：可选的 :class:`~thesis_format_checker.core.word_facts.WordFacts`
        #: （复用其缓存的段落/范围对象）。
        self.document = document
        self.facts = facts
        self.calls = 0
        self._total: int | None = None
        self._starts: list[int | None] | None = None
        self._valid: list[tuple[int, int]] | None = None
        self._pages: dict[int, int | None] = {}
        self._start_pages: dict[int, int | None] = {}
        self._section_of: dict[int, int | None] = {}
        self._page_at: dict[int, int | None] = {}
        self._sections: list[dict] | None = None
        self._content: list[int] | None = None

    def _bump(self) -> None:
        self.calls += 1

    def _range(self, index: int):
        if self.facts is not None:
            return self.facts.range(index)
        self._bump()
        return self.document.Paragraphs(index).Range

    # -- 页边界 ----------------------------------------------------------

    def _ensure_page_starts(self) -> list[int | None]:
        """每页的首页字符位置；Word 定位不到的页记 ``None``。"""

        if self._starts is None:
            try:
                self.document.Repaginate()
            except Exception:
                pass
            starts: list[int | None] = []
            try:
                total = self.total_pages or 0
            except Exception:
                total = 0
            for number in range(1, int(total) + 1):
                self._bump()
                try:
                    value = int(self.document.GoTo(1, 1, number).Start)  # wdGoToPage
                except Exception:
                    value = None
                starts.append(value)
            self._starts = starts
        return self._starts

    def _valid_pages(self) -> list[tuple[int, int]]:
        """``[(页码, 可定位的首页字符位置)]``，按位置严格递增。"""

        if self._valid is None:
            valid: list[tuple[int, int]] = []
            last = -1
            for number, start in enumerate(self._ensure_page_starts() or [], start=1):
                if start is None or start <= last:
                    continue
                valid.append((number, start))
                last = start
            self._valid = valid
        return self._valid

    @property
    def total_pages(self) -> int | None:
        """总页数（``ComputeStatistics(wdStatisticPages)``）。"""

        if self._total is None:
            self._bump()
            try:
                self._total = int(self.document.ComputeStatistics(2))
            except Exception:
                self._total = 0
        return self._total

    def page_at(self, position: int) -> int | None:
        """字符位置 → 页码。

        用与旧实现相同的 ``Range(position, position+1).Information(3)`` 精确读取
        （页边界二分在"Word 定位不到的空白页"上会算错），结果按位置缓存。
        """

        key = int(position)
        if key not in self._page_at:
            self._bump()
            try:
                end = int(self.document.Content.End)
                self._page_at[key] = int(
                    self.document.Range(key, min(key + 1, end)).Information(WD_ACTIVE_END_PAGE_NUMBER))
            except Exception:
                self._page_at[key] = None
        return self._page_at[key]

    # -- 段落页码 / 节号 --------------------------------------------------

    def page_of(self, index) -> int | None:
        """Word 段号 → 页码（按需读取并缓存）。"""

        try:
            key = int(index)
        except (TypeError, ValueError):
            return None
        if key not in self._pages:
            self._bump()
            try:
                self._pages[key] = int(self._range(key).Information(WD_ACTIVE_END_PAGE_NUMBER))
            except Exception:
                self._pages[key] = None
        return self._pages[key]

    def start_page_of(self, index) -> int | None:
        """Word 段号 → 段落首字符所在页，而不是跨页段落的末页。"""

        try:
            key = int(index)
        except (TypeError, ValueError):
            return None
        if key not in self._start_pages:
            try:
                start = int(self._range(key).Start)
                self._start_pages[key] = self.page_at(start)
            except Exception:
                self._start_pages[key] = None
        return self._start_pages[key]

    def section_of(self, index) -> int | None:
        """Word 段号 → 节号（按需读取并缓存）。"""

        try:
            key = int(index)
        except (TypeError, ValueError):
            return None
        if key not in self._section_of:
            self._bump()
            try:
                self._section_of[key] = int(self._range(key).Information(WD_ACTIVE_END_SECTION_NUMBER))
            except Exception:
                self._section_of[key] = None
        return self._section_of[key]

    # -- 节页范围 --------------------------------------------------------

    @property
    def sections(self) -> list[dict]:
        """``[{index, first_page, last_page}]``：由"有内容的页"的节号推出。

        只统计有内容的页，与旧实现（按段落所属节汇总）口径一致：空白页不归属任何节。
        """

        if self._sections is None:
            starts = dict(self._valid_pages())
            ranges: dict[int, list[int]] = {}
            for number in self.content_pages():
                start = starts.get(number)
                if start is None:
                    continue
                self._bump()
                try:
                    section = int(self.document.Range(start, start + 1)
                                  .Information(WD_ACTIVE_END_SECTION_NUMBER))
                except Exception:
                    continue
                ranges.setdefault(section, []).append(number)
            self._sections = [{"index": section,
                               "first_page": min(pages),
                               "last_page": max(pages)}
                              for section, pages in sorted(ranges.items())]
        return self._sections

    # -- 空白页 ----------------------------------------------------------

    def content_pages(self) -> list[int]:
        """有内容（文字或图形）的页号。

        能定位的页按"该页到下一页之间的文字/图形"判断；对 Word 定位不到的页
        （真正的空白页，或分节产生的幻影页）用"下一页之前那个字符的页码"探针确认，
        避免把其实有内容的幻影页误判成空白页。
        """

        if self._content is None:
            valid = self._valid_pages()
            pages: list[int] = []
            if valid:
                try:
                    content_end = int(self.document.Content.End)
                except Exception:
                    content_end = valid[-1][1]
                for position, (number, start) in enumerate(valid):
                    self._bump()
                    try:
                        end = valid[position + 1][1] if position + 1 < len(valid) else content_end
                        range_ = self.document.Range(start, end)
                        text = str(range_.Text).strip(_PAGE_MARKERS)
                    except Exception:
                        continue
                    # 图形计数：不同 Word 版本对 Range 暴露的集合不同（InlineShapes /
                    # ShapeRange），读不到就按"没有图形"处理，绝不影响文字判定。
                    shapes = 0
                    for attr in ("InlineShapes", "ShapeRange"):
                        try:
                            shapes += int(getattr(range_, attr).Count)
                        except Exception:
                            continue
                    if text or shapes:
                        pages.append(number)
                # 探针：定位不到的页，看"下一张能定位的页之前一个字符"落在第几页。
                total = int(self.total_pages or 0)
                located = {number for number, _ in valid}
                known = set(pages)
                for number in range(1, total + 1):
                    if number in located or number in known:
                        continue
                    probe = next((start for page, start in valid if page > number), None)
                    if probe is None or probe <= 0:
                        continue
                    self._bump()
                    try:
                        actual = int(self.document.Range(probe - 1, probe)
                                     .Information(WD_ACTIVE_END_PAGE_NUMBER))
                    except Exception:
                        continue
                    if actual == number:
                        pages.append(number)
            self._content = sorted(set(pages))
        return self._content

    def blank_pages(self) -> list[int]:
        """完全没有内容的页号（1 起，按页序）。"""

        total = self.total_pages or 0
        if not total:
            return []
        content = set(self.content_pages())
        return [number for number in range(1, int(total) + 1) if number not in content]

    # -- 序列化 ----------------------------------------------------------

    def as_dict(self) -> dict:
        """可 JSON 化的事实字典（写进 ``RunSummary.render_facts``，供缓存与报告）。"""

        return {"total_pages": self.total_pages,
                "sections": list(self.sections),
                "content_pages": list(self.content_pages()),
                "blank_pages": list(self.blank_pages()),
                # 旧字段保留空表：段落页码已改为按需查询，不再全量缓存。
                "pages": {}, "section_of": {}}

    def stats(self) -> dict:
        """给报告用的读取统计。"""

        return {"paragraphs_queried": len(self._pages), "com_calls": self.calls}


def paragraph_page(facts, index) -> int | None:
    """取某个 Word 段号所在页；越界或读不到时返回 None。

    同时支持 :class:`PageFacts`（按需查询）与旧的分页事实字典。
    """

    if isinstance(facts, PageFacts):
        return facts.page_of(index)
    try:
        return (facts.get("pages") or {}).get(int(index))
    except (AttributeError, TypeError, ValueError):
        return None


def paragraph_start_page(facts, index) -> int | None:
    """取段落首字符所在页；旧事实没有首页数据时才退回段落末页。"""

    if isinstance(facts, PageFacts):
        return facts.start_page_of(index)
    try:
        start_pages = facts.get("start_pages") or {}
        return start_pages.get(int(index), (facts.get("pages") or {}).get(int(index)))
    except (AttributeError, TypeError, ValueError):
        return None


def section_pages(facts, index) -> tuple[int | None, int | None]:
    """取第 ``index`` 节（1 起）的物理页范围。"""

    items = facts.sections if isinstance(facts, PageFacts) else (facts.get("sections") or [])
    for item in items:
        if item.get("index") == index:
            return item.get("first_page"), item.get("last_page")
    return None, None


__all__ = [
    "PageFacts", "paragraph_page", "paragraph_start_page",
    "read_page_facts", "section_pages",
]
