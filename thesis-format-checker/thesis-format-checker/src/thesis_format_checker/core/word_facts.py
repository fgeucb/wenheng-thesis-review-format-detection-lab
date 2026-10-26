"""Word 段落事实表：属性按需读取一次，之后全程复用（性能优化 Tier 0）。

为什么需要
==========

旧版 Word 流程会把同一份文档从头到尾走**四遍**：正文段、标题、分页事实、空段扫描。
每一遍都逐段访问 COM（样例论文约 2 万次调用），其中大部分是重复读取同一批属性
（文本、样式名、大纲级别、是否在表格）。本模块把这些读取收敛成一张"读一次、到处用"
的事实表：

* :class:`WordFacts` —— 段落文本 / 样式名 / 大纲级别 / 是否在表格 / 起始位置，
  **每个属性第一次被问到时才读**，之后直接命中缓存；
* 页码与节号的按需查询在 :class:`thesis_format_checker.core.word_pages.PageFacts`。

精度约定
========

所有属性仍然走**与逐段读取完全相同的 COM 接口**（``Range.Text``、
``Style.NameLocal``、``OutlineLevel``、``Range.Information(12)`` …），只是不再重复读，
因此结论与优化前逐字节一致。反过来，凡是"换个接口读同一个值"的取巧做法
（例如用 ``Content.Text`` 批量切分冒充每段文本）都**没有**采用：
``tools/verify_word_facts.py`` 的对照实测显示，表格单元格与行结束标记会让批量文本
与 ``Range.Text`` 发生错位（每行表格差 1 个字符、位置累计漂移），因此不作为事实来源。
"""

from __future__ import annotations

from typing import Any, Iterator

from thesis_format_checker.core.effective_format import clean_word_text

#: Word 把"正文文本"（不在任何标题级别）表示为 10。
WD_OUTLINE_BODY_TEXT = 10

#: ``Range.Information`` 的取值含义（只列本模块用到的）。
WD_ACTIVE_END_SECTION_NUMBER = 2
WD_ACTIVE_END_PAGE_NUMBER = 3
WD_WITH_IN_TABLE = 12


class WordFacts:
    """Word 文档的段落事实表（懒加载 + 缓存）。"""

    def __init__(self, document, *, label: str = "") -> None:
        self.document = document
        self.label = label
        #: 统计访问 COM 的次数，便于验证"省了多少"。
        self.calls = 0
        self._count: int | None = None
        self._paragraphs: dict[int, Any] = {}
        self._ranges: dict[int, Any] = {}
        self._text: dict[int, str] = {}
        self._style: dict[int, str] = {}
        self._outline: dict[int, int | None] = {}
        self._in_table: dict[int, bool] = {}
        self._start: dict[int, int | None] = {}
        #: 由调度层在建立分页事实后挂上（见 ``word_pages.PageFacts``）。
        self.pagination: Any = None

    # -- 基础访问 --------------------------------------------------------

    def _bump(self) -> None:
        self.calls += 1

    @property
    def count(self) -> int:
        """段落总数（与 ``document.Paragraphs.Count`` 一致）。"""

        if self._count is None:
            self._bump()
            self._count = int(self.document.Paragraphs.Count)
        return self._count

    def indices(self) -> range:
        """1 起的段落号序列。"""

        return range(1, self.count + 1)

    def paragraph(self, index: int):
        """Word 段落对象（缓存；调用方拿到的是同一个 COM 对象）。"""

        item = self._paragraphs.get(index)
        if item is None:
            self._bump()
            item = self.document.Paragraphs(index)
            self._paragraphs[index] = item
        return item

    def range(self, index: int):
        """段落 Range（缓存，避免重复构造）。"""

        item = self._ranges.get(index)
        if item is None:
            self._bump()
            item = self.paragraph(index).Range
            self._ranges[index] = item
        return item

    # -- 逐属性（懒加载） -------------------------------------------------

    def text(self, index: int) -> str:
        """段落原文（含段落标记，与 ``Range.Text`` 一致；调用方自行 clean）。"""

        if index not in self._text:
            self._bump()
            try:
                self._text[index] = str(self.range(index).Text)
            except Exception:
                self._text[index] = ""
        return self._text[index]

    def clean_text(self, index: int) -> str:
        """去掉段落标记后的文字（等价于 ``clean_word_text(Range.Text)``）。"""

        return clean_word_text(self.text(index))

    def stripped_text(self, index: int) -> str:
        """再去掉首尾空白，供"是不是空段""像不像标题"这类判断使用。"""

        return self.clean_text(index).strip()

    def style_name(self, index: int) -> str:
        """段落样式名。

        先读 ``Range.Style``（正文/标题采集的历史口径）；个别段落——例如目录域里的
        空段——``Range.Style`` 会是 ``None``，此时退回 ``Paragraph.Style``（目录检测的
        历史口径，样例论文里只有 2 段如此）。两种读法都是同一个 COM 接口，
        只是"从 Range 取"还是"从 Paragraph 取"，实测只有这两段会不同。
        """

        if index not in self._style:
            self._bump()
            name = ""
            for source in (self.range(index), self.paragraph(index)):
                try:
                    style = source.Style
                except Exception:
                    style = None
                if style is None:
                    continue
                for attr in ("NameLocal", "Name"):
                    value = getattr(style, attr, None)
                    if value:
                        name = str(value)
                        break
                if name:
                    break
            # 读不到名字时返回空串：与 ``Style.NameLocal or ""`` 的读取口径一致，
            # 不让 "None" 这样的字符串混进样式比较。
            self._style[index] = name
        return self._style[index]

    def outline_level(self, index: int) -> int | None:
        """大纲级别（与 ``Paragraph.OutlineLevel`` 一致）；读不到返回 ``None``。"""

        if index not in self._outline:
            self._bump()
            try:
                self._outline[index] = int(self.paragraph(index).OutlineLevel)
            except Exception:
                self._outline[index] = None
        return self._outline[index]

    def in_table(self, index: int) -> bool:
        """段落是否位于表格内（与 ``Range.Information(12)`` 一致）。"""

        if index not in self._in_table:
            self._bump()
            try:
                self._in_table[index] = bool(self.range(index).Information(WD_WITH_IN_TABLE))
            except Exception:
                self._in_table[index] = False
        return self._in_table[index]

    def start(self, index: int) -> int | None:
        """段落起始字符位置（与 ``Range.Start`` 一致）。"""

        if index not in self._start:
            self._bump()
            try:
                self._start[index] = int(self.range(index).Start)
            except Exception:
                self._start[index] = None
        return self._start[index]

    # -- 便捷视图 --------------------------------------------------------

    def non_empty(self) -> Iterator[int]:
        """非空段落的段号（第一次调用时只读文本，其余属性仍按需）。"""

        for index in self.indices():
            if self.clean_text(index).strip():
                yield index

    def texts(self) -> list[str]:
        """全部段落原文（按段号顺序）。"""

        return [self.text(index) for index in self.indices()]

    def stats(self) -> dict:
        """给报告用的读取统计。"""

        return {"label": self.label, "paragraphs": self.count, "com_calls": self.calls}


__all__ = ["WD_OUTLINE_BODY_TEXT", "WordFacts"]
