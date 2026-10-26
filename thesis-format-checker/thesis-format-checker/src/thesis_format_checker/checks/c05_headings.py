"""五、标题与编号检测（规则 045–055）。

大类目标
========
各级标题的字体、字号、加粗、对齐与间距是否符合模板的对应层级，标题层级
是否跳级，编号是否连续、格式是否统一、是否使用多级列表自动编号。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：045、046、047、048、049、050、051、052、053、054、055

本大类的 11 条规则全部有实现。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class HeadingChecker(CategoryChecker):
    """五、标题与编号检测（045–055）。"""

    CATEGORY = "c05"

    def check_045(self) -> list[Finding]:
        """【045】一级标题的字体、字号、加粗或对齐不符合要求（旧版流程）。"""
        return self.evidence("045")

    def check_046(self) -> list[Finding]:
        """【046】二级标题的字体、字号、缩进或段前段后间距不符合要求（旧版流程）。"""
        return self.evidence("046")

    def check_047(self) -> list[Finding]:
        """【047】三级及以下标题格式与对应层级模板样式不一致（旧版流程）。"""
        return self.evidence("047")

    def check_048(self) -> list[Finding]:
        """【048】标题层级跳级（如一级标题后直接出现三级标题）。"""
        return self.evidence("048")

    def check_049(self) -> list[Finding]:
        """【049】同一层级标题使用了不同的字体、字号、缩进或编号形式。"""
        return self.evidence("049")

    def check_050(self) -> list[Finding]:
        """【050】章节编号不连续，存在跳号、重号或错误排序。"""
        return self.evidence("050")

    def check_051(self) -> list[Finding]:
        """【051】标题编号格式错误（“第1章”“1.1”等混用不当）。"""
        return self.evidence("051")

    def check_052(self) -> list[Finding]:
        """【052】标题编号与标题文字之间的空格、制表符或标点使用错误。"""
        return self.evidence("052")

    def check_053(self) -> list[Finding]:
        """【053】无编号标题被错误编号，或应编号标题缺少编号。"""
        return self.evidence("053")

    def check_054(self) -> list[Finding]:
        """【054】标题位于页面末行，与其后的正文、图表或公式分离。"""
        return self.evidence("054")

    def check_055(self) -> list[Finding]:
        """【055】使用手工输入编号而非多级列表，编号无法自动更新。"""
        return self.evidence("055")
