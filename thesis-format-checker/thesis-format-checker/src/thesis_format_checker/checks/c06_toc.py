"""六、目录检测（规则 056–066）。

大类目标
========
目录是否由 Word 域自动生成且来源正确（标题样式/大纲级别）、条目是否与正文
标题一一对应、页码是否正确、层级范围是否符合要求、条目与标题的格式。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：056、057、058、059、060、061、062、063、064、065、066

本大类的 11 条规则全部有实现。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class TocChecker(CategoryChecker):
    """六、目录检测（056–066）。"""

    CATEGORY = "c06"

    def check_056(self) -> list[Finding]:
        """【056】目录未自动生成，而是手工输入或手工修改。"""
        return self.evidence("056")

    def check_057(self) -> list[Finding]:
        """【057】正文已有标题未出现在目录中，或目录含正文不存在的标题。"""
        return self.evidence("057")

    def check_058(self) -> list[Finding]:
        """【058】目录中的标题文字、编号或层级与正文不一致。"""
        return self.evidence("058")

    def check_059(self) -> list[Finding]:
        """【059】目录页码与正文实际页码不对应。"""
        return self.evidence("059")

    def check_060(self) -> list[Finding]:
        """【060】目录层级显示范围不符合要求（层级过多或过少）。"""
        return self.evidence("060")

    def check_061(self) -> list[Finding]:
        """【061】目录各级条目的字体、字号、加粗或缩进格式错误。"""
        return self.evidence("061")

    def check_062(self) -> list[Finding]:
        """【062】目录条目与页码之间缺少点引导线或引导线样式不一致。"""
        return self.evidence("062")

    def check_063(self) -> list[Finding]:
        """【063】目录页码未右对齐，或不同条目的页码位置不齐。"""
        return self.evidence("063")

    def check_064(self) -> list[Finding]:
        """【064】“目录”标题的字体、字号、位置或段落间距不符合模板。"""
        return self.evidence("064")

    def check_065(self) -> list[Finding]:
        """【065】目录中出现孤立条目、空白条目、重复条目或错误换行。"""
        return self.evidence("065")

    def check_066(self) -> list[Finding]:
        """【066】修改正文后未更新目录，标题或页码仍为旧内容。"""
        return self.evidence("066")
