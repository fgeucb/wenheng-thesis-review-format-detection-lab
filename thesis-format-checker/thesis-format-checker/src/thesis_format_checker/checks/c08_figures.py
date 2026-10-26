"""八、图表检测（规则 078–088）。

大类目标
========
图题表题的位置与格式、图表编号的连续性与书写形式、三线表与表格结构、
图表尺寸与清晰度、跨页表格的处理，以及图注/数据来源标注。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：078、079、080、081、083、084
* 待实现：082、085、087、088
* 已删除：086
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class FigureTableChecker(CategoryChecker):
    """八、图表检测（078–088）。"""

    CATEGORY = "c08"

    def check_078(self) -> list[Finding]:
        """【078】图题未置于图片下方，或表题未置于表格上方。"""
        return self.evidence("078")

    def check_079(self) -> list[Finding]:
        """【079】图题、表题的字体、字号、对齐方式或段落间距不符合要求。"""
        return self.evidence("079")

    def check_080(self) -> list[Finding]:
        """【080】图表编号不连续，存在跳号、重号或章节号错误。"""
        return self.evidence("080")

    def check_081(self) -> list[Finding]:
        """【081】图表编号格式错误（多种写法无规则混用）。"""
        return self.evidence("081")

    def check_082(self) -> list[Finding]:
        """【082】正文提及的图表编号不存在，或图表未被引用——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_083(self) -> list[Finding]:
        """【083】表格未按要求采用三线表，存在多余竖线或横线。"""
        return self.evidence("083")

    def check_084(self) -> list[Finding]:
        """【084】表头、表身、合计行的字体、对齐、行高或底纹格式不规范。"""
        return self.evidence("084")

    def check_085(self) -> list[Finding]:
        """【085】图表超出页边距、尺寸过小或跨页位置不合理——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_086(self) -> list[Finding]:
        """【086】图片分辨率过低、文字模糊或颜色不适合打印——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_087(self) -> list[Finding]:
        """【087】跨页表格未重复标题行，或表格被不合理拆分——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_088(self) -> list[Finding]:
        """【088】图注、表注、数据来源或单位标注缺失或位置错误——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）
