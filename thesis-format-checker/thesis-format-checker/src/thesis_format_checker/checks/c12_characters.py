"""十二、字符、符号、标点、数字与空格检测（规则 122–132）。

大类目标
========
中英文标点误用、标点成对性、全半角混用、中英文边界空格、多余空格与空行、
日期编号形式，以及单位与特殊符号书写。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：122、123、124、125、128、129
* 待实现：126、127、130、131、132
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class CharacterChecker(CategoryChecker):
    """十二、字符、符号、标点、数字与空格检测（122–132）。"""

    CATEGORY = "c12"

    def check_122(self) -> list[Finding]:
        """【122】中文句子误用英文标点，或英文句子误用中文标点（确定性检测）。"""
        return self.evidence("122")

    def check_123(self) -> list[Finding]:
        """【123】标点或括号缺失、多余、成对不完整（确定性检测）。"""
        return self.evidence("123")

    def check_124(self) -> list[Finding]:
        """【124】全角与半角字符混用，导致标点、字母或数字宽度不一致。"""
        return self.evidence("124")

    def check_125(self) -> list[Finding]:
        """【125】中文与英文、数字之间的空格使用不符合统一规则。"""
        return self.evidence("125")

    def check_126(self) -> list[Finding]:
        """【126】行首行尾、标点前后或词语之间存在多余空格——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_127(self) -> list[Finding]:
        """【127】英文单词之间、数字与单位之间存在缺失或错误空格——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_128(self) -> list[Finding]:
        """【128】连续出现多个空行，或用空行控制分页和段落间距。"""
        return self.evidence("128")

    def check_129(self) -> list[Finding]:
        """【129】日期、时间、年份或编号格式不统一。"""
        return self.evidence("129")

    def check_130(self) -> list[Finding]:
        """【130】计量单位大小写、正斜体、前缀或单位符号书写错误——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_131(self) -> list[Finding]:
        """【131】百分号、摄氏度、范围号、破折号或省略号使用不规范——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_132(self) -> list[Finding]:
        """【132】文档包含乱码、不可见控制字符或无意义特殊符号——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）
