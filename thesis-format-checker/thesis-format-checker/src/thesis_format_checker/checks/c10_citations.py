"""十、引用、注释与交叉引用检测（规则 100–110）。

大类目标
========
引用标注规范、注释类型（脚注/尾注）与编号、引用序号格式，以及交叉引用
是否使用域、是否显示域错误信息。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：105、106、107、108、110
* 待实现：103、104、109
* 已删除：100、101、102
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class CitationNoteChecker(CategoryChecker):
    """十、引用、注释与交叉引用检测（100–110）。"""

    CATEGORY = "c10"

    def check_100(self) -> list[Finding]:
        """【100】直接引用他人文字却未标注来源、页码或引号——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_101(self) -> list[Finding]:
        """【101】正文引用标注的位置错误——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_102(self) -> list[Finding]:
        """【102】顺序编码制与著者—出版年制混用——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_103(self) -> list[Finding]:
        """【103】引用序号的括号、上标、字体或字号格式不统一——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_104(self) -> list[Finding]:
        """【104】引用序号不连续、重复或未按首次出现顺序排列——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_105(self) -> list[Finding]:
        """【105】脚注与尾注混用，或注释类型不符合学校要求。"""
        return self.evidence("105")

    def check_106(self) -> list[Finding]:
        """【106】脚注编号未按页或全文连续编号（跳号、重号、重置错误）。"""
        return self.evidence("106")

    def check_107(self) -> list[Finding]:
        """【107】脚注文字的字体、字号、行距、缩进或分隔线格式错误。"""
        return self.evidence("107")

    def check_108(self) -> list[Finding]:
        """【108】交叉引用由手工输入，更新后引用编号未同步变化。"""
        return self.evidence("108")

    def check_109(self) -> list[Finding]:
        """【109】交叉引用指向错误的章节、图、表、公式等——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_110(self) -> list[Finding]:
        """【110】链接或交叉引用显示“错误！未找到引用源”等域错误信息。"""
        return self.evidence("110")
