"""十五、附录及特殊内容检测（规则 155–165）。

大类目标
========
附录与正文的引用关系、附录编号与标题格式、附录内图表公式的独立编号、
附录对象格式与正文规则的一致性，以及致谢、成果条目、特殊材料、特殊版式
和文档末尾残留内容。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：155、156、157、158、159、164
* 待实现：160、161、165
* 已删除：162、163

边界说明：159 只比较可读取的格式规则，不判断附录资料来源的真实性、
完整性，也不做参考文献引用判定。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class AppendixChecker(CategoryChecker):
    """十五、附录及特殊内容检测（155–165）。"""

    CATEGORY = "c15"

    def check_155(self) -> list[Finding]:
        """【155】正文已引用附录但文末缺少相应附录，或附录未被正文引用。"""
        return self.evidence("155")

    def check_156(self) -> list[Finding]:
        """【156】附录编号不连续，或“附录A、附录B”等编号格式错误。"""
        return self.evidence("156")

    def check_157(self) -> list[Finding]:
        """【157】附录标题的字体、字号、对齐方式或编号格式不符合要求。"""
        return self.evidence("157")

    def check_158(self) -> list[Finding]:
        """【158】附录中的图、表、公式未采用独立编号（如“图A.1”）。"""
        return self.evidence("158")

    def check_159(self) -> list[Finding]:
        """【159】附录图表、公式、参考资料格式与正文共同规则冲突。"""
        return self.evidence("159")

    def check_160(self) -> list[Finding]:
        """【160】致谢标题、正文的字体、行距、措辞或署名格式不符——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_161(self) -> list[Finding]:
        """【161】攻读学位期间成果条目的格式、排序或作者标注不规范——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_162(self) -> list[Finding]:
        """【162】问卷、访谈提纲、源代码等特殊材料排版混乱——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_163(self) -> list[Finding]:
        """【163】保密内容、个人隐私或敏感数据未按要求脱敏——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_164(self) -> list[Finding]:
        """【164】超大表格、横向页面、折页或彩色页面未按特殊版式设置。"""
        return self.evidence("164")

    def check_165(self) -> list[Finding]:
        """【165】文档末尾存在多余空白页、无关内容或未删除的模板示例——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）
