"""七、字体与段落检测（规则 067–077）。

大类目标
========
正文的中文字体、英文与数字字体、字号、对齐方式、首行缩进、行距、段前段后
间距，以及不应缩进对象的缩进与分页设置。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：067、068、069、070、071、072、073、074、075、076、077

注意：069 同时由两条路径产出，:attr:`~thesis_format_checker.checks.catalog.RuleSpec.providers`
会列出全部来源，便于排查。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class ParagraphFormatChecker(CategoryChecker):
    """七、字体与段落检测（067–077）。"""

    CATEGORY = "c07"

    def check_067(self) -> list[Finding]:
        """【067】正文中文字体不是学校规定的指定字体（旧版流程）。"""
        return self.evidence("067")

    def check_068(self) -> list[Finding]:
        """【068】正文英文及数字字体不是规定的指定字体（旧版流程）。"""
        return self.evidence("068")

    def check_069(self) -> list[Finding]:
        """【069】字号不符合对应样式要求（正文字号 + 注释字号两条证据）。"""
        return self.evidence("069")

    def check_070(self) -> list[Finding]:
        """【070】同一用途文字出现字体、字号、颜色或加粗格式不一致。"""
        return self.evidence("070")

    def check_071(self) -> list[Finding]:
        """【071】正文未按要求两端对齐，或误设为居中/左对齐/右对齐。"""
        return self.evidence("071")

    def check_072(self) -> list[Finding]:
        """【072】正文首行未缩进 2 个字符，或用空格、制表符模拟缩进（旧版流程）。"""
        return self.evidence("072")

    def check_073(self) -> list[Finding]:
        """【073】标题、图题、表题等不应缩进的对象被设置了首行缩进。"""
        return self.evidence("073")

    def check_074(self) -> list[Finding]:
        """【074】行距不是学校规定的数值（旧版流程，按模板有效行距比较）。"""
        return self.evidence("074")

    def check_075(self) -> list[Finding]:
        """【075】段前、段后间距不符合要求，或用空行代替段落间距（旧版流程）。"""
        return self.evidence("075")

    def check_076(self) -> list[Finding]:
        """【076】段落中出现异常悬挂缩进、左缩进或右缩进。"""
        return self.evidence("076")

    def check_077(self) -> list[Finding]:
        """【077】孤行控制、段中不分页或与下段同页等分页设置使用不当。"""
        return self.evidence("077")
