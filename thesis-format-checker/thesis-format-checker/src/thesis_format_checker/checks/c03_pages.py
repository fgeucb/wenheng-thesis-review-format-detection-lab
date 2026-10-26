"""三、页面设置检测（规则 023–033）。

大类目标
========
纸张大小与方向、四向页边距、装订线、页眉页脚与纸张边缘的距离、分节与
分页方式的合理性，以及内容是否超出可打印区域。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：023、024、025、026、027、029、030、031、033
* 待实现：028、032
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class PageSetupChecker(CategoryChecker):
    """三、页面设置检测（023–033）。"""

    CATEGORY = "c03"

    def check_023(self) -> list[Finding]:
        """【023】纸张大小不是 A4 或页面宽高设置错误（旧版 Word 流程）。"""
        return self.evidence("023")

    def check_024(self) -> list[Finding]:
        """【024】页面方向错误（应为纵向却设为横向，或反之）。"""
        return self.evidence("024")

    def check_025(self) -> list[Finding]:
        """【025】上、下、左、右页边距不符合模板（旧版 Word 流程）。"""
        return self.evidence("025")

    def check_026(self) -> list[Finding]:
        """【026】装订线位置或宽度设置错误/缺失。"""
        return self.evidence("026")

    def check_027(self) -> list[Finding]:
        """【027】页眉距顶端、页脚距底端的距离不符合要求。"""
        return self.evidence("027")

    def check_028(self) -> list[Finding]:
        """【028】同一文档存在不必要的页面尺寸/页边距差异——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_029(self) -> list[Finding]:
        """【029】分节符使用错误导致页面设置、页码或页眉页脚异常变化。"""
        return self.evidence("029")

    def check_030(self) -> list[Finding]:
        """【030】章节未另起一页，或误用大量空行代替分页符。"""
        return self.evidence("030")

    def check_031(self) -> list[Finding]:
        """【031】奇偶页、首页不同等版式选项与模板要求不一致。"""
        return self.evidence("031")

    def check_032(self) -> list[Finding]:
        """【032】存在非必要空白页或应保留的空白页被删除——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_033(self) -> list[Finding]:
        """【033】页面内容超出可打印区域，存在被裁切的风险。"""
        return self.evidence("033")
