"""四、页眉页脚与页码检测（规则 034–044）。

大类目标
========
页眉页脚是否按模板出现、内容与格式是否正确、奇偶页/首页变体是否分别设置、
是否错误链接到前一节，以及页码的数字形式、起始值与位置。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：034、035、036、037、038、039、040、041、042、043、044

本大类的 11 条规则全部有实现，是本项目完成度最高的一类。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class HeaderFooterChecker(CategoryChecker):
    """四、页眉页脚与页码检测（034–044）。"""

    CATEGORY = "c04"

    def check_034(self) -> list[Finding]:
        """【034】应设置页眉的页面缺少页眉，或不应有页眉的页面出现页眉。"""
        return self.evidence("034")

    def check_035(self) -> list[Finding]:
        """【035】页眉文字内容与学校名称/论文题目/当前章节名称不一致（旧版流程）。"""
        return self.evidence("035")

    def check_036(self) -> list[Finding]:
        """【036】页眉字体、字号、对齐方式或下边框样式不符合要求。"""
        return self.evidence("036")

    def check_037(self) -> list[Finding]:
        """【037】奇数页与偶数页的页眉内容或位置未按要求分别设置。"""
        return self.evidence("037")

    def check_038(self) -> list[Finding]:
        """【038】页眉未与前一节断开，或错误链接到前一节。"""
        return self.evidence("038")

    def check_039(self) -> list[Finding]:
        """【039】页脚包含多余文字、横线或与模板不一致的内容。"""
        return self.evidence("039")

    def check_040(self) -> list[Finding]:
        """【040】页码缺失、重复、跳号或出现两个页码。"""
        return self.evidence("040")

    def check_041(self) -> list[Finding]:
        """【041】前置部分未用罗马数字，或正文未用阿拉伯数字编号。"""
        return self.evidence("041")

    def check_042(self) -> list[Finding]:
        """【042】页码起始值错误，正文首页未按规定从 1 开始编号。"""
        return self.evidence("042")

    def check_043(self) -> list[Finding]:
        """【043】页码位置、字体、字号或对齐方式不符合模板要求。"""
        return self.evidence("043")

    def check_044(self) -> list[Finding]:
        """【044】封面、声明页等不应显示页码的页面显示了页码。"""
        return self.evidence("044")
