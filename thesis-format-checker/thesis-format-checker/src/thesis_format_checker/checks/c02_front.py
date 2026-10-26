"""二、封面与前置部分检测（规则 012–022）。

大类目标
========
封面题目与基本信息是否填写且与模板一致、封面元素的版式、原创性声明的
必填内容、中英文摘要的排版一致性，以及前置部分是否错误出现页眉页脚页码。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：012、013、015、016、018、019、021、022
* 待实现：014、017、020
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class FrontMatterChecker(CategoryChecker):
    """二、封面与前置部分检测（012–022）。"""

    CATEGORY = "c02"

    def check_012(self) -> list[Finding]:
        """【012】封面题目缺失，或与正文、题名页中的题目不一致。"""
        return self.evidence("012")

    def check_013(self) -> list[Finding]:
        """【013】封面题目的字体、字号、加粗或行距不符合模板。"""
        return self.evidence("013")

    def check_014(self) -> list[Finding]:
        """【014】作者姓名、学号、学院等基本信息缺失或填写错误——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_015(self) -> list[Finding]:
        """【015】指导教师姓名、职称或单位信息缺失或格式不统一。"""
        return self.evidence("015")

    def check_016(self) -> list[Finding]:
        """【016】提交/答辩/学位授予日期缺失或格式错误。"""
        return self.evidence("016")

    def check_017(self) -> list[Finding]:
        """【017】学校名称、校徽的位置/尺寸/颜色/清晰度不符——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_018(self) -> list[Finding]:
        """【018】封面元素的对齐方式、上下间距或左右位置与模板不一致。"""
        return self.evidence("018")

    def check_019(self) -> list[Finding]:
        """【019】密级、分类号、UDC 等字段缺失或填写位置错误。"""
        return self.evidence("019")

    def check_020(self) -> list[Finding]:
        """【020】原创性声明中的签名、日期或必填信息缺失——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_021(self) -> list[Finding]:
        """【021】中英文摘要的标题、正文或关键词排版格式不一致。"""
        return self.evidence("021")

    def check_022(self) -> list[Finding]:
        """【022】前置部分错误显示页眉、页脚或阿拉伯数字页码。"""
        return self.evidence("022")
