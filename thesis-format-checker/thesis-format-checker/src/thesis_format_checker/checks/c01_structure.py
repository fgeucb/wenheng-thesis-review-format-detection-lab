"""一、结构完整性检测（规则 001–011）。

大类目标
========
判断论文是否包含学校要求的组成部分（封面、声明、摘要、关键词、目录、
致谢、附录、成果等），以及这些组成部分的排列顺序是否符合模板规定。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：001、002、005、006、009、010
* 待实现：003、004、007、011
* 待确认：008

细粒度拆分（本类已拆完，批注按下面的编号逐条输出）
================================================

* 001 → 001.1 缺少封面、001.2 缺少题名页
* 002 → 002.1 缺少原创性声明或独创性声明、002.2 缺少授权使用声明、002.3 声明只有标题没有正文
* 005 → 005.1/005.2 中文关键词缺少/数量不符、005.3/005.4 英文关键词缺少/数量不符、
  005.5 关键词标签后未填写内容
* 006 → 006.1 缺少目录、006.2 目录为手工输入、006.3 目录未使用标题样式或大纲级别
* 009 → 009.1 缺少致谢、009.2 缺少附录、009.3 缺少攻读学位期间成果
* 010 → 010.1 组成部分排列顺序不符合模板要求

维护提示
========
``check_0NN`` 方法默认从 :class:`~thesis_format_checker.checks.base.CheckResult`
取回既有证据；若要改动某条规则的算法，直接在该方法里用
``self.fail()`` / ``self.unknown()`` 产出结论，并把该编号从
``checks/catalog.py`` 中对应 provider 的 ``rules`` 里移除。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class StructureChecker(CategoryChecker):
    """一、结构完整性检测（001–011）。"""

    CATEGORY = "c01"

    def check_001(self) -> list[Finding]:
        """【001】缺少学校要求的封面或题名页（模板强制要求时判定缺失）。"""
        return self.evidence("001")

    def check_002(self) -> list[Finding]:
        """【002】缺少原创性声明、独创性声明或授权使用声明（或只有标题无正文）。"""
        return self.evidence("002")

    def check_003(self) -> list[Finding]:
        """【003】缺少中文摘要或摘要内容不完整——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_004(self) -> list[Finding]:
        """【004】缺少英文摘要或与中文摘要不对应——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_005(self) -> list[Finding]:
        """【005】缺少中/英文关键词或关键词数量不符合模板要求。"""
        return self.evidence("005")

    def check_006(self) -> list[Finding]:
        """【006】缺少目录，或目录不是由标题样式/大纲级别自动生成。"""
        return self.evidence("006")

    def check_007(self) -> list[Finding]:
        """【007】正文缺少绪论、结论等必要章节——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_008(self) -> list[Finding]:
        """【008】缺少参考文献部分——待确认（业务口径）"""
        return []  # 状态见 docs/规则清单.md：待确认（业务口径）

    def check_009(self) -> list[Finding]:
        """【009】致谢、附录、攻读学位期间成果等必需部分缺失。"""
        return self.evidence("009")

    def check_010(self) -> list[Finding]:
        """【010】各组成部分的排列顺序不符合模板声明的要求。"""
        return self.evidence("010")

    def check_011(self) -> list[Finding]:
        """【011】存在无内容章节、重复章节或遗漏章节——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）
