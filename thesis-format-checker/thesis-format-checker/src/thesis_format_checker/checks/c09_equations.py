"""九、公式检测（规则 089–099）。

大类目标
========
公式是否用公式编辑器排版、独立公式与编号的对齐、编号格式与序列、
正文引用是否指向存在的公式、公式字体字号与前后间距。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：090、091、092、093、096、097
* 待实现：089、094、095、098
* 已删除：099
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class EquationChecker(CategoryChecker):
    """九、公式检测（089–099）。"""

    CATEGORY = "c09"

    def check_089(self) -> list[Finding]:
        """【089】公式未使用公式编辑器排版——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_090(self) -> list[Finding]:
        """【090】独立公式未居中，或公式编号未在同一行右对齐。"""
        return self.evidence("090")

    def check_091(self) -> list[Finding]:
        """【091】公式编号格式不符合要求（“(1-1)”与“式（1）”混用等）。"""
        return self.evidence("091")

    def check_092(self) -> list[Finding]:
        """【092】公式编号不连续，存在跳号、重号或章节号错误。"""
        return self.evidence("092")

    def check_093(self) -> list[Finding]:
        """【093】正文引用的公式编号与实际公式编号不一致或不存在。"""
        return self.evidence("093")

    def check_094(self) -> list[Finding]:
        """【094】变量未使用斜体，向量/矩阵/常量的正斜体使用错误——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_095(self) -> list[Finding]:
        """【095】上标、下标、分式、根号等层级和位置错误——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_096(self) -> list[Finding]:
        """【096】公式中中文、英文、数字和符号的字体或字号不统一。"""
        return self.evidence("096")

    def check_097(self) -> list[Finding]:
        """【097】公式前后段落间距过大、过小或与正文行距不协调。"""
        return self.evidence("097")

    def check_098(self) -> list[Finding]:
        """【098】长公式未合理换行，超出页边距或被裁切——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_099(self) -> list[Finding]:
        """【099】公式后标点缺失或误用——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）
