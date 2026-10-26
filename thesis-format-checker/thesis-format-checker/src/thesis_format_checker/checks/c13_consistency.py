"""十三、一致性与关联性检测（规则 133–143）。

大类目标
========
同一个文档内部各处的表述、编号、格式、数据是否互相一致：题目、目录、
图表公式引用、同类对象格式、统计数字等。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：134、136、137、139
* 待确认：141
* 已删除：133、135、138、140、142、143

口径说明：本类的比较基准是“文档自身”，因此结论表述为“内部一致性要求
（非模板固定值）”，避免把模板示例当作论文内容标准。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class ConsistencyChecker(CategoryChecker):
    """十三、一致性与关联性检测（133–143）。"""

    CATEGORY = "c13"

    def check_133(self) -> list[Finding]:
        """【133】同一术语、人名、机构名称写法不一致——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_134(self) -> list[Finding]:
        """【134】论文题目在封面、声明、摘要、页眉和正文中的表述不一致。"""
        return self.evidence("134")

    def check_135(self) -> list[Finding]:
        """【135】中英文摘要的题目、关键词、数据或结论不对应——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_136(self) -> list[Finding]:
        """【136】目录中的章节标题、编号或页码与正文不一致。"""
        return self.evidence("136")

    def check_137(self) -> list[Finding]:
        """【137】正文引用的图、表、公式、章节编号与实际对象不一致。"""
        return self.evidence("137")

    def check_138(self) -> list[Finding]:
        """【138】图表中的数据、单位或名称与正文描述不一致——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_139(self) -> list[Finding]:
        """【139】同级标题、同类图表、同类公式或同类注释的格式不一致。"""
        return self.evidence("139")

    def check_140(self) -> list[Finding]:
        """【140】同一变量、缩写或符号在不同章节中的含义不一致——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_141(self) -> list[Finding]:
        """【141】正文引用与参考文献表对应关系错误——待确认（业务口径）"""
        return []  # 状态见 docs/规则清单.md：待确认（业务口径）

    def check_142(self) -> list[Finding]:
        """【142】章节之间出现重复内容、顺序颠倒或逻辑断裂——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_143(self) -> list[Finding]:
        """【143】统计总数、百分比、样本量或合计值前后不一致——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）
