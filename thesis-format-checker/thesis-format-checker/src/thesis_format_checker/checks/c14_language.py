"""十四、文字与语言规范检测（规则 144–154）。

大类目标
========
错别字与用词、语法与句式、术语与缩略语、学术语言风格、数字与名称表述，
以及文档中残留的批注、修订痕迹与模板提示语。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 待实现：154（批注残留、修订痕迹、模板提示语或占位文字）——可以确定性检测。
* 已删除：144–153（错别字、用词、语法、术语、缩略语、语言风格、数字与名称表述）
  ——必须依赖大模型语义理解，按需求整体删除。

删除原因：这些规则需要理解句子语义，确定性格式检测无法给出可靠结论，误报会直接
影响用户对检测系统的信任。将来若接入大模型或语言模型服务，可按下面的建议恢复。

接入建议
========

1. 在本模块为要恢复的规则实现 ``check_14x``，用 ``self.fail`` / ``self.unknown``
   产出结论；
2. 在 ``checks/rule_specs.py`` 的 ``FINE_RULE_ROWS`` 中登记对应细粒度规则
   （编号、标题、等级、状态、定位方式），并把该族从 ``REMOVED_FAMILIES`` 中移除；
3. 运行 ``python tools/generate_rule_catalog.py`` 重新生成清单，再跑一次测试；
4. 如果结论来自模型，请把“模型不确定”的情况用 ``self.unknown`` 表达，
   不要用 ``self.fail``——保持“宁可无法判定，不误报”的既有原则。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class LanguageChecker(CategoryChecker):
    """十四、文字与语言规范检测（144–154，本阶段全部未实现）。"""

    CATEGORY = "c14"

    def check_144(self) -> list[Finding]:
        """【144】错别字、形近字误用、漏字或重复字——词表部分已实现（144.1/.2）"""
        return self.evidence("144")

    def check_145(self) -> list[Finding]:
        """【145】用词不当、词语搭配错误或口语化表达——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_146(self) -> list[Finding]:
        """【146】成分残缺、赘余、语序不当或句式杂糅——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_147(self) -> list[Finding]:
        """【147】句子过长、指代不清或歧义——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_148(self) -> list[Finding]:
        """【148】专有名词、学科术语或专业缩写写法不规范——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_149(self) -> list[Finding]:
        """【149】缩略语首次出现未给出全称或说明——首次出现全称检查已实现（149.1）"""
        return self.evidence("149")

    def check_150(self) -> list[Finding]:
        """【150】语言风格不符合学术写作要求——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_151(self) -> list[Finding]:
        """【151】数字、量词、年代、世纪或序数词表达不规范——写法统一检查已实现（151.1）"""
        return self.evidence("151")

    def check_152(self) -> list[Finding]:
        """【152】国家、机构、民族、地区等名称表述不规范——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_153(self) -> list[Finding]:
        """【153】政策、法规、党政用语或规范性文件名称引用不准确——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_154(self) -> list[Finding]:
        """【154】文档含有批注残留、修订痕迹、模板提示语或占位文字——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）
