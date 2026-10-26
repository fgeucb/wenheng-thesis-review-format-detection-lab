"""十一、参考文献检测（规则 111–121）。

大类目标
========
参考文献著录结构是否符合 GB/T 7714—2015 或学校指定标准、必填字段是否齐全、
作者与期刊年卷期信息格式、标点使用，以及文末排列顺序。

规则与实现状态
==============

* 逐条细粒度规则（编号、等级、状态、定位方式）见 ``docs/规则清单.md``；
  数据源是 ``checks/rule_specs.py``，状态由 ``checks/catalog.py`` 汇总。
* 已实现：111、113、114、115、117、118、121
* 待实现：112
* 待确认：120
* 已删除：116、119

边界说明：本类只读取文末参考文献区域，不联网核对文献真实性与完整性
（见 159 的说明）。
"""

from __future__ import annotations

from thesis_format_checker.core.models import Finding
from thesis_format_checker.checks.base import CategoryChecker


class BibliographyChecker(CategoryChecker):
    """十一、参考文献检测（111–121）。"""

    CATEGORY = "c11"

    def check_111(self) -> list[Finding]:
        """【111】参考文献著录格式不符合 GB/T 7714—2015 或学校标准。"""
        return self.evidence("111")

    def check_112(self) -> list[Finding]:
        """【112】文献类型标识错误或缺失（[J]、[M]、[D] 等）——待实现（可确定性检测）"""
        return []  # 状态见 docs/规则清单.md：待实现（可确定性检测）

    def check_113(self) -> list[Finding]:
        """【113】作者姓名著录格式错误，作者人数/省略规则使用不当。"""
        return self.evidence("113")

    def check_114(self) -> list[Finding]:
        """【114】题名、期刊名、出版社、年份、卷期或页码信息缺失。"""
        return self.evidence("114")

    def check_115(self) -> list[Finding]:
        """【115】期刊的年、卷、期、起止页码排列顺序或标点格式错误。"""
        return self.evidence("115")

    def check_116(self) -> list[Finding]:
        """【116】不同类型文献套用了错误模板——已删除（依赖大模型语义）"""
        return []  # 状态见 docs/规则清单.md：已删除（依赖大模型语义）

    def check_117(self) -> list[Finding]:
        """【117】DOI、URL 或引用日期缺失、失效或书写格式不规范。"""
        return self.evidence("117")

    def check_118(self) -> list[Finding]:
        """【118】中文文献使用英文标点，或英文文献误用中文标点。"""
        return self.evidence("118")

    def check_119(self) -> list[Finding]:
        """【119】英文文献的大小写与期刊缩写不统一——作者姓名写法一致性已实现（119.1）"""
        return self.evidence("119")

    def check_120(self) -> list[Finding]:
        """【120】正文引用与参考文献表无法一一对应——待确认（业务口径）"""
        return []  # 状态见 docs/规则清单.md：待确认（业务口径）

    def check_121(self) -> list[Finding]:
        """【121】参考文献排序方式错误（文末编号或作者字母顺序）。"""
        return self.evidence("121")
