"""大类检测器注册表：大类 key → 检测器类。

新增一个大类时，只需要：

1. 在 ``checks/catalog.py`` 的 :data:`~thesis_format_checker.checks.catalog.CATEGORIES`
   中登记元数据；
2. 新建 ``checks/cXX_*.py`` 并实现检测器类；
3. 在本文件把两者登记到一起。

检测引擎、CLI 规则列表、自检测试都从这里取数据，不需要改动其他位置。
"""

from __future__ import annotations

from thesis_format_checker.checks.base import CategoryChecker, CheckResult
from thesis_format_checker.checks.c01_structure import StructureChecker
from thesis_format_checker.checks.c02_front import FrontMatterChecker
from thesis_format_checker.checks.c03_pages import PageSetupChecker
from thesis_format_checker.checks.c04_headers import HeaderFooterChecker
from thesis_format_checker.checks.c05_headings import HeadingChecker
from thesis_format_checker.checks.c06_toc import TocChecker
from thesis_format_checker.checks.c07_paragraphs import ParagraphFormatChecker
from thesis_format_checker.checks.c08_figures import FigureTableChecker
from thesis_format_checker.checks.c09_equations import EquationChecker
from thesis_format_checker.checks.c10_citations import CitationNoteChecker
from thesis_format_checker.checks.c11_bibliography import BibliographyChecker
from thesis_format_checker.checks.c12_characters import CharacterChecker
from thesis_format_checker.checks.c13_consistency import ConsistencyChecker
from thesis_format_checker.checks.c14_language import LanguageChecker
from thesis_format_checker.checks.c15_appendix import AppendixChecker

#: 大类 key → 检测器类。顺序与《格式检测错误总结》中的大类顺序一致。
CATEGORY_CHECKERS: dict[str, type[CategoryChecker]] = {
    "c01": StructureChecker,
    "c02": FrontMatterChecker,
    "c03": PageSetupChecker,
    "c04": HeaderFooterChecker,
    "c05": HeadingChecker,
    "c06": TocChecker,
    "c07": ParagraphFormatChecker,
    "c08": FigureTableChecker,
    "c09": EquationChecker,
    "c10": CitationNoteChecker,
    "c11": BibliographyChecker,
    "c12": CharacterChecker,
    "c13": ConsistencyChecker,
    "c14": LanguageChecker,
    "c15": AppendixChecker,
}


def checker_class(key: str) -> type[CategoryChecker]:
    """按大类 key 取检测器类，key 不存在时抛出 ``KeyError``。"""

    try:
        return CATEGORY_CHECKERS[key]
    except KeyError as exc:
        raise KeyError(f"未注册的大类：{key}") from exc


def build_checkers(result: CheckResult) -> dict[str, CategoryChecker]:
    """为一次检测创建全部大类检测器实例（按大类顺序）。"""

    return {key: cls(result) for key, cls in CATEGORY_CHECKERS.items()}


__all__ = ["CATEGORY_CHECKERS", "build_checkers", "checker_class"]
