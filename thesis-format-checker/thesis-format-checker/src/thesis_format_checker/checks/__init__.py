"""规则层：**一个大类一个类，一条规则一个方法**。

目录结构
========

``catalog.py``
    15 个大类、165 条规则的唯一权威清单（纯数据，不依赖 python-docx）。
``base.py``
    :class:`CheckResult` 结果容器与 :class:`CategoryChecker` 基类。
``cXX_*.py``（15 个文件）
    每个大类一个检测器类；每条规则一个 ``check_<编号>`` 方法，
    与 ``detectors/`` 中同名模块一一对应。
``registry.py``
    大类 key → 检测器类。

调度与报告在 :mod:`thesis_format_checker.engine`。

快速上手
========

查看全部规则与实现状态::

    python -m thesis_format_checker rules

在代码里按大类排查::

    from thesis_format_checker.engine.pipeline import CheckEngine

    engine = CheckEngine(template_ctx, document_ctx)
    result = engine.run()
    result.for_category("c07")            # 字体与段落类的全部结论
    engine.category("c07").check_074()   # 单条规则的方法
"""

from thesis_format_checker.checks.base import CategoryChecker, CheckResult
from thesis_format_checker.checks.catalog import (
    ALL_FAMILY_IDS,
    ALL_RULE_IDS,
    CATEGORIES,
    CATEGORY_BY_KEY,
    CategorySpec,
    FINE_RULES,
    PROVIDERS,
    PARTIALLY_PENDING_RULE_IDS,
    PENDING_RULE_IDS,
    REMOVED_REASONS,
    REMOVED_RULE_IDS,
    REMOVED_RULES,
    RULES,
    RuleItem,
    RuleSpec,
    RuleStatus,
    fine_rules_of_category,
    fine_rules_of_family,
    rules_of_category,
    severity_of,
    skipped_reasons,
)
from thesis_format_checker.checks.registry import CATEGORY_CHECKERS, build_checkers

__all__ = [
    "ALL_FAMILY_IDS",
    "ALL_RULE_IDS",
    "CATEGORIES",
    "CATEGORY_BY_KEY",
    "CATEGORY_CHECKERS",
    "CategoryChecker",
    "CategorySpec",
    "CheckResult",
    "FINE_RULES",
    "PROVIDERS",
    "REMOVED_REASONS",
    "REMOVED_RULE_IDS",
    "REMOVED_RULES",
    "RULES",
    "RuleItem",
    "RuleSpec",
    "RuleStatus",
    "build_checkers",
    "fine_rules_of_category",
    "fine_rules_of_family",
    "rules_of_category",
    "severity_of",
    "skipped_reasons",
]
