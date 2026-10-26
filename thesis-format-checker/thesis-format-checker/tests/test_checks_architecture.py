"""架构自检：大类 / 规则族 / 细粒度规则 / 等级 / 目录与实现一致。

这些断言只依赖 ``thesis_checker.checks`` 包本身（不需要 Microsoft Word），
用来防止后续维护时漏加规则、漏写方法、等级缺失或目录与实现脱节。
"""

from __future__ import annotations

import ast
import importlib
import inspect
import re
from pathlib import Path

import pytest

import thesis_format_checker
from thesis_format_checker.checks import CATEGORY_CHECKERS, CATEGORIES, RULES
from thesis_format_checker.checks.base import CheckResult
from thesis_format_checker.checks.catalog import (
    ALL_FAMILY_IDS,
    DEFERRED_RULE_IDS,
    FINE_RULES,
    LEGACY_RULE_IDS,
    PENDING_RULE_IDS,
    PROVIDERS,
    REMOVED_REASONS,
    REMOVED_RULE_IDS,
    SKIPPED_SUMMARY_LINES,
    STATUS_LABELS,
    RuleStatus,
    fine_rules_of_category,
    implemented_rule_ids,
    rules_of_category,
    skipped_reasons,
)
from thesis_format_checker.core.models import Finding, Severity, Status


def test_catalog_covers_all_families_and_fine_rules():
    """15 个大类、165 个规则族；细粒度规则必须信息完整。"""

    assert len(CATEGORIES) == 15
    assert list(ALL_FAMILY_IDS) == [f"{number:03d}" for number in range(1, 166)]
    assert len(FINE_RULES) >= 300
    for category in CATEGORIES:
        assert len(rules_of_category(category.key)) == 11
        # 大类十四只保留“批注/修订/占位残留”，因此条数按实际清单校验
        assert len(fine_rules_of_category(category.key)) >= 3
    for code, item in FINE_RULES.items():
        assert item.code == code
        assert item.family in set(ALL_FAMILY_IDS)
        assert item.title and item.location
        assert isinstance(item.severity, Severity)
        assert item.status in STATUS_LABELS


def test_every_family_has_a_class_and_a_method():
    """一个大类一个类、一个规则族一个方法。"""

    for category in CATEGORIES:
        checker_cls = CATEGORY_CHECKERS.get(category.key)
        assert checker_cls is not None, f"大类 {category.key} 未注册检测器类"
        assert checker_cls.CATEGORY == category.key
        for spec in rules_of_category(category.key):
            method = getattr(checker_cls, spec.method, None)
            assert callable(method), f"大类 {category.key} 缺少方法 {spec.method}"
            assert method.__name__ == spec.method


def test_checker_classes_live_in_declared_modules():
    """目录里声明的大类模块名、类名必须和真实代码一致。"""

    for category in CATEGORIES:
        module = importlib.import_module(f"thesis_format_checker.checks.{category.module}")
        assert getattr(module, category.checker) is CATEGORY_CHECKERS[category.key]


def test_status_sets_partition_all_families():
    """已实现 / 待实现 / 待确认 / 已删除 四类互斥，且覆盖全部 165 个规则族。"""

    implemented = set(implemented_rule_ids())
    pending = set(PENDING_RULE_IDS)
    deferred = set(DEFERRED_RULE_IDS)
    removed = set(REMOVED_RULE_IDS)
    groups = [implemented, pending, deferred, removed]
    for index, left in enumerate(groups):
        for right in groups[index + 1:]:
            assert left & right == set()
    assert implemented | pending | deferred | removed == set(ALL_FAMILY_IDS)
    # 大类十四：154（残留批注/修订/占位）与 144/149/151 的“词表/写法统一”部分已实现，
    # 其余（用词、语法、术语、风格、名称规范）仍需大模型语义，保持已删除。
    assert {f"{i:03d}" for i in range(144, 155)} - removed == {"144", "149", "151", "154"}


def test_family_metadata_is_consistent_with_items():
    """规则族的等级、聚合标记必须与族内细粒度规则一致。"""

    for family, spec in RULES.items():
        assert isinstance(spec.severity, Severity)
        for item in spec.items:
            assert item.family == family
            assert item.category == spec.category
        if spec.aggregated:
            assert any(item.status is RuleStatus.AGGREGATED for item in spec.items)
        if spec.status is RuleStatus.PENDING:
            assert spec.pending_items


def _module_rule_ids(path: Path) -> set[str]:
    """从取证模块源码中提取该模块声明的规则编号（_TITLES/TITLES/RULES 等常量）。"""

    tree = ast.parse(path.read_text(encoding="utf-8"))
    found: set[str] = set()
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if not isinstance(target, ast.Name):
                continue
            if not target.id.upper().endswith(("TITLES", "RULES")):
                continue
            try:
                value = ast.literal_eval(node.value)
            except ValueError:  # 非字面量，跳过
                continue
            if isinstance(value, dict):
                # 只统计形如规则编号的键（001 / 005.2），避免 *_RULES 之类的辅助字典误判
                found.update(str(key) for key in value
                             if re.fullmatch(r"\d{3}(?:\.\d+)?", str(key)))
    return found


def test_provider_rules_match_detector_modules():
    """规则目录里登记的编号必须与取证模块内部声明的编号完全一致。"""

    package_dir = Path(thesis_format_checker.__file__).parent
    for provider in PROVIDERS:
        if provider.key == "legacy":
            assert provider.rules == LEGACY_RULE_IDS
            continue
        relative = Path(*provider.module.split(".")[1:]).with_suffix(".py")
        declared = _module_rule_ids(package_dir / relative)
        assert declared, f"{provider.module} 未声明任何规则编号"
        assert set(provider.rules) == declared, provider.module


def test_skipped_reasons_cover_pending_and_deferred_items():
    """未实现 / 待确认说明必须覆盖每一条细粒度规则，并保留固定汇总行。"""

    notes = skipped_reasons()
    items = [item for item in FINE_RULES.values()
             if item.status in (RuleStatus.PENDING, RuleStatus.DEFERRED)]
    assert items
    for item in items:
        assert any(f"【{item.code}】" in note for note in notes), item.code
    for line in SKIPPED_SUMMARY_LINES:
        assert line in notes
    assert len(notes) == len(set(notes))


def test_removed_families_document_their_reason():
    """删除的规则族必须留痕：编号与原因一一对应。"""

    assert set(REMOVED_RULE_IDS) == set(REMOVED_REASONS)
    assert len(REMOVED_RULE_IDS) >= 20
    assert all(reason for reason in REMOVED_REASONS.values())
    # 已删除的族不再产出结论
    assert set(REMOVED_RULE_IDS) & set(implemented_rule_ids()) == set()


def test_category_checker_reads_evidence_and_isolates_failures():
    """类方法默认取回既有证据；单条规则异常降级为“无法判定”而不中断。"""

    finding = Finding("067", RULES["067"].title, Status.FAIL, "既有证据")
    checker = CATEGORY_CHECKERS["c07"](CheckResult(findings=[finding]))
    assert checker.check_067() == [finding]
    # 每条规则都是本大类类上的独立方法（同一规则不可能跨类调用）。
    assert inspect.isfunction(type(checker).check_070)
    with pytest.raises(KeyError):
        checker.rule("001")
    assert checker.evidence("070") == []

    from thesis_format_checker.checks.c07_paragraphs import ParagraphFormatChecker

    class BrokenChecker(ParagraphFormatChecker):
        def check_067(self):  # type: ignore[override]
            raise ValueError("object cannot be read")

    broken = BrokenChecker(CheckResult())
    findings = broken.run()
    assert len(findings) == 1
    assert findings[0].rule_id == "067"
    assert findings[0].status is Status.UNKNOWN
    assert findings[0].message.startswith("【无法判定】")   # 抬头不再写规则编号


def test_all_checkers_are_runnable_without_document():
    """15 个大类都能在“无文档”的情况下运行：已实现规则返回空，未实现规则不报错。"""

    result = CheckResult()
    for key, checker_cls in CATEGORY_CHECKERS.items():
        assert checker_cls(result).run() == [], key
