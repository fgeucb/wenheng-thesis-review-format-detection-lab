"""分类检测器的公共基类与结果容器。

本模块定义重构后的三个核心概念：

* :class:`CheckResult` —— 一次检测的完整结果：按批注顺序排列的问题列表、
  未实现/跳过说明，以及按“规则编号”“大类”建立的索引；
* :class:`CategoryChecker` —— **每个大类一个类**的基类；子类只声明
  :attr:`CategoryChecker.CATEGORY`，并为每条规则实现 ``check_<编号>`` 方法；
* 结果构造函数 :meth:`CategoryChecker.fail` / :meth:`unknown` / :meth:`skipped`
  —— 统一批注文案格式，保证迁移算法后生成的批注与旧实现同格式。

设计要点
========

1. **单一事实来源**：检测算法目前仍集中在原有的 ``detectors/`` 模块与旧版 Word
   流程中，:meth:`CategoryChecker.evidence` 从 :class:`CheckResult` 里按规则编号
   取回这些证据。重构只是把“入口”对象化，不复制算法，避免两处实现不一致。
2. **可逐条替换**：维护者可以把某条规则的算法直接写进对应方法里
   （用 ``self.fail`` / ``self.unknown`` 产出结论），此时把该规则从
   :data:`~thesis_format_checker.checks.catalog.PROVIDERS` 的 ``rules`` 中移除即可，
   引擎的调度、汇总和批注写入都不需要改动。
3. **异常隔离**：:meth:`CategoryChecker.run` 会捕获单条规则的异常并转成
   “无法判定”，避免一条规则出错导致整个检测中断。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Iterator

from thesis_format_checker.core.models import (
    SEVERITY_LABELS,
    Finding,
    Severity,
    Status,
    build_comment,
)
from thesis_format_checker.core.models import severity_label as _severity_label
from thesis_format_checker.checks.catalog import (
    CATEGORY_BY_KEY,
    RULES,
    CategorySpec,
    RuleSpec,
    RuleStatus,
    rules_of_category,
)

#: 未实现规则的统一说明前缀（与旧版“无法判定”文案保持一致）。
UNAVAILABLE_PREFIX = "本条暂未实现"
UNKNOWN_EXPECTED = "仅按模板中可确认的对应要求检测"


def severity_label(severity: Severity, status: Status = Status.FAIL) -> str:
    """批注抬头里的等级文字：错误 / 警告 / 提醒，或“无法判定”“未实现”。"""

    return _severity_label(severity, status)


def build_message(severity: Severity, rule_id: str, title: str, expected: Any, actual: Any,
                  status: Status = Status.FAIL, location: str = "") -> str:
    """按统一格式生成批注文案（一条批注只讲一件事）。

    统一模板::

        【错误】封面题目格式不符合模板
        应该：字号 19 磅、加粗
        实际：字号 16 磅、未加粗

    * 抬头 = 等级（错误/警告/提醒/无法判定/未实现）+ 这一条在讲什么（不含规则编号）；
    * ``应该`` = 模板要求（期望值）；
    * ``实际`` = 当前文档的真实值；
    * 批注本身锚定在出问题的文字上，普通批注不再写“位置：”一行；段号等定位信息
      保留在结构化结果的 ``location`` 字段里（JSON 报告）。
    """

    return build_comment(severity, rule_id, title, expected, actual, status=status, location=location)


@dataclass
class CheckResult:
    """一次文档检测的结构化结果。

    Attributes:
        findings: 全部检出项，**顺序即批注写入顺序**（与重构前一致）。
        skipped: 未实现/跳过规则的说明行，用于 CLI 摘要与批注兜底。
        baseline: 模板基准（各取证命名空间的原始数据），
            便于调试时查看“模板要求是从哪儿读出来的”。
    """

    findings: list[Finding] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    baseline: dict[str, Any] = field(default_factory=dict)
    _by_rule: dict[str, list[Finding]] = field(default_factory=dict, repr=False)

    def __post_init__(self) -> None:
        self.reindex()

    def reindex(self) -> None:
        """按规则编号重建索引。修改 ``findings`` 后需要重新调用。"""

        index: dict[str, list[Finding]] = {}
        for finding in self.findings:
            index.setdefault(finding.rule_id, []).append(finding)
        self._by_rule = index

    def for_rule(self, rule_id: str) -> list[Finding]:
        """取某条规则产出的全部问题（无则为空列表）。

        传入规则族编号时，会把它的细粒度规则（``005.1``、``005.2`` …）一并返回，
        因此检测代码拆成原子结论后，按族查询的旧写法依然有效。
        """

        exact = self._by_rule.get(rule_id)
        if exact is not None:
            return list(exact)
        prefix = f"{rule_id}."
        return [finding for finding in self.findings if finding.rule_id.startswith(prefix)]

    def for_category(self, key: str) -> list[Finding]:
        """按批注顺序取某个大类下的全部问题。"""

        category = CATEGORY_BY_KEY[key]
        wanted = set(category.rule_ids)
        return [finding for finding in self.findings if finding.rule_id in wanted]

    def failed_counts(self) -> dict[str, int]:
        """统计每条规则的问题数（只统计 FAIL，与旧版 CLI 行为一致）。"""

        counts: dict[str, int] = {}
        for finding in self.findings:
            if finding.status == Status.FAIL:
                counts[finding.rule_id] = counts.get(finding.rule_id, 0) + 1
        return counts

    def categories_with_findings(self) -> list[str]:
        """返回有检出项的大类 key，按大类序号排列。"""

        keys = {RULES[finding.rule_id].category
                for finding in self.findings if finding.rule_id in RULES}
        return sorted(keys)

    def describe(self) -> list[tuple[str, RuleSpec, int]]:
        """生成“规则 → 问题数”的报告行，供 CLI/日志打印。"""

        counts = self.failed_counts()
        rows: list[tuple[str, RuleSpec, int]] = []
        for category in CATEGORY_BY_KEY.values():
            for spec in rules_of_category(category.key):
                rows.append((category.name, spec, counts.get(spec.rule_id, 0)))
        return rows

    def __iter__(self) -> Iterator[Finding]:
        return iter(self.findings)


class CategoryChecker:
    """一个大类 = 一个类；一条规则 = 一个 ``check_<编号>`` 方法。

    子类只需：

    1. 声明 :attr:`CATEGORY`（目录 key，例如 ``"c07"``）；
    2. 为该大类的 11 条规则各写一个方法 ``check_0NN``，方法体内可以是
       ``return self.evidence("0NN")``（委托现有实现），也可以直接写算法。
    """

    #: 所属大类的 key，必须存在于 :data:`~thesis_format_checker.checks.catalog.CATEGORIES`。
    CATEGORY: str = ""

    def __init__(self, result: CheckResult) -> None:
        self.result = result

    # -- 元数据 -------------------------------------------------------------

    @property
    def spec(self) -> CategorySpec:
        """本大类的元数据。"""

        return CATEGORY_BY_KEY[self.CATEGORY]

    def rules(self) -> tuple[RuleSpec, ...]:
        """本大类按编号顺序排列的 11 条规则。"""

        return rules_of_category(self.CATEGORY)

    def rule(self, rule_id: str) -> RuleSpec:
        """取单条规则的元数据，并校验它确实属于本大类。"""

        spec = RULES[rule_id]
        if spec.category != self.CATEGORY:
            raise KeyError(f"规则 {rule_id} 不属于大类 {self.CATEGORY}")
        return spec

    # -- 执行 ---------------------------------------------------------------

    def run(self) -> list[Finding]:
        """逐条执行本大类的规则方法，单条异常降级为“无法判定”。"""

        findings: list[Finding] = []
        for spec in self.rules():
            try:
                findings.extend(getattr(self, spec.method)())
            except Exception as exc:  # 一条规则出错（含方法缺失）不影响其他规则
                findings.append(self._unavailable(spec, f"{type(exc).__name__}"))
        return findings

    # -- 取值与构造 ---------------------------------------------------------

    def evidence(self, rule_id: str) -> list[Finding]:
        """取该规则已经产出的证据（委托现有实现时使用）。

        返回空列表代表“本次检测该规则没有发现问题”，与旧实现语义一致。
        """

        return self.result.for_rule(rule_id)

    def fail(self, rule_id: str, expected: Any, actual: Any, paragraph: Any = None) -> Finding:
        """构造一条“不符合模板”的结论（与旧实现同格式）。"""

        spec = self.rule(rule_id)
        finding = Finding(spec.rule_id, spec.title, Status.FAIL, "",
                          severity=spec.severity, expected=expected, actual=actual)
        finding.message = build_message(spec.severity, spec.rule_id, spec.title, expected, actual)
        finding.anchor = paragraph
        return finding

    def unknown(self, rule_id: str, reason: str, paragraph: Any = None) -> Finding:
        """构造一条“无法判定”的结论，用于对象不可读等情形。"""

        spec = self.rule(rule_id)
        actual = "暂无法判定：" + str(reason)
        finding = Finding(spec.rule_id, spec.title, Status.UNKNOWN, "",
                          severity=spec.severity, expected=UNKNOWN_EXPECTED, actual=actual)
        finding.message = build_message(spec.severity, spec.rule_id, spec.title,
                                        UNKNOWN_EXPECTED, actual, status=Status.UNKNOWN)
        finding.anchor = paragraph
        return finding

    def skipped(self, rule_id: str, reason: str = "") -> Finding:
        """构造一条“本阶段未实现”的说明，用于规则尚未开发时。"""

        spec = self.rule(rule_id)
        detail = reason or spec.note or STATUS_REASONS[spec.status]
        actual = f"{UNAVAILABLE_PREFIX}：{detail}"
        finding = Finding(spec.rule_id, spec.title, Status.SKIPPED, "",
                          severity=spec.severity, expected=UNKNOWN_EXPECTED, actual=actual)
        finding.message = build_message(spec.severity, spec.rule_id, spec.title,
                                        UNKNOWN_EXPECTED, actual, status=Status.SKIPPED)
        finding.anchor = None
        return finding

    # -- 内部工具 -----------------------------------------------------------

    def _unavailable(self, spec: RuleSpec, reason: str) -> Finding:
        """单条规则执行失败时兜底，沿用旧版的“读取受限”文案。"""

        if spec.status in (RuleStatus.PENDING, RuleStatus.DEFERRED, RuleStatus.REMOVED):
            return self.skipped(spec.rule_id)
        return self.unknown(spec.rule_id, f"规则执行失败（{reason}）")


#: 各类未实现状态的默认说明。
STATUS_REASONS: dict[RuleStatus, str] = {
    RuleStatus.IMPLEMENTED: "该规则已实现，但本次未产出结论",
    RuleStatus.AGGREGATED: "该规则已实现（与同族其他条目合并输出），本次未产出结论",
    RuleStatus.LEGACY: "该规则由旧版 Word 有效格式流程检测，本次未产出结论",
    RuleStatus.PENDING: "本条可以确定性检测，尚未实现",
    RuleStatus.DEFERRED: "本条业务口径待确认",
    RuleStatus.REMOVED: "本条依赖大模型语义，已整体删除",
}


__all__ = [
    "CategoryChecker",
    "CheckResult",
    "STATUS_REASONS",
    "UNAVAILABLE_PREFIX",
    "UNKNOWN_EXPECTED",
    "build_message",
    "severity_label",
]
