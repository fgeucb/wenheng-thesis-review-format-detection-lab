"""检测调度：模板基准 → 逐类取证 → 旧版结果适配 → 结果汇总。

职责边界
========

* **本模块只做调度**：决定“按什么顺序、调用哪些取证实现、异常怎么降级、
  结果怎么汇总”，不实现任何一条规则的具体算法。
* 规则算法在 ``detectors/``（按大类分模块）；一条规则的结论由
  ``checks/cXX_*.py`` 里的 ``check_<编号>`` 方法对外提供。
* 页面写入、批注落盘等 I/O 也集中在这里，便于替换或单独测试。

调用顺序
========

`:data:`SECTION_RULES` 的先后顺序就是取证顺序，也就是问题（批注）的生成顺序。
顺序与旧版实现保持一致，保证输出文档的批注排列不因重构而改变。
"""

from __future__ import annotations

import importlib
import os
import re
import threading
import time
from copy import deepcopy
from dataclasses import asdict, is_dataclass
from io import BytesIO
from pathlib import Path
from tempfile import NamedTemporaryFile, TemporaryDirectory
from zipfile import ZipFile

from docx import Document
from thesis_format_checker.core.xmlutil import qn
from lxml import etree

from thesis_format_checker.checks.base import CategoryChecker, CheckResult, build_message
from thesis_format_checker.checks.catalog import (
    ALL_RULE_IDS,
    LEGACY_PROVIDER_KEY,
    LEGACY_RULE_IDS,
    PROVIDERS,
    RULES,
    ProviderSpec,
    RuleSpec,
    severity_of,
)
from thesis_format_checker.checks.catalog import skipped_reasons as catalog_skipped_reasons
from thesis_format_checker.checks.registry import build_checkers
from thesis_format_checker.core.cache import ResultCache, summary_from_payload, summary_payload
from thesis_format_checker.core.docx_context import DocxContext, UNREADABLE, unknown
from thesis_format_checker.core.models import Finding, RunSummary, Status
from thesis_format_checker.core.word_session import WordAutomationError

#: 大类分组 → 该组的规则编号（分组信息与顺序统一由规则目录提供）。
SECTION_RULES: dict[str, tuple[str, ...]] = {
    provider.key: provider.rules
    for provider in PROVIDERS
    if provider.key != LEGACY_PROVIDER_KEY
}

#: 由旧版 Word 有效格式流程提供的规则编号。
WORD_EFFECTIVE_RULE_IDS: tuple[str, ...] = LEGACY_RULE_IDS

def _real_word_flow(legacy) -> bool:
    """判断 ``legacy`` 是不是项目自带的 Word 流程（而不是测试替身）。

    结果缓存只在真实 Word 流程下默认开启：测试里注入的替身函数行为与实现无关，
    若也走缓存会互相污染。
    """

    return (getattr(legacy, "__module__", "") == "thesis_format_checker.detectors.word_effective.runner"
            and getattr(legacy, "__name__", "") == "compare_word_documents")


def _prewarm_docx_scans(*contexts) -> None:
    """把 docx 侧最贵的全篇扫描先算出来（结果只进上下文缓存，不产生任何结论）。

    性能优化 Tier 1：Word 阶段大部分时间在等 COM，正好用这段等待把"同一份文档
    只算一次"的扫描并行做完。检测顺序、结论都与串行执行完全一致——这里只写缓存。
    """

    try:
        from thesis_format_checker.detectors.equations import _read_records
        from thesis_format_checker.detectors.fields import scan_body_fields
        from thesis_format_checker.detectors.figures import collect_caption_records
        from thesis_format_checker.detectors.headings import (
            _toc_paragraph_ids,
            collect_heading_records,
        )
        from thesis_format_checker.detectors.toc import _snapshot
    except Exception:
        return
    for context in contexts:
        for produce in (collect_heading_records, _toc_paragraph_ids, collect_caption_records,
                        scan_body_fields, _snapshot, _read_records):
            try:
                produce(context)
            except Exception:
                continue


def _cache_params(mapping: dict | None, **values):
    """把影响检测结果的可选参数整理成可比较的缓存参数（键统一成字符串）。"""

    from thesis_format_checker.core.cache import clean_params

    payload = {name: clean_params(value) for name, value in values.items()}
    if mapping:
        payload["mapping"] = clean_params(mapping)
    return payload


# ---------------------------------------------------------------------------
# 基准提取与逐类取证
# ---------------------------------------------------------------------------

def _providers(sections: tuple[str, ...] | list[str] | None = None) -> tuple[ProviderSpec, ...]:
    """按调用顺序返回取证实现；``sections`` 可只取其中几组（用于单测）。"""

    if sections is None:
        return PROVIDERS
    wanted = set(sections)
    return tuple(provider for provider in PROVIDERS if provider.key in wanted)


def _load(provider: ProviderSpec):
    """导入取证实现所在模块（延迟导入，便于替换与测试打桩）。"""

    return importlib.import_module(provider.module)


def _extract(provider: ProviderSpec, context, mapping) -> dict:
    """调用某个取证实现的模板基准提取函数，并归一化返回的命名空间。"""

    module = _load(provider)
    extract = getattr(module, provider.extract)
    additions = extract(context, mapping=mapping) if provider.extract_takes_mapping else extract(context)
    if provider.namespace == "front":
        # 封面/前置部分历史上返回裸数据字典，这里补上命名空间键。
        return {"front": additions}
    return additions


def extract_baseline(context, existing: dict | None = None, mapping: dict | None = None,
                     sections: tuple[str, ...] | list[str] | None = None) -> dict:
    """提取模板基准，按 :data:`SECTION_RULES` 的顺序追加命名空间。

    Args:
        context: 模板 A 的 :class:`~thesis_format_checker.core.docx_context.DocxContext`。
        existing: 已有基准，函数只做追加，绝不覆盖调用方的键。
        mapping: 可选的区域/节映射配置。
        sections: 只提取指定的几组（单测使用），默认为全部。

    Returns:
        基准字典，键为大类名（``structure``、``headers``、``equations`` …）。
    """

    baseline = dict(existing or {})
    for provider in _providers(sections):
        if provider.extract is None:
            continue
        try:
            additions = _extract(provider, context, mapping)
        except Exception as exc:  # 单组读不出来不影响其他组
            additions = {provider.namespace: {
                "unreadable": f"{UNREADABLE}；模板读取异常：{type(exc).__name__}"}}
        for key, value in additions.items():
            if key in baseline:
                raise ValueError(f"模板基准键冲突，未覆盖已有值：{key}")
            baseline[key] = value
    return baseline


def check_sections(baseline: dict, context, mapping: dict | None = None,
                   sections: tuple[str, ...] | list[str] | None = None) -> list[Finding]:
    """按顺序对待检测文档执行各组的检查函数，返回有序问题列表。

    每一组的异常都会被降级为“无法判定”，不会中断其他组；这也是
    “一条规则读不出来时，其余规则照常给结论”的实现位置。
    """

    findings: list[Finding] = []
    for provider in _providers(sections):
        if provider.extract is None:
            continue
        data = baseline.get(provider.namespace, {})
        if isinstance(data, dict) and data.get("unreadable"):
            findings.extend(
                unknown(rule, "模板对象读取受限", data["unreadable"]) for rule in provider.rules
            )
            continue
        try:
            module = _load(provider)
            check = getattr(module, provider.check)
            result = check(baseline, context, mapping=mapping) if provider.check_takes_mapping \
                else check(baseline, context)
            findings.extend(result)
        except Exception as exc:
            findings.extend(
                unknown(rule, "对象读取受限", f"{UNREADABLE}；读取异常：{type(exc).__name__}")
                for rule in provider.rules
            )
    return findings


def skipped_reasons() -> list[str]:
    """未实现/跳过规则的说明列表（由规则目录统一生成）。"""

    return catalog_skipped_reasons()


def legacy_unknown_findings(unknowns: list[str],
                            rule_ids: tuple[str, ...] = WORD_EFFECTIVE_RULE_IDS) -> list[Finding]:
    """把旧版流程写在文本里的“无法判定”还原成报告中的结构化条目。

    旧版 Word 流程只把限制写进 ``RunSummary.unknowns`` 文本，CLI 能看到而
    本函数按 ``【编号】`` 前缀还原条目，并尽量解析
    “第 N 段/第 N 节”的定位信息。
    """

    findings: list[Finding] = []
    for reason in list(unknowns):
        match = re.search(r"【([0-9/]+)】", reason)
        if not match:
            continue
        # 批注正文不写规则编号：去掉旧流程写在文本里的“【035】”前缀。
        detail = reason.replace(match.group(0), "", 1).strip() or reason
        for rule in match.group(1).split("/"):
            if rule not in rule_ids:
                continue
            item = unknown(rule, "既有检测基准或读取受限", detail)
            paragraph = re.search(r"第\s*(\d+)\s*段", reason)
            section = re.search(r"第\s*(\d+)\s*节", reason)
            if paragraph:
                item.paragraph_index = int(paragraph.group(1))
            elif section:
                item.section_index = int(section.group(1))
            findings.append(item)
    return findings


def adapt_legacy_findings(findings: list[Finding], baseline: dict, context,
                          mapping: dict | None = None) -> list[Finding]:
    """适配旧版流程的结论（目前用于页眉【035】的动态内容）。

    旧版检测基于 Word 缓存的页眉文字，动态域（章节名、页码）无法静态比对，
    因此这里统一交给页眉取证模块判定：无法证明不一致时降级为“无法判定”，
    而不是把缓存值当作错误。
    """

    from thesis_format_checker.detectors.headers import adapt_legacy_header_findings

    try:
        return adapt_legacy_header_findings(findings, baseline, context, mapping=mapping)
    except Exception as exc:
        result: list[Finding] = []
        for finding in findings:
            if finding.rule_id == "035":
                finding = deepcopy(finding)
                finding.status = Status.UNKNOWN
                finding.actual = (f"{UNREADABLE}；页眉动态内容适配异常：{type(exc).__name__}；"
                                  "原检测证据：" + " ".join(finding.message.split()))
                finding.message = unknown("035", finding.title, finding.actual).message
            result.append(finding)
        if not any(finding.rule_id == "035" for finding in result):
            result.append(unknown("035", "页眉读取受限", f"{UNREADABLE}；{type(exc).__name__}"))
        return result


# ---------------------------------------------------------------------------
# 引擎：把上面各步组装成一次完整检测
# ---------------------------------------------------------------------------

class CheckEngine:
    """一次对照检测的调度器。

    典型用法::

        engine = CheckEngine(template_context, document_context, mapping)
        result = engine.run()
        result.findings              # 有序问题列表（即批注写入顺序）
        result.for_rule("074")       # 只看某条规则
        result.for_category("c07")   # 只看某个大类
        engine.category("c07").check_074()   # 直接调用单条规则
    """

    def __init__(self, template_context, document_context, mapping: dict | None = None,
                 baseline: dict | None = None) -> None:
        self.template_context = template_context
        self.document_context = document_context
        self.mapping = mapping
        self._baseline = baseline
        self._checkers: dict[str, CategoryChecker] | None = None

    @property
    def baseline(self) -> dict:
        """模板基准，首次访问时按需提取。"""

        if self._baseline is None:
            self._baseline = extract_baseline(self.template_context, mapping=self.mapping)
        return self._baseline

    @staticmethod
    def providers() -> tuple[ProviderSpec, ...]:
        """本次检测用到的取证实现（按调用顺序）。"""

        return PROVIDERS

    def category(self, key: str) -> CategoryChecker:
        """取单个大类检测器，例如 ``engine.category("c07")``。"""

        if self._checkers is None:
            self._checkers = build_checkers(CheckResult(baseline=self._baseline or {}))
        return self._checkers[key]

    def run(self) -> CheckResult:
        """执行全部检测并返回结构化结果。"""

        baseline = self.baseline
        findings = check_sections(baseline, self.document_context, self.mapping)
        result = CheckResult(findings=findings, skipped=skipped_reasons(), baseline=baseline)
        self._checkers = build_checkers(result)
        return result

    @staticmethod
    def report(result: CheckResult) -> list[tuple[str, RuleSpec, int]]:
        """生成“大类 → 规则 → 问题数”的报告行。"""

        return result.describe()


# ---------------------------------------------------------------------------
# 结果文案与批注写入
# ---------------------------------------------------------------------------

def _display(value) -> str:
    """把模板/实际值渲染成批注里可读的中文说明。"""

    if is_dataclass(value):
        value = asdict(value)
    if isinstance(value, dict):
        names = {"top": "上", "bottom": "下", "left": "左", "right": "右", "before": "段前",
                 "after": "段后", "points": "磅值", "characters": "字符数", "line_points": "行距磅值",
                 "line_rule": "行距规则"}
        return "；".join(f"{names.get(key, key)}：{item}" for key, item in value.items())
    return str(value)


def _legacy_message(finding: Finding) -> None:
    """把旧版流程的结论改写为与 docx 检测一致的批注文案（就地修改）。"""

    expected = (_display(finding.expected) if finding.expected is not None
                else "按模板对应用途的格式要求（差异见实际记录）")
    if finding.rule_id == "023":
        expected = "A4（本阶段保留的既有检查要求）"
    elif finding.rule_id == "072":
        expected = (_display(finding.expected) if finding.expected is not None
                    else "首行2字符，使用段落缩进（本阶段例外）")
    actual = _display(finding.actual) if finding.actual is not None else finding.message
    # 期望值/实际值同时保留在字段里，供统一批注渲染（normalize_finding）使用。
    finding.expected = expected
    finding.actual = actual
    finding.message = (f"【{finding.rule_id}｜{finding.title}】\n"
                       f"模板要求：{expected}\n"
                       f"当前文档实际：{actual}\n"
                       f"请核对调整。")


def _paragraph_anchor(context, finding: Finding):
    """为一条结论找到最合适的批注锚点段落。

    批注必须锚定在真正出问题的文字上：只有“段号命中的段落”或“按摘录文字找到的
    段落”核对得上时才返回锚点；核对不上时返回 None，由调用方把结论留在 JSON 报告里，
    而不是换一个位置（例如文首封面）写批注。

    同一条结论会在去重、合并、写批注三个阶段各问一次锚点；而每次都要按摘录文字
    扫描全篇一千多个段落。这里把结果记在结论对象上（``_anchor_done``），
    **同一次检测内只解析一次**——实测这一项单独占 60 秒以上。
    """

    if getattr(finding, "_anchor_done", False):
        return getattr(finding, "_anchor_resolved", None)
    anchor = _resolve_paragraph_anchor(context, finding)
    finding._anchor_done = True
    finding._anchor_resolved = anchor
    return anchor


def _resolve_paragraph_anchor(context, finding: Finding):
    """解析锚点段落的实际逻辑（见 :func:`_paragraph_anchor`）。"""

    explicit = getattr(finding, "anchor", None)
    if explicit is not None and context.index(explicit) >= 0:
        return explicit
    if finding.paragraph_index is not None:
        index = finding.paragraph_index - 1
        if 0 <= index < len(context.paragraphs):
            candidate = context.paragraphs[index]
            if not finding.excerpt or finding.excerpt in candidate.text:
                return candidate
        if finding.excerpt:
            matches = [paragraph for paragraph in context.paragraphs if finding.excerpt in paragraph.text]
            if matches:
                if 0 <= index < len(context.paragraphs):
                    # Word 流程的段号包含文本框等段落，与 docx 段号存在偏移：
                    # 同一段文字可能多处出现，取与记录段号最接近的一处。
                    return min(matches, key=lambda paragraph: abs(context.index(paragraph) - index))
                if len(matches) == 1:
                    return matches[0]
        # 段号与摘录文字都对不上：不猜位置，本条只进报告。
        finding.location_note = "段号与摘录文字都无法在当前文档中核对，本条未写批注"
    if finding.section_index is not None:
        return next((paragraph for paragraph in context.paragraphs
                     if context.section_index(paragraph) == finding.section_index - 1 and paragraph.text.strip()),
                    None)
    return None


#: 同一段落内不同位置可能是不同问题的规则（字符级/标点级）：去重时必须带文字范围。
_POSITION_RULES = frozenset({
    "103", "104", "124", "125", "126", "127", "130", "131", "132",
})


def dedupe_findings(findings: list[Finding], context) -> tuple[list[Finding], int]:
    """同一段落上的同一条问题只保留一条。

    去重按**内容**（规则编号 + 等级状态 + 锚点段落 + 应该/实际 + 文字范围）而不是
    文案措辞：旧版 Word 流程与新取证模块对同一问题会给出不同措辞，只按文案去重会漏掉。
    文字范围只对字符级/标点级规则生效，其余规则按“段落 + 应该/实际”即视为同一问题。
    """

    result: list[Finding] = []
    seen = set()
    dropped = 0
    for finding in findings:
        anchor = _paragraph_anchor(context, finding)
        index = context.index(anchor) if anchor is not None else -1
        if finding.page_story:
            key = (finding.rule_id, finding.status, finding.section_index,
                   finding.page_story, finding.page_scope, finding.page_variant,
                   finding.header_kind, str(finding.expected), str(finding.actual))
        elif index >= 0:
            # 字号/字体这类“整段属性”规则不再带文字范围：一段文字被按脚本切成多个
            # 片段时它们是同一个问题，重复写批注会让当事人看到“同一句话说了好几遍”。
            key = [finding.rule_id, finding.status, index, str(finding.expected), str(finding.actual)]
            if finding.rule_id.split(".")[0] in _POSITION_RULES:
                key.extend([finding.start, finding.end])
            key = tuple(key)
        else:
            # 定位不到段落时不能按“段落”合并，改用更细的字段避免误合并。
            key = (finding.rule_id, finding.status, finding.paragraph_index, finding.section_index,
                   finding.header_kind, str(finding.expected), str(finding.actual),
                   " ".join((finding.message or "").split()))
        if key in seen:
            dropped += 1
            continue
        seen.add(key)
        result.append(finding)
    return result, dropped


def merge_repeated_problems(findings: list[Finding], context, minimum: int = 3
                            ) -> tuple[list[Finding], int]:
    """同一规则、同一“应该/实际”的问题跨段落重复时，合并成一条汇总批注。

    例如“54 处三级标题中文字体应为黑体、实际宋体”合成一条，位置列出全部段落；
    少于 ``minimum`` 处的不合并，保持逐条批注的精确性。
    """

    grouped: dict[tuple, list[tuple[int, Finding]]] = {}
    for finding in findings:
        if finding.status != Status.FAIL or finding.merged_count or finding.page_story:
            continue
        anchor = _paragraph_anchor(context, finding)
        index = context.index(anchor) if anchor is not None else -1
        if index < 0:
            continue  # 定位不到具体段落的问题不参与合并
        grouped.setdefault((finding.rule_id, str(finding.expected), str(finding.actual)),
                           []).append((index, finding))
    dropped: set[int] = set()
    merged = 0
    for items in grouped.values():
        if len(items) < minimum:
            continue
        items.sort(key=lambda pair: pair[0])
        head = items[0][1]
        # 同一段可能产生多条同规则结论（例如同一段落被切成多个字号片段）：
        # 位置列表去重，避免批注里出现“第 517、517 段”这样的重复位置。
        head.merged_locations = sorted({index + 1 for index, _ in items})
        head.merged_count = len(head.merged_locations)
        for _, extra in items[1:]:
            dropped.add(id(extra))
        merged += len(items) - 1
    if not dropped:
        return findings, 0
    return [finding for finding in findings if id(finding) not in dropped], merged


def _location_text(finding: Finding, context, anchor) -> str:
    """生成“位置：”一行：优先用段号/节号，其次用批注锚点所在段号。"""

    if finding.merged_locations:
        shown = "、".join(str(index) for index in finding.merged_locations[:8])
        if len(finding.merged_locations) > 8:
            shown += " 等"
        return f"第 {shown} 段（共 {finding.merged_count} 处同类问题）"
    if finding.section_index is not None:
        base = f"第 {finding.section_index} 节"
    elif finding.paragraph_index is not None:
        base = f"第 {finding.paragraph_index} 段"
    else:
        base = ""
    if anchor is not None:
        index = context.index(anchor)
        if index >= 0:
            paragraph = f"第 {index + 1} 段"
            base = paragraph if not base or base == paragraph else f"{base}（{paragraph}）"
    note = getattr(finding, "location_note", "")
    if note:
        base = f"{base}；{note}" if base else note
    return base


def normalize_finding(finding: Finding, context, anchor=None) -> None:
    """统一批注的等级与文案：``等级 + 应该 + 实际``。

    全项目只有这一处渲染批注正文，因此等级、文案格式不会出现两套写法。
    批注本身已经锚定在出问题的文字上，不再在正文里重复写“位置：第 N 段”。
    定位信息仍保留在结构化结果（``finding.location``）里，供 JSON 报告与人工核对。
    """

    finding.severity = severity_of(finding.rule_id)
    location = (_location_text(finding, context, anchor) or finding.location
                or "未定位到具体段落")
    finding.location = location
    expected = _display(finding.expected) if finding.expected is not None else "按模板对应要求"
    actual = _display(finding.actual) if finding.actual is not None else finding.message
    # 普通批注已锚定在出问题的文字上，不再重复写“位置：”一行；只有可选的汇总批注
    # （显式传 merge_repeats>0 时生成）需要位置，否则看不出它覆盖哪些段落。
    summary_location = location if getattr(finding, "merged_locations", None) else ""
    finding.message = build_message(finding.severity, finding.rule_id, finding.title,
                                    expected, actual, status=finding.status,
                                    location=summary_location)


def semantic_tree(element):
    """只忽略批注标记的文档语义树，用于校验“除批注外没有任何改动”。"""

    tree = deepcopy(element)
    for node in list(tree.iter()):
        if node.tag in (qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")):
            parent = node.getparent()
            parent.remove(node)
            if parent.tag == qn("w:r") and not any(child.tag != qn("w:rPr") for child in parent):
                parent.getparent().remove(parent)

    def shape(node):
        return (node.tag, tuple(sorted(node.attrib.items())), node.text or "", node.tail or "",
                tuple(shape(child) for child in node))

    return shape(tree)


def write_page_comments(summary: RunSummary, context, source: Path, document, findings: list) -> tuple:
    """按真实页面归属写零长度边界批注；同一共享配置只写一次，其余页面记录在报告。"""
    from thesis_format_checker.core.comments import comment_allowed
    rejected = [f for f in findings if not comment_allowed(f)]
    findings = [f for f in findings if comment_allowed(f)]

    if not findings:
        return [], rejected
    if not getattr(context, "render_facts", None):
        return [], rejected + list(findings)
    # 排版几何已经顺带算出按页锚点时不重复打开 Word（省一次 3–4 分钟的遍历）。
    pagination = (getattr(context, "render_layout", None) or {}).get("pagination")
    if pagination is None or not pagination.boundaries_verified:
        try:
            from thesis_format_checker.core.word_layout import read_pagination

            pagination = read_pagination(source, context)
        except Exception as exc:
            summary.unknowns.append(
                f"页眉页脚批注未写入：排版快照读取失败（{type(exc).__name__}）；这些结论见报告")
            return [], rejected + list(findings)
    from thesis_format_checker.core.point_comments import add_boundary_comment

    written = []
    unplaced = list(rejected)
    grouped = {}
    for finding in findings:
        try:
            section = int(finding.section_index) - 1
        except (TypeError, ValueError):
            unplaced.append(finding)
            continue
        pages = pagination.page_of_section(section)
        if not pages:
            summary.unknowns.append(
                f"【{finding.rule_id}】第 {section + 1} 节没有可定位的页面，未写页眉页脚批注")
            unplaced.append(finding)
            continue
        story = finding.page_story or "header"
        variant = finding.page_variant or {1: "default", 2: "first", 3: "even"}.get(finding.header_kind)
        if finding.page_scope == "even_pages":
            first_page = pages[0].number
            pages = [p for p in pages if p.adjusted_number is not None
                     and p.adjusted_number % 2 == 0
                     and getattr(p, story + "_variant") is not None
                     and getattr(p, story + "_variant") != "first"
                     and not (finding.page_skip_first and p.number == first_page)]
        elif finding.page_scope == "first_page":
            pages = pages[:1]
        elif variant:
            pages = [p for p in pages if getattr(p, story + "_variant") == variant]
        elif finding.page_scope == "variant":
            pages = []   # 缺变体证据时不能按“普通页”猜测
        finding.affected_pages = [p.number for p in pages]
        if not pages:
            # The section exists, but this header/footer variant is not used by
            # any physical page in it. A dormant definition is not an affected
            # page and must not become a report-only error or a misplaced note.
            finding.status = Status.NOT_APPLICABLE
            finding.location_note = "该页眉页脚变体未在本节任何实际页面启用"
            summary.not_applicable.append(
                f"【{finding.rule_id}】{finding.actual}（该变体未在本节任何实际页面启用）")
            continue
        # 奇偶页不同是 document/settings.xml 的全局配置；同值差异不逐节重复批注。
        key = (finding.rule_id, story, finding.page_scope, variant,
               None if finding.page_scope == "even_pages" else section,
               str(finding.expected), str(finding.actual))
        if finding.page_config:
            key = ("configuration", story, finding.page_config)
        elif finding.page_issue:
            key = ("page_property", section, story, finding.page_issue)
        elif (finding.rule_id in {"035", "022.1", "029"} and story == "header"
              and ("无内容" in str(finding.expected) or "没有页眉" in str(finding.expected))):
            key = ("unexpected_header_content", section, story, variant)
        grouped.setdefault(key, []).append((finding, pages))
    # Visible content/distance differences are annotated on every affected
    # page. A document-wide switch remains one configuration comment.
    expanded = []
    missing_pages = {}
    for group in grouped.values():
        if group[0][0].page_config:
            expanded.append(group)
        else:
            by_page = {}
            for item, pages in group:
                for page in pages:
                    by_page.setdefault(page.number, []).append((item, [page]))
            expanded.extend(by_page.values())
    for group in expanded:
        finding = group[0][0]
        story = finding.page_story or "header"
        pages = {p.number: p for _, items in group for p in items}
        affected = sorted(pages)
        candidates = [(p, p.top if story == "header" else p.bottom)
                      for p in sorted(pages.values(), key=lambda p: p.number)]
        placed = next(((p, a) for p, a in candidates if a is not None), None)
        if placed is None:
            for item, _ in group:
                missing_pages.setdefault(id(item), set()).update(affected)
                unplaced.append(item)
            continue
        page, anchor = placed
        normalize_finding(finding, context, None)
        try:
            comment = add_boundary_comment(document, context.paragraphs[anchor.paragraph],
                                           anchor.offset, finding.message, edge=anchor.edge)
            if comment is None:
                raise WordAutomationError("批注文案不符合写入条件")
        except Exception as exc:
            summary.unknowns.append(
                f"【{finding.rule_id}】第 {page.number} 页批注写入失败：{type(exc).__name__}")
            unplaced.extend(item for item, _ in group)
            for item, _ in group:
                missing_pages.setdefault(id(item), set()).update(affected)
            continue
        for item, _ in group:
            item.location = f"{'页顶' if story == 'header' else '页底'}边界；受影响文件页序：" + "、".join(map(str, item.affected_pages))
        written.append({"page": page.number, "paragraph": anchor.paragraph + 1,
                        "offset": anchor.offset, "rule": finding.rule_id,
                        "story": story, "comment_id": comment.comment_id,
                        "position": anchor.position, "length": 0,
                        "edge": anchor.edge,
                        "affected_pages": affected, "variant": finding.page_variant,
                        "scope": finding.page_scope, "config": finding.page_config,
                        "rules": sorted({item.rule_id for item, _ in group})})
    if written:
        setattr(summary, "page_comments", written)
    unique_unplaced = list({id(f): f for f in unplaced}.values())
    for item in unique_unplaced:
        if id(item) in missing_pages:
            item.location_note = "部分受影响页面没有可安全定位的对应边界；未标注页：" + "、".join(
                map(str, sorted(missing_pages[id(item)])))
    return written, unique_unplaced


def save_comments_only(source: Path, document, output: Path, *,
                       allow_run_splits: bool = False, show_comments: bool = False) -> None:
    """逐件复制原包，只替换批注与锚点相关的部件。

    这是“输入原件不被修改、输出包除批注外逐字节一致”的保证：写盘前还会
    比较正文语义树，一旦发现批注以外的主文档改动就中止并保留原件。

    ``allow_run_splits``：页眉页脚问题的批注要锚在正文的**某一个字符**上（Word 不允许
    把批注写进页眉页脚），这需要按字符无损拆分 run；此时改用 :func:`content_tree`
    校验——它忽略批注并把相邻同格式 run 规范化，仍能证明“文字与格式没有变化”。
    ``show_comments``：在输出副本里打开“显示批注”开关，打开文件即可看到批注。
    """

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    comment_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    comment_rel = next((rel for rel in document.part.rels.values() if rel.reltype == comment_type), None)
    comment_name = str(comment_rel.target_part.partname).lstrip("/") if comment_rel else None
    allowed = {"word/document.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
    if comment_name:
        allowed.add(comment_name)
    with ZipFile(source) as original, ZipFile(buffer) as annotated:
        before = etree.fromstring(original.read("word/document.xml"))
        after = etree.fromstring(annotated.read("word/document.xml"))
        compare_tree = semantic_tree
        if allow_run_splits:
            from thesis_format_checker.core.point_comments import content_tree

            compare_tree = content_tree
        if compare_tree(before) != compare_tree(after):
            raise WordAutomationError("批注写入改变了非批注正文结构，已停止输出并保留原件。")
        replacements = {name: annotated.read(name) for name in allowed if name in annotated.namelist()}
        if show_comments and "word/settings.xml" in original.namelist():
            from thesis_format_checker.core.review_settings import visible_comments_settings

            replacements["word/settings.xml"] = visible_comments_settings(
                original.read("word/settings.xml"))
        # 关系与内容类型从原树构造，绝不重写无关元数据。
        for name, attr, value in (
            ("word/_rels/document.xml.rels", "Type", comment_type),
            ("[Content_Types].xml", "PartName", "/" + comment_name if comment_name else ""),
        ):
            if name in original.namelist():
                old = etree.fromstring(original.read(name))
            else:
                old = etree.Element("{http://schemas.openxmlformats.org/package/2006/relationships}Relationships")
            new = etree.fromstring(annotated.read(name))
            existing = [node for node in old if node.get(attr) == value]
            if not existing:
                added = next((node for node in new if node.get(attr) == value), None)
                if added is not None:
                    old.append(deepcopy(added))
            replacements[name] = etree.tostring(old, encoding="UTF-8", xml_declaration=True, standalone=True)
        output.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(suffix=".docx", prefix=".thesis-comments-", dir=output.parent, delete=False) as handle:
            temporary = Path(handle.name)
        try:
            with ZipFile(temporary, "w") as target:
                for info in original.infolist():
                    target.writestr(info, replacements.pop(info.filename, original.read(info.filename)))
                for name, blob in replacements.items():
                    target.writestr(name, blob)
            # 落盘前先验证原生批注可被 python-docx 读回，且锚点引用完整。
            verified = Document(temporary)
            ids = {str(comment.comment_id) for comment in verified.comments}
            references = verified.element.xpath(
                ".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference")
            if any(node.get(qn("w:id")) not in ids for node in references):
                raise WordAutomationError("批注锚点引用了不存在的批注，已停止输出并保留原件。")
            os.replace(temporary, output)
        finally:
            if temporary.exists():
                temporary.unlink()


# ---------------------------------------------------------------------------
# 对外流程
# ---------------------------------------------------------------------------

def run_comparison(legacy, template_path, document_path, output_path, **kwargs) -> RunSummary:
    """核心对照流程。

    Args:
        legacy: 旧版 Word 有效格式流程（``legacy(template, document, output, **kwargs)``），
            .doc 输入或需要 Word 解析有效格式时使用；测试中常注入替身。
        template_path / document_path / output_path: 模板 A、待检测文档 B、输出 .docx。
        **kwargs: ``mapping``、``overwrite`` 及各容差参数，原样透传给 Word 流程；
            ``merge_repeats`` / ``show_comments`` / ``use_cache`` / ``cache_dir``
            只属于本调度层，不会透传。

    Returns:
        检测摘要：问题列表、无法判定与跳过说明、输出的绝对路径。
    """

    template_path, document_path, output_path = map(Path, (template_path, document_path, output_path))
    # 合并阈值只属于本调度层，不能透传给旧版 Word 流程（它没有这个参数）。
    # 默认 0：批注逐处写出，不做跨段落汇总（需要汇总时显式传 merge_repeats=N）。
    merge_repeats = int(kwargs.pop("merge_repeats", 0) or 0)
    # “显示批注”开关同样只属于调度层（旧版 Word 流程没有这个参数）。
    show_comments = bool(kwargs.pop("show_comments", False))
    # 结果缓存（性能优化第 5 条）：默认只在真实 Word 流程下开启，
    # 测试注入的替身函数不走缓存，避免互相污染。
    use_cache = kwargs.pop("use_cache", None)
    cache_dir = kwargs.pop("cache_dir", None)
    if use_cache is None:
        use_cache = _real_word_flow(legacy)
    cache = ResultCache(directory=cache_dir, enabled=bool(use_cache))
    if template_path.suffix.lower() != ".docx" or document_path.suffix.lower() != ".docx":
        # .doc 输入仍走旧版 Word 流程，行为不变。
        return legacy(template_path, document_path, output_path, **kwargs)
    if output_path.resolve() in (template_path.resolve(), document_path.resolve()):
        raise WordAutomationError("输出不能覆盖模板或待检测文档原件。")
    if output_path.suffix.lower() != ".docx":
        raise WordAutomationError("输出文件扩展名必须是 .docx。")
    if output_path.exists() and not kwargs.get("overwrite", False):
        raise WordAutomationError(f"输出文件已存在：{output_path}。如需覆盖，请显式使用 --overwrite。")
    try:
        template = DocxContext(Document(template_path))
        output_doc = Document(document_path)
        context = DocxContext(output_doc)
    except Exception as exc:
        raise WordAutomationError(f"无法读取输入DOCX：{type(exc).__name__}；原件未修改。") from exc

    # 节对应关系自动推导：论文与模板的分节数常不同（封面/声明/摘要/目录可能合并），
    # 由“节角色”自动对应，避免让使用者手写 section_map；显式配置优先。
    mapping = dict(kwargs.get("mapping") or {})
    if not mapping.get("section_map"):
        from thesis_format_checker.detectors.pages import auto_section_map

        auto_map = auto_section_map(template, context)
        if auto_map:
            mapping["section_map"] = auto_map
    if not mapping.get("header_section_map"):
        from thesis_format_checker.detectors.pages import auto_header_section_map

        header_map = auto_header_section_map(template, context)
        if header_map:
            mapping["header_section_map"] = header_map
    kwargs["mapping"] = mapping

    summary = RunSummary(output_path=str(output_path.resolve()))
    started = time.perf_counter()
    timings: dict[str, float] = {}
    output_path.parent.mkdir(parents=True, exist_ok=True)
    findings: list[Finding] = []
    work = TemporaryDirectory(prefix=".thesis-compare-", dir=output_path.parent, ignore_cleanup_errors=True)
    # 排版几何先在 docx 侧按"文档里有什么对象"筛一遍，交给 Word 流程**同一次会话**
    # 顺手读出来（省掉为几何再开一次 Word + 重排）；批注统一由本层写入，
    # 因此让 Word 流程跳过会被丢弃的副本批注与保存。
    from thesis_format_checker.detectors import rendered as rendered_module

    layout_request = rendered_module.content_layout_request(context)
    prewarm: threading.Thread | None = None
    if _real_word_flow(legacy):
        kwargs["document_context"] = context
        kwargs["layout_request"] = layout_request
        kwargs["write_comments"] = False
        kwargs["word_cache"] = cache
        # 并行预热：Word 阶段以等待 COM 为主，正好把 docx 侧的全篇扫描提前算完。
        prewarm = threading.Thread(target=_prewarm_docx_scans, args=(template, context),
                                   name="docx-prewarm", daemon=True)
        prewarm.start()
    # 缓存键：模板与论文的**内容**指纹 + 源码指纹 + 影响结论的参数。
    word_flow_key = cache.key(
        "word-flow", template=template_path, document=document_path,
        params=_cache_params(
            mapping,
            overwrite=bool(kwargs.get("overwrite", False)),
            margin_tolerance_mm=kwargs.get("margin_tolerance_mm"),
            font_size_tolerance_pt=kwargs.get("font_size_tolerance_pt"),
            indent_tolerance_pt=kwargs.get("indent_tolerance_pt"),
        ),
    )
    cached_flow = cache.load("word-flow", word_flow_key, "Word 段落属性与分页事实")
    try:
        if cached_flow is not None:
            # 命中缓存：整趟 Word 流程（模板基准 + 逐段段落属性 + 分页统计）都被跳过。
            summary = summary_from_payload(cached_flow)
            summary.output_path = str(output_path.resolve())
        else:
            try:
                summary = legacy(template_path, document_path, Path(work.name) / "legacy.docx", **kwargs)
            except Exception as exc:
                for rule in WORD_EFFECTIVE_RULE_IDS:
                    findings.append(unknown(rule, "既有检测暂未完成",
                                            f"原Word检测不可用：{type(exc).__name__}；本条未执行，不计为通过"))
                summary.unknowns.append(
                    f"既有Word检测未完成：{type(exc).__name__}；仅输出已完成的新规则及限制批注")
            else:
                # 只缓存成功完成的 Word 流程结论；失败结果不缓存，下次仍会重试。
                cache.store("word-flow", word_flow_key, summary_payload(summary),
                            "Word 段落属性与分页事实")
    finally:
        try:
            work.cleanup()
        except OSError:
            # 清理失败（文件占用/权限）不能吞掉已完成的检测结果。
            summary.unknowns.append(f"临时检测副本未能清理，请关闭文件占用后处理：{work.name}")
        if prewarm is not None:
            # 预热只写缓存、不产生结论；即便超时也只会让后面的检测多算一次，结果不变。
            prewarm.join(timeout=300)
    timings["Word段落属性与分页"] = round(time.perf_counter() - started, 1)
    # Word 分页事实先挂到 docx 上下文：需要排版结果的规则（054/030/040）据此给出
    # 确定结论，而不是一律“无法判定”。读取失败时字典为空，行为与原来一致。
    context.render_facts = dict(getattr(summary, "render_facts", None) or {})
    engine = CheckEngine(template, context, mapping)
    baseline = engine.baseline
    static_started = time.perf_counter()
    findings.extend(check_sections(baseline, context, mapping))
    timings["docx静态检测"] = round(time.perf_counter() - static_started, 1)
    # ------------------------------------------------------------------
    # 静态先筛（性能优化第 3 条）+ 排版几何复核
    #
    # 只有真正跑过 Word 流程（拿到分页事实）才继续用排版事实复核：单测与 docx
    # 单跑不应因为这条去启动 Word。复核分两层，先便宜后昂贵：
    #   1) 只用分页事实的规则（032 空白页 / 165.1 文末空白页）——不读几何；
    #   2) 需要坐标的规则（078 题注位置 / 087 跨页表格 / 033 越界 / 098 公式
    #      越界 / 059 目录页码）——由 :func:`layout_requirements` 静态先筛出
    #      “真的会用到哪些段落”，只对它们调 Word；筛不出需求就不读排版。
    # ------------------------------------------------------------------
    if getattr(summary, "render_facts", None):
        from thesis_format_checker.detectors import rendered as rendered_module

        fact_families = rendered_module.FACT_ONLY_FAMILIES
        if any(finding.status == Status.UNKNOWN
               and finding.rule_id.split(".")[0] in fact_families for finding in findings):
            facts_findings = rendered_module.check_fact_errors(baseline, context)
            findings = [finding for finding in findings
                        if not (finding.status == Status.UNKNOWN
                                and finding.rule_id.split(".")[0] in fact_families)]
            findings.extend(facts_findings)
        request = rendered_module.layout_requirements(context, findings)
        geo_families = rendered_module.GEO_LAYOUT_FAMILIES
        layout_key = cache.key("layout", document=document_path)
        # ① Word 会话里已经顺手读出的几何（正常流程走这条：不再第二次打开 Word）；
        # ② 否则看缓存；③ 再否则按需单独打开 Word 读一次（测试替身/无预读时）。
        preview = getattr(summary, "render_layout", None) or {}
        layout = preview or {}
        if preview:
            if cache.enabled:
                cache.store("layout", layout_key, preview, "排版几何快照")
        else:
            layout = cache.load("layout", layout_key, "排版几何快照") or {}
        if not layout and request is not None:
            try:
                from thesis_format_checker.core.word_layout import read_layout

                layout = read_layout(document_path, context,
                                     paragraph_indices=request["indices"],
                                     want_figures=request["want_figures"],
                                     want_tables=request["want_tables"])
            except Exception as exc:
                layout = {}
                summary.unknowns.append(
                    f"排版几何读取失败：{type(exc).__name__}；相关规则保持“无法判定”")
            else:
                if layout:
                    cache.store("layout", layout_key, layout, "排版几何快照")
        if layout:
            context.render_layout = layout
            try:
                rendered = rendered_module.check_layout_errors(baseline, context, layout)
            except Exception as exc:
                rendered = []
                summary.unknowns.append(f"排版几何复核失败：{type(exc).__name__}")
            findings = [finding for finding in findings
                        if not (finding.status == Status.UNKNOWN
                                and finding.rule_id.split(".")[0] in geo_families)]
            findings.extend(rendered)
    for finding in summary.findings:
        _legacy_message(finding)
    summary.findings = adapt_legacy_findings(summary.findings, baseline, context, mapping=kwargs.get("mapping"))
    # 旧版 unknowns 文本也进入结构化报告，所有未知结论均不写批注。
    findings.extend(legacy_unknown_findings(summary.unknowns))
    summary.findings.extend(findings)
    summary.output_path = str(output_path.resolve())
    summary.skipped.extend(skipped_reasons())
    # 先做“同一段落同一问题”去重，再把跨段落重复的同类问题合并成汇总批注；
    # 报告与批注使用同一份结论列表，避免“报告 2441 条、文档只有 1643 条”的不一致。
    summary.findings, summary.duplicates_removed = dedupe_findings(summary.findings, context)
    if merge_repeats > 1:
        summary.findings, summary.merged_repeats = merge_repeated_problems(
            summary.findings, context, minimum=merge_repeats)
    page_scoped: list[Finding] = []
    annotate_started = time.perf_counter()
    for finding in summary.findings:
        anchor = _paragraph_anchor(context, finding)
        if finding.status == Status.NOT_APPLICABLE:
            # 模板没有对应内容：直接跳过，不写批注、不计入“无法判定”。
            summary.not_applicable.append(f"【{finding.rule_id}】{finding.actual}")
            continue
        if finding.status == Status.UNKNOWN:
            summary.unknowns.append(f"【{finding.rule_id}】{finding.actual}")
        from thesis_format_checker.core.comments import comment_allowed
        if not comment_allowed(finding):
            normalize_finding(finding, context, anchor)
            summary.report_only.append(f"【{finding.rule_id}】{finding.actual}")
            continue
        if finding.page_story:
            normalize_finding(finding, context, None)
            page_scoped.append(finding)
            continue
        if anchor is None:
            # 定位不到具体段落时不动位置：批注只写在确实有问题的位置上，
            # 这条结论留在 JSON 报告里由人工核对。
            normalize_finding(finding, context, anchor)
            summary.report_only.append(
                f"【{finding.rule_id}】{finding.actual}（未能定位到具体段落，未写批注）")
            continue
        # 统一批注格式：等级 + 应该 + 实际 + 位置。
        normalize_finding(finding, context, anchor)
        from thesis_format_checker.core.comments import add_word_comment, paragraph_comment_limitation
        limitation = paragraph_comment_limitation(anchor)
        if limitation:
            summary.report_only.append(f"【{finding.rule_id}】{finding.actual}（{limitation}）")
            continue
        try:
            add_word_comment(output_doc, finding, anchor)
        except WordAutomationError as exc:
            summary.report_only.append(f"【{finding.rule_id}】{finding.actual}（{exc}，未写批注）")
    page_written, page_unplaced = write_page_comments(
        summary, context, document_path, output_doc, page_scoped)
    # 定位失败只进报告，严禁退回节首、文首或无关正文。
    for finding in page_unplaced:
        normalize_finding(finding, context, None)
        summary.report_only.append(
            f"【{finding.rule_id}】{finding.actual}（未能确认对应页面边界，未写批注）")
    timings["批注锚定与写入"] = round(time.perf_counter() - annotate_started, 1)
    # ``show_comments`` 默认关闭：输出包除批注外逐字节保持不变（README/架构说明的承诺）；
    # 需要“打开就显示批注”时用 ``--show-comments`` 显式打开。
    comment_started = time.perf_counter()
    save_comments_only(document_path, output_doc, output_path,
                       allow_run_splits=True, show_comments=show_comments)
    timings["输出校验与落盘"] = round(time.perf_counter() - comment_started, 1)
    # 结果缓存统计只用于向使用者解释“这次为什么快/慢”，不参与任何判定。
    summary.cache_stats = cache.stats()
    timings["总耗时"] = round(time.perf_counter() - started, 1)
    summary.timings = timings
    return summary


def compare_documents(template_path: Path, document_path: Path, output_path: Path,
                      mapping: dict | None = None, overwrite: bool = False,
                      margin_tolerance_mm: float = 0.2, font_size_tolerance_pt: float = 0.1,
                      indent_tolerance_pt: float = 0.5, merge_repeats: int = 0,
                      show_comments: bool = False, use_cache: bool | None = None,
                      cache_dir: Path | None = None) -> RunSummary:
    """公共入口：以模板 A 对照检测文档 B，并把问题写成 Word 原生批注。

    ``.doc`` 输入由本机 Word 转换后检测；输出统一为 ``.docx``；输入原件
    永不被修改。``use_cache`` 控制结果缓存（默认开启：同一组模板/论文重复检测
    时复用 Word 段落属性与排版快照）；``cache_dir`` 可指定缓存目录。
    """

    from thesis_format_checker.detectors.word_effective.runner import compare_word_documents

    return run_comparison(
        compare_word_documents, template_path, document_path, output_path,
        mapping=mapping, overwrite=overwrite, margin_tolerance_mm=margin_tolerance_mm,
        font_size_tolerance_pt=font_size_tolerance_pt, indent_tolerance_pt=indent_tolerance_pt,
        merge_repeats=merge_repeats, show_comments=show_comments,
        use_cache=use_cache, cache_dir=cache_dir,
    )


__all__ = [
    "ALL_RULE_IDS",
    "CheckEngine",
    "RULES",
    "SECTION_RULES",
    "WORD_EFFECTIVE_RULE_IDS",
    "adapt_legacy_findings",
    "check_sections",
    "compare_documents",
    "extract_baseline",
    "legacy_unknown_findings",
    "normalize_finding",
    "run_comparison",
    "save_comments_only",
    "semantic_tree",
    "skipped_reasons",
]
