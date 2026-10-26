from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


class Status(str, Enum):
    FAIL = "fail"
    UNKNOWN = "unknown"
    #: 模板里不存在该对象/样例，本条规则对这份论文不适用（不写批注，只在报告里统计）。
    NOT_APPLICABLE = "not_applicable"
    SKIPPED = "skipped"


class Severity(str, Enum):
    """问题等级：错误必须修改，警告建议修改，提醒可改可不改。"""

    ERROR = "error"
    WARNING = "warning"
    HINT = "hint"


#: 等级的中文名称，统一用于批注、终端摘要与报告。
SEVERITY_LABELS: dict[Severity, str] = {
    Severity.ERROR: "错误",
    Severity.WARNING: "警告",
    Severity.HINT: "提醒",
}

#: 等级排序权重（数值越大越严重），用于取“规则族内的最高等级”。
SEVERITY_ORDER: dict[Severity, int] = {
    Severity.HINT: 0,
    Severity.WARNING: 1,
    Severity.ERROR: 2,
}


def severity_label(severity: Severity, status: "Status" = Status.FAIL) -> str:
    """批注抬头里的等级文字：错误 / 警告 / 提醒，或“无法判定”“未实现”。"""

    if status is Status.UNKNOWN:
        return "无法判定"
    if status is Status.NOT_APPLICABLE:
        return "不适用"
    if status is Status.SKIPPED:
        return "未实现"
    return SEVERITY_LABELS[severity]


def build_comment(severity: Severity, rule_id: str, title: str, expected: Any, actual: Any,
                  status: "Status" = Status.FAIL, location: str = "") -> str:
    """统一的批注文案模板（全项目唯一实现）。

    ::

        【错误】封面题目格式不符合模板
        应该：字号 19 磅、加粗
        实际：字号 16 磅、未加粗

    批注面向普通用户：只写等级与问题标题，不写规则编号（编号在 JSON 报告里），
    也不写“位置：第 N 段”（批注本身已经贴在出问题的文字上）。
    """

    head = f"【{severity_label(severity, status)}】{title}"
    lines = [head, f"应该：{expected}", f"实际：{actual}"]
    if location:
        lines.append(f"位置：{location}")
    return "\n".join(lines)


@dataclass(frozen=True)
class MarginSpec:
    top: float
    bottom: float
    left: float
    right: float


@dataclass(frozen=True)
class IndentSpec:
    points: float
    characters: float | None = None


@dataclass(frozen=True)
class SpacingSpec:
    line_rule: int
    line_points: float
    before: float
    after: float


@dataclass(frozen=True)
class HeadingSpec:
    chinese_font: str | None
    latin_font: str | None
    size_pt: float | None
    bold: bool | None
    alignment: int | None
    indent: IndentSpec | None
    before: float | None
    after: float | None


@dataclass(frozen=True)
class BodySample:
    """模板中**某一个正文段落**的实际格式。

    论文格式检测是模板比对：论文第 i 个同类正文段落只需要和模板里
    同一层级（同一样式）、同一相对位置的段落比较，因此基准要保留到段落粒度，
    而不是先在全篇模板里求出“唯一标准”。
    """

    #: 模板中的 Word 段落序号（1 起），用于定位与调试。
    index: int
    #: 段落样式名（“同一层级”的判定键）。
    style_name: str
    chinese_font: str | None = None
    latin_font: str | None = None
    size_pt: float | None = None
    first_line_indent: IndentSpec | None = None
    spacing: SpacingSpec | None = None
    #: 该模板段落的字符数，用于合成“典型样例”时按篇幅加权。
    weight: int = 1


@dataclass(frozen=True)
class HeadingSample:
    """模板中**某一个标题**的层级与格式（同级标题按出现顺序对应）。"""

    index: int
    level: int
    spec: HeadingSpec


@dataclass
class TemplateProfile:
    chinese_font: str | None = None
    latin_font: str | None = None
    font_size_pt: float | None = None
    first_line_indent: IndentSpec | None = None
    body_spacing: SpacingSpec | None = None
    #: 逐段基准：``同层级 + 同位置``比对的依据（见 detectors/word_effective/correspondence.py）。
    body_samples: list[BodySample] = field(default_factory=list)
    #: 逐个标题基准：同级标题按出现顺序与论文对应。
    heading_samples: list[HeadingSample] = field(default_factory=list)
    #: 全模板完全一致时保留的标题基准；仅作为无法定位同级样例时的兜底。
    headings: dict[int, HeadingSpec] = field(default_factory=dict)
    headers: list[dict[int, str]] = field(default_factory=list)
    margins: list[MarginSpec] = field(default_factory=list)
    uniform_margin: MarginSpec | None = None
    body_style_names: list[str] = field(default_factory=list)
    unknown_reasons: list[str] = field(default_factory=list)


@dataclass
class Finding:
    rule_id: str
    title: str
    status: Status
    message: str
    #: 问题等级（错误 / 警告 / 提醒）；由检测调度层按规则目录统一填写。
    severity: Severity | None = None
    #: 精准定位说明，例如「第 12 段」「第 3 节」「页眉（普通页）」。
    location: str = ""
    expected: Any = None
    actual: Any = None
    paragraph_index: int | None = None
    section_index: int | None = None
    header_kind: int | None = None
    #: 页眉/页脚结论的承载位置：``header`` 或 ``footer``；配合 ``section_index``
    #: 由调度层把批注锚在该页版心顶部/底部的正文上（Word 不允许批注写进页眉页脚）。
    page_story: str = ""
    #: 页面范围来自取证元数据，不从批注文案反推；空变体表示节级问题。
    page_variant: str | None = None
    page_scope: str = "section"
    page_skip_first: bool = False
    #: 同一底层配置可由多个规则报告；仅批注按该键合并，保留全部规则结论。
    page_config: str = ""
    #: 同一页面上的同一属性差异跨规则去重；与全局配置键不同，仍逐页标注。
    page_issue: str = ""
    affected_pages: list[int] = field(default_factory=list)
    start: int | None = None
    end: int | None = None
    excerpt: str = ""
    #: 合并同类重复问题时的全部段落号（1 起，按出现顺序）；为空表示未合并。
    merged_locations: list[int] = field(default_factory=list)
    #: 合并了多少处（≥2 表示这是一条汇总批注）。
    merged_count: int = 0


@dataclass
class RunSummary:
    sections_checked: int = 0
    body_paragraphs_checked: int = 0
    findings: list[Finding] = field(default_factory=list)
    unknowns: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    #: 模板未涉及、因而跳过的规则说明（不写批注，也不计入“无法判定”）。
    not_applicable: list[str] = field(default_factory=list)
    #: 同一段落上重复出现的同一条问题被去掉的条数。
    duplicates_removed: int = 0
    #: 同一问题跨段落重复、合并进一条汇总批注的条数。
    merged_repeats: int = 0
    #: 只在报告里出现、不写进文档批注的“无法判定”条目（需要 Word 渲染或对象不可读）。
    report_only: list[str] = field(default_factory=list)
    #: Word 分页事实（段落页码、节页范围、总页数）；docx 侧规则据此给出确定结论。
    render_facts: dict = field(default_factory=dict)
    #: 结果缓存统计（命中/未命中/说明）；只在报告里展示，不影响结论。
    cache_stats: dict = field(default_factory=dict)
    #: 分段耗时（秒）：Word 阶段 / docx 静态检测 / 批注写入与校验 / 总耗时。
    timings: dict = field(default_factory=dict)
    #: 按页写入的页眉页脚批注明细（页码、锚点段落/字符偏移、规则、批注 id）。
    page_comments: list = field(default_factory=list)
    output_path: str = ""

    def counts_by_rule(self) -> dict[str, int]:
        """按规则族统计问题**处数**：合并成汇总批注的按 ``merged_count`` 计。"""

        result: dict[str, int] = {}
        for finding in self.findings:
            if finding.status == Status.FAIL:
                count = finding.merged_count or 1
                result[finding.rule_id] = result.get(finding.rule_id, 0) + count
        return result

    def counts_by_severity(self) -> dict[Severity, int]:
        """按等级统计问题数（未标注等级的历史结论按“错误”计）。"""

        result: dict[Severity, int] = {severity: 0 for severity in Severity}
        for finding in self.findings:
            if finding.status != Status.FAIL:
                continue
            severity = finding.severity or Severity.ERROR
            result[severity] = result.get(severity, 0) + 1
        return result
