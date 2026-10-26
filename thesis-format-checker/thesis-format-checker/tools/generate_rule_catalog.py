"""由当前规则目录生成交付版 docs/规则清单.md。

运行：``python tools/generate_rule_catalog.py``。规则状态只从
``checks/rule_specs.py``、``checks/catalog.py`` 和 ``PROVIDERS`` 读取。
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from thesis_format_checker.checks import catalog  # noqa: E402
from thesis_format_checker.core.models import SEVERITY_LABELS  # noqa: E402

OUTPUT = ROOT / "docs" / "规则清单.md"

TYPES = {
    "outline": "结构识别", "structure": "结构与顺序比对",
    "front": "封面/前置字段比对", "cover_media": "图片与版式几何比对",
    "pages": "节属性与分页比对", "headers": "页眉页脚与页码比对",
    "headings": "标题层级与自动编号比对", "toc": "TOC 域与目录条目比对",
    "paragraphs": "段落格式与分页属性比对", "figures": "图表、题注与表格结构比对",
    "equations": "公式对象、编号与格式比对", "notes": "引用、注释与域比对",
    "bibliography": "参考文献结构化著录比对", "characters": "字符与标点规则",
    "consistency": "跨区域关联一致性比对", "formats": "同类对象格式一致性比对",
    "appendices": "附录结构与编号比对", "hygiene": "残留内容与文末检查",
    "language_lite": "确定性词表与写法规则", "legacy": "Word COM 有效格式比对",
}

PLANS = {
    "028": "先定义允许变化的节类型或白名单，才能判断页面差异是否必要。",
    "040": "需用 Word 打印页码建立逐页序列，并处理分节重启、域结果和空白页。",
    "056": "需比较 Word 更新后的 TOC 域结果、域代码和正文重新生成结果。",
    "057": "已由 136 覆盖；若单独输出，应复用目录—正文映射并去重。",
    "058": "已由 136 覆盖；若拆分输出，应复用同一映射并按属性去重。",
    "063": "需读取目录页码的 Word 渲染右边界，不能用段落整体对齐代替。",
    "065": "需结合 Word 分页坐标并先定义孤立条目的页首/页尾阈值。",
    "069": "标题字号已由 045.3、046.3、047.3 输出，本条仅作兼容去重。",
    "080": "需在现有编号序列上增加重号定位，并与跳号、章节号统一去重。",
    "082": "已由 137.1 覆盖；若恢复，应复用图表引用索引并避免重复批注。",
    "085": "超出版心已由 033.1 覆盖；尺寸过小或比例失真需量化阈值。",
    "088": "需先按模板或配置定义题注、来源和单位的适用条件。",
    "089": "需可靠识别本应为公式的普通文本或图片。",
    "093": "已由 137.3 覆盖；若恢复，应复用公式引用索引。",
    "094": "变量、常量、函数、向量和矩阵的正斜体依赖数学语义。",
    "095": "上下标层级是否正确依赖数学表达语义。",
    "096": "需增加公式内部字符级字体解析并区分数学字体和正文文字。",
    "106": "需由 Word 读取脚注显示编号并处理每节/每页重新编号。",
    "107": "需从脚注故事范围读取有效字号并与模板脚注样例比较。",
    "109": "已由 137.2 覆盖；若恢复，应复用域目标索引并去重。",
    "120": "需归一化正文引用和文献条目，建立双向映射。",
    "127": "英文单词间是否缺空格需要词典或语义识别合法连写。",
    "141": "与 120 职责重叠，启用前应统一正文引用—文献表口径。",
    "161": "需学校先定义成果字段、排序键和作者标注规则。",
}


def _providers(family):
    return [catalog.provider_by_key(key) for key in family.providers]


def _kind(family) -> str:
    if not family.providers:
        return "业务/语义口径待定"
    return " + ".join(dict.fromkeys(TYPES.get(key, key) for key in family.providers))


def _implementation(item, family) -> str:
    providers = _providers(family)
    evidence = "；".join(dict.fromkeys(p.label for p in providers))
    if item.status is catalog.RuleStatus.IMPLEMENTED:
        return f"{evidence}；按模板同用途对象建立基准，问题定位到{item.location}。"
    if item.status is catalog.RuleStatus.AGGREGATED:
        return (f"{evidence}；已能检出并定位到{item.location}，但当前与同族属性合并输出；"
                "后续可拆为独立结论。")
    plan = PLANS.get(item.family) or family.note or "业务口径尚未确定。"
    return f"未实现。{plan}"


def _summary() -> list[str]:
    items = list(catalog.FINE_RULES.values())
    statuses = Counter(item.status for item in items)
    severity = Counter(item.severity.value for item in items)
    return [
        "# 规则清单（交付版）", "",
        "本清单由 `checks/rule_specs.py`、`checks/catalog.py` 和 `catalog.PROVIDERS` 自动生成。",
        "生成命令：`python tools/generate_rule_catalog.py`。实现状态以源码和架构自检为准。", "",
        "## 一、统计与分类口径", "",
        f"* 检测大类：{len(catalog.CATEGORIES)} 个。",
        f"* 规则族：{len(catalog.ALL_FAMILY_IDS)} 个（保留 {len(catalog.RULES)} 个，"
        f"按范围删除语义规则族 {len(catalog.REMOVED_RULES)} 个）。",
        f"* 细粒度规则：{len(items)} 条；错误 {severity['error']} 条、"
        f"警告 {severity['warning']} 条、提醒 {severity['hint']} 条。",
        f"* 完全实现：{statuses[catalog.RuleStatus.IMPLEMENTED]} 条。",
        f"* 已实现但有优化空间：{statuses[catalog.RuleStatus.AGGREGATED]} 条。",
        f"* 未实现：{statuses[catalog.RuleStatus.PENDING] + statuses[catalog.RuleStatus.DEFERRED]} 条；"
        f"另有 {len(catalog.REMOVED_RULES)} 个开放式语义规则族暂不实现。", "",
        "| 分类 | 判定标准 | 交付含义 |",
        "| --- | --- | --- |",
        "| 完全实现 | 有确定性取证、独立结论和对应定位 | 可用于生产检测 |",
        "| 已实现但有优化空间 | 已能检测，但同族属性仍聚合输出 | 可使用，后续可继续拆分批注粒度 |",
        "| 未实现 | 口径待定、与其他规则重叠，或需语义/更深 Word 解析 | 当前不产出确定结论 |", "",
        "运行时的“无法判定”不等于规则未实现。模板无基准、对象不可读或 Word 取证失败时，",
        "原因只进入 JSON 报告，不写入 Word 批注。", "",
    ]


def _status_table(title: str, wanted: set[catalog.RuleStatus]) -> list[str]:
    total = sum(item.status in wanted for item in catalog.FINE_RULES.values())
    lines = [f"## {title}（{total} 条）", ""]
    for category in catalog.CATEGORIES:
        items = [item for item in catalog.fine_rules_of_category(category.key)
                 if item.status in wanted]
        if not items:
            continue
        lines += [f"### {category.key} {category.name}（{len(items)} 条）", "",
                  "| 序号 | 检测规则 | 等级 | 检测类型 | 实现方案 |",
                  "| --- | --- | --- | --- | --- |"]
        for item in items:
            family = catalog.RULES[item.family]
            lines.append(f"| {item.code} | {item.title} | {SEVERITY_LABELS[item.severity]} | "
                         f"{_kind(family)} | {_implementation(item, family)} |")
        lines.append("")
    return lines


def _removed() -> list[str]:
    lines = ["## 五、按范围暂不实现的语义规则族", "",
             "以下原始规则族不计入 356 条确定性细粒度规则。它们需要理解论文内容、数学或图像语义，",
             "按当前交付范围暂不实现；编号保留用于追溯。", "",
             "| 序号 | 检测类型 | 未实现原因/后续方案 |", "| --- | --- | --- |"]
    for rule_id, reason in catalog.REMOVED_REASONS.items():
        lines.append(f"| {rule_id} | 开放式语义判断 | 当前不实现。{reason}；后续如启用，"
                     "需引入可审计的语义模型和人工复核机制。 |")
    lines += ["", "## 六、维护与验收", "",
              "1. 编号、状态、等级和定位方式只在 `checks/rule_specs.py` 维护。",
              "2. 取证归属在 `checks/catalog.py` 的 `PROVIDERS` 维护。",
              "3. 重新运行本生成器，不手工改统计数字。",
              "4. 运行 `python -m pytest tests/test_checks_architecture.py` 核对目录与实现声明。",
              "5. 交付前运行全量单元测试和真实 Word 集成测试。", ""]
    return lines


def build() -> str:
    lines = _summary()
    lines += _status_table("二、完全实现", {catalog.RuleStatus.IMPLEMENTED})
    lines += _status_table("三、已实现但有优化空间", {catalog.RuleStatus.AGGREGATED})
    lines += _status_table("四、未实现（不含已删除语义规则族）",
                           {catalog.RuleStatus.PENDING, catalog.RuleStatus.DEFERRED})
    lines += _removed()
    return "\n".join(lines)


def main() -> None:
    OUTPUT.write_text(build(), encoding="utf-8")
    print(f"已生成：{OUTPUT}")


if __name__ == "__main__":
    main()
