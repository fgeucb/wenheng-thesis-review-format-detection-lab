"""对全部错误样本执行 docx 静态取证，输出目标命中和规则分布供回归分析。"""

from __future__ import annotations

import json
import sys
import time
from collections import Counter
from pathlib import Path

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from thesis_format_checker.core.docx_context import DocxContext  # noqa: E402
from thesis_format_checker.core.models import Status  # noqa: E402
from thesis_format_checker.engine.pipeline import CheckEngine  # noqa: E402

SOURCE = ROOT / "samples" / "错误样本" / "错误样本"
OUTPUT = ROOT / "samples" / "错误样本" / "检测结果_20261014_修正版" / "静态预检汇总.json"
TEMPLATE = SOURCE / "南京理工大学硕士学位论文撰写格式.docx"

EXPECTED_BY_NAME = {
    "关键词缺失和数量不符": ["005"], "空章节和重复章节": ["011"],
    "目录非自动生成": ["006"], "缺少参考文献": ["008"],
    "缺少封二": ["001"], "缺少封一": ["001"], "缺少目录": ["006"],
    "缺少声明和授权声明": ["002"], "缺少英文摘要": ["004"],
    "缺少正文章节": ["007"], "缺少致谢和附录": ["009"],
    "缺少中文摘要": ["003"], "组成部分排列顺序错误": ["010"],
    "封面题目缺失": ["012"], "页码缺失重复跳号": ["040"],
}


def family(rule_id: str) -> str:
    return rule_id.split(".", 1)[0]


def expected(category_number: int, stem: str) -> list[str]:
    for marker, families in EXPECTED_BY_NAME.items():
        if marker in stem:
            return families
    # 其余综合样本只说明“本类别含若干故意错误”，没有逐条真值标注。
    # 不能把整类 11 个规则族都假定为必然命中，否则会制造虚假的“漏报”。
    return []


def main() -> None:
    baseline_context = DocxContext(Document(TEMPLATE))
    baseline = CheckEngine(baseline_context, baseline_context).baseline
    records = []
    for directory in sorted((p for p in SOURCE.iterdir() if p.is_dir()),
                            key=lambda p: int(p.name.split(".", 1)[0])):
        category_number = int(directory.name.split(".", 1)[0])
        for path in sorted(directory.glob("*.docx")):
            started = time.perf_counter()
            result = CheckEngine(baseline_context, DocxContext(Document(path)),
                                 baseline=baseline).run()
            elapsed = round(time.perf_counter() - started, 3)
            fails = [item for item in result.findings if item.status is Status.FAIL]
            unknowns = [item for item in result.findings if item.status is Status.UNKNOWN]
            fail_counts = Counter(family(item.rule_id) for item in fails)
            expected_families = expected(category_number, path.stem)
            category_start = (category_number - 1) * 11 + 1
            category_families = {
                f"{number:03d}" for number in range(category_start, category_start + 11)
            }
            records.append({
                "category": directory.name, "source": str(path), "seconds": elapsed,
                "expected_families": expected_families,
                "expectation_basis": "文件名明确真值" if expected_families else "无逐条真值，仅统计本类命中",
                "category_hit_families": sorted(category_families & set(fail_counts)),
                "hit_families": sorted(set(expected_families) & set(fail_counts)),
                "missed_families": sorted(set(expected_families) - set(fail_counts)),
                "fail_count": len(fails), "unknown_count": len(unknowns),
                "fail_counts_by_family": dict(sorted(fail_counts.items())),
                "target_findings": [
                    {"rule_id": item.rule_id, "title": item.title, "actual": str(item.actual),
                     "location": item.location, "excerpt": item.excerpt}
                    for item in fails if family(item.rule_id) in expected_families
                ],
            })
            print(f"{directory.name} / {path.name}: {elapsed:.1f}s，"
                  f"命中={records[-1]['hit_families']}，缺失={records[-1]['missed_families']}",
                  flush=True)
    common = Counter()
    for record in records:
        common.update(record["fail_counts_by_family"].keys())
    payload = {"template": str(TEMPLATE), "records": records,
               "families_by_sample_count": dict(common.most_common())}
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已写入：{OUTPUT}")


if __name__ == "__main__":
    main()
