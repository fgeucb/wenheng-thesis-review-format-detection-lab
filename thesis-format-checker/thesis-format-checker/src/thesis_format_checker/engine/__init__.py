"""调度与报告层。

* :mod:`~thesis_format_checker.engine.pipeline` —— 模板基准、逐类取证、批注落盘；
* :mod:`~thesis_format_checker.engine.reporting` —— 命令行摘要、规则清单、JSON 报告。
"""

from thesis_format_checker.engine.pipeline import (
    SECTION_RULES,
    CheckEngine,
    check_sections,
    compare_documents,
    extract_baseline,
    run_comparison,
    save_comments_only,
    skipped_reasons,
)

__all__ = [
    "SECTION_RULES",
    "CheckEngine",
    "check_sections",
    "compare_documents",
    "extract_baseline",
    "run_comparison",
    "save_comments_only",
    "skipped_reasons",
]
