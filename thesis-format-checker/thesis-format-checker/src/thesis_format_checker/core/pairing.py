"""模板 ←→ 论文的“同一层级 + 同一位置”配对工具（纯函数，不依赖 Word/docx）。

口径
====

论文格式检测是**模板比对**：

* **同一层级**——按段落样式名、标题级别、题注种类等把两份文档各自分组；
* **同一位置**——组内按相对位置一一对应（组内数量相同即为逐个对应；
  模板样例更少时按比例落点，绝不越界）；
* 只有模板完全没有该层级的样例时，才需要退回“全篇主导值”或输出“无法判定”。

``detectors/`` 与 ``detectors/word_effective/`` 都复用本模块，避免两套口径。
"""

from __future__ import annotations

import re
from collections import defaultdict
from typing import Any, Iterable, TypeVar

T = TypeVar("T")


def style_key(name: Any) -> str:
    """层级键：去掉空格、统一大小写（“正文 ”与“正文”视为同一层级）。"""

    return re.sub(r"\s+", "", str(name or "")).casefold()


def partner_index(ordinal: int, sample_count: int, item_count: int) -> int:
    """把论文组内第 ``ordinal`` 个（0 起）按相对位置映射到模板组内的下标。

    * 组内数量相同 → 逐个对应（第 3 个对第 3 个）；
    * 模板样例更少 → 按比例落点，绝不越界；
    * 任一组为空 → ``-1``，由调用方决定兜底方式。
    """

    if sample_count <= 0 or item_count <= 0:
        return -1
    if sample_count == 1 or item_count == 1:
        return 0
    ratio = ordinal / (item_count - 1)
    return max(0, min(sample_count - 1, int(round(ratio * (sample_count - 1)))))


def sample_at(samples: "list[T] | tuple[T, ...]", ordinal: int, item_count: int) -> T | None:
    """按“同一层级 + 同一位置”取模板样例；没有样例时返回 ``None``。"""

    position = partner_index(ordinal, len(samples), item_count)
    return None if position < 0 else samples[position]


def dominant_value(weighted: Iterable[tuple[T | None, int]]) -> T | None:
    """取一组加权值里的主导值（并列时保留先出现的那个）。

    用于“同一段落内部”或“模板确实只有一个层级样例”时的兜底，
    **不**用于要求模板全文唯一。
    """

    totals: dict[T, int] = defaultdict(int)
    best: T | None = None
    best_weight = 0
    for value, weight in weighted:
        if value is None or weight <= 0:
            continue
        totals[value] += weight
        if totals[value] > best_weight:
            best, best_weight = value, totals[value]
    return best


def group_by(items: Iterable[T], key) -> dict[Any, list[T]]:
    """按 ``key`` 分组并保持原顺序。"""

    groups: dict[Any, list[T]] = defaultdict(list)
    for item in items:
        groups[key(item)].append(item)
    return groups
