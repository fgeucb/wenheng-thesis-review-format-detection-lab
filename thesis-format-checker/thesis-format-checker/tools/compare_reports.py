"""对照两份检测报告（``--report-json`` 产物），核对"优化前后结论一致"。

用法::

    python tools/compare_reports.py 旧报告.json 新报告.json

比对口径：把每条结论归一化成 ``(规则编号, 状态, 应该, 实际, 段号, 节号)`` 后做集合差；
**结论条数与逐条内容必须完全一致**，任何差异都会按规则族列出前若干条明细。
两份报告还可能是"同一份文档在不同代码版本下跑出来的"，因此本工具只比结论，
不比对耗时/缓存等运行信息。

退出码：0 = 完全一致；1 = 存在差异；2 = 参数/文件错误。
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path


def finding_key(item: dict) -> tuple:
    """结论的归一化键：只保留"判定内容"，忽略文案措辞与展示字段。"""

    return (str(item.get("rule_id")), str(item.get("status")), str(item.get("expected")),
            str(item.get("actual")), item.get("paragraph_index"), item.get("section_index"))


def describe(item: dict) -> str:
    return (f"[{item.get('rule_id')}] 应={item.get('expected')!r} 实={item.get('actual')!r} "
            f"段={item.get('paragraph_index')}")


def load(path: Path) -> list[dict]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    findings = payload.get("findings")
    if not isinstance(findings, list):
        raise ValueError(f"{path} 不是检测报告（缺少 findings）")
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description="对照两份检测报告的结论")
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path)
    parser.add_argument("--show", type=int, default=10, help="每个规则族最多列出的差异条数")
    args = parser.parse_args()
    try:
        before, after = load(args.before), load(args.after)
    except (OSError, ValueError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 2
    left, right = Counter(map(finding_key, before)), Counter(map(finding_key, after))
    removed = left - right
    added = right - left
    print(f"结论条数：{len(before)} → {len(after)}（差 {len(after) - len(before):+d}）")
    print(f"只出现在旧报告：{sum(removed.values())} 条；只出现在新报告：{sum(added.values())} 条")
    for label, bucket in (("旧 → 新（缺失）", removed), ("新 → 旧（新增）", added)):
        if not bucket:
            continue
        print(f"\n== {label} ==")
        by_family: dict[str, list[tuple]] = {}
        for key, count in bucket.items():
            by_family.setdefault(key[0].split(".")[0], []).extend([key] * count)
        for family, keys in sorted(by_family.items()):
            print(f"  规则 {family}：{len(keys)} 条")
            for key in keys[:args.show]:
                print("    -", key)
    return 0 if not (removed or added) else 1


if __name__ == "__main__":
    raise SystemExit(main())
