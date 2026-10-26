from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from uuid import uuid4

import pytest

from thesis_format_checker.checks.catalog import FINE_RULES, RULES


def comment_heads(comments):
    """批注抬头里的标题集合（批注不再写规则编号，测试按标题核对）。"""

    heads = set()
    for comment in comments:
        first = (comment.text or "").splitlines()[0]
        if "】" in first:
            heads.add(first.split("】", 1)[-1].strip())
    return heads


def family_titles(*families):
    """规则族 → 可能出现在批注抬头里的标题（族标题 + 该族各细粒度标题）。"""

    wanted = set(families)
    titles = {spec.title for key, spec in RULES.items() if key.split(".")[0] in wanted}
    titles |= {item.title for code, item in FINE_RULES.items() if code.split(".")[0] in wanted}
    return titles


#: 批注抬头里表示“确定问题”的等级（“无法判定/不适用/未实现”不算确定结论）
_CONFIRMED_LABELS = ("【错误】", "【警告】", "【提醒】")


def confirmed_heads(comments):
    """确定问题批注的抬头标题集合（批注不写规则编号，按标题核对）。"""

    heads = set()
    for comment in comments:
        first = (comment.text or "").splitlines()[0]
        if first.startswith(_CONFIRMED_LABELS) and "】" in first:
            heads.add(first.split("】", 1)[-1].strip())
    return heads


def assert_reported_comments(report_path, comments, families):
    """以报告的规则编号核对确定问题，再核对其实际批注标题，避免目录标题别名误报。"""
    import json
    findings = json.loads(report_path.read_text(encoding="utf-8"))["findings"]
    heads = confirmed_heads(comments)
    for family in families:
        failures = [f for f in findings if f["rule_id"].split(".")[0] == family
                    and f["status"] == "fail"]
        assert failures, f"报告中缺少规则 {family} 的确定问题"
        assert any(f["title"] in heads for f in failures), f"确定问题 {family} 没有对应原生批注"


def title_families():
    """批注抬头标题 → 规则族（用于从批注回溯规则编号做统计）。"""

    mapping = {}
    for key, spec in RULES.items():
        mapping.setdefault(spec.title, key.split(".")[0])
    for code, item in FINE_RULES.items():
        mapping.setdefault(item.title, code.split(".")[0])
    return mapping


@pytest.fixture
def runtime_dir() -> Path:
    """项目内测试运行目录；避免依赖系统临时目录权限。"""
    path = Path(__file__).parent / "_runtime" / uuid4().hex
    path.mkdir(parents=True, exist_ok=False)
    return path


@dataclass
class FakePageSetup:
    TopMargin: float
    BottomMargin: float
    LeftMargin: float
    RightMargin: float


@dataclass
class FakeSection:
    PageSetup: FakePageSetup


class FakeFormat:
    def __init__(self, points=0.0, characters=0.0):
        self.FirstLineIndent = points
        self.CharacterUnitFirstLineIndent = characters


class FakeParagraph:
    def __init__(self, points=0.0, characters=0.0):
        self.Format = FakeFormat(points, characters)
