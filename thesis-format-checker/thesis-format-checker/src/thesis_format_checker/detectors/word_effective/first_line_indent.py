from __future__ import annotations

from thesis_format_checker.core.effective_format import normalized_number
from thesis_format_checker.core.models import Finding, IndentSpec, Status

#: 哨兵：区分"没传实际缩进（自己去读）"与"读了但没读到（None）"。
_UNSET = object()


def read_indent(paragraph) -> IndentSpec | None:
    points = normalized_number(paragraph.Format.FirstLineIndent)
    if points is None:
        return None
    characters = None
    try:
        characters = normalized_number(paragraph.Format.CharacterUnitFirstLineIndent)
    except Exception:
        pass
    return IndentSpec(points=round(points, 3), characters=characters)


def compare_first_line_indent(
    paragraph,
    text: str,
    expected: IndentSpec | None,
    paragraph_index: int,
    tolerance_pt: float = 0.5,
    tolerance_chars: float = 0.05,
    actual_indent=_UNSET,
) -> list[Finding]:
    """比较首行缩进。

    ``actual_indent``：调用方已经读过本段缩进时直接传入，避免同一属性读两遍
    （COM 每次访问都要过一遍 Word；正文段落里这是最常见的重复读）。
    """

    findings: list[Finding] = []
    prefix = ""
    for char in text:
        if char in (" ", "\u3000", "\t"):
            prefix += char
        else:
            break
    if prefix:
        counts = [("半角空格", prefix.count(" ")), ("全角空格", prefix.count("\u3000")), ("制表符（Tab）", prefix.count("\t"))]
        visible = "、".join(f"{count} 个{name}" for name, count in counts if count)
        findings.append(
            Finding(
                rule_id="072",
                title="正文首行缩进",
                status=Status.FAIL,
                message=f"【072 正文首行缩进】段首使用了{visible}模拟缩进，请改用段落首行缩进。",
                expected="段落格式缩进",
                actual=visible,
                paragraph_index=paragraph_index,
                start=0,
                end=max(1, len(prefix)),
                excerpt=prefix,
            )
        )
    actual = read_indent(paragraph) if actual_indent is _UNSET else actual_indent
    if actual is None:
        return findings
    if expected is None:
        expected = IndentSpec(0, 2)
        if actual.characters is None:
            return findings
    # 用户明确要求 2 字符；模板 A 的磅值只用于无法读到字符单位时辅助判断。
    expected_chars = 2.0
    if expected.characters is not None and actual.characters is not None:
        differs = abs(expected_chars - actual.characters) > tolerance_chars
        expected_label = "2 字符"
        actual_label = f"{actual.characters:g} 字符（{actual.points:g} 磅）"
    else:
        if expected.characters is None or abs(expected.characters - 2) > tolerance_chars:
            return findings  # 无法可靠地把模板磅值换算为 2 字符
        differs = abs(expected.points - actual.points) > tolerance_pt
        expected_label = f"2 字符（模板约 {expected.points:g} 磅）"
        actual_label = f"{actual.points:g} 磅"
    if differs:
        findings.append(
            Finding(
                rule_id="072",
                title="正文首行缩进",
                status=Status.FAIL,
                message=(
                    f"【072 正文首行缩进】模板 A：{expected_label}；当前段落：{actual_label}。"
                    "请核对本段首行缩进。"
                ),
                expected=expected_label,
                actual=actual_label,
                paragraph_index=paragraph_index,
                start=0,
                end=min(max(len(text), 1), 2),
                excerpt=text[:12],
            )
        )
    return findings
