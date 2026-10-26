from __future__ import annotations

import re

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.models import Finding, Status, build_comment

CN_PUNCT = "，。；：？！"
EN_PUNCT = ",.;:?!"
OPEN_TO_CLOSE = {"（": "）", "(": ")", "“": "”", "‘": "’", "「": "」", "『": "』", "[": "]", "【": "】"}
CLOSE_TO_OPEN = {v: k for k, v in OPEN_TO_CLOSE.items()}


#: 标点类细粒度规则的“应该是什么”
PUNCT_EXPECTED = {
    "122.1": "中文句子使用中文标点",
    "122.2": "英文句子使用英文标点",
    "123.1": "引号成对出现",
    "123.2": "括号成对出现",
    "123.3": "不出现连续重复标点",
    "123.4": "句末使用句号等句末标点",
}


def _pair_code(char: str) -> str:
    """成对符号 → 引号（123.1）或括号（123.2）。"""

    return "123.1" if char in "“”‘’\"'" else "123.2"


def _finding(rule: str, title: str, text: str, index: int, pos: int, message: str) -> Finding:
    """构造一条标点结论：标题取自规则目录，文案走统一模板。"""

    item = FINE_RULES.get(rule)
    label = item.title if item else title
    severity = item.severity if item else Severity.HINT
    expected = PUNCT_EXPECTED.get(rule, "符合中文/英文标点使用规范")
    finding = Finding(rule, label, Status.FAIL, "", severity=severity, expected=expected, actual=message,
                      paragraph_index=index, start=pos, end=pos + 1,
                      excerpt=text[max(0, pos - 8):pos + 9])
    finding.message = build_comment(severity, rule, label, expected, message)
    return finding


def inspect_punctuation(text: str, index: int) -> list[Finding]:
    """确定性检查：邻接语言、重复标点、成对符号；不推测语义性漏逗号。"""
    findings = []
    stack: list[tuple[str, int]] = []
    for pos, char in enumerate(text):
        if char in OPEN_TO_CLOSE:
            stack.append((char, pos))
        elif char in CLOSE_TO_OPEN:
            if stack and stack[-1][0] == CLOSE_TO_OPEN[char]:
                stack.pop()
            else:
                findings.append(_finding(_pair_code(char), "标点成对", text, index, pos, f"“{char}”缺少匹配的左侧符号或嵌套顺序错误。"))
        if char in EN_PUNCT + CN_PUNCT and 0 < pos < len(text) - 1:
            left, right = text[pos - 1], text[pos + 1]
            if char in ",.;:?!" and re.search(r"[\u3400-\u9fff]", left) and re.search(r"[\u3400-\u9fff]", right):
                # 小数点、英文缩写不会由两侧汉字包围。
                findings.append(_finding("122.1", "中英文标点混用", text, index, pos, f"中文字符之间使用了英文标点“{char}”。"))
            elif char in CN_PUNCT and left.isascii() and left.isalpha() and right.isascii() and right.isalpha():
                findings.append(_finding("122.2", "中英文标点混用", text, index, pos, f"英文字母之间使用了中文标点“{char}”。"))
        if pos and char in ",，。;；:：?!？！" and text[pos - 1] == char:
            findings.append(_finding("123.3", "标点重复", text, index, pos, f"连续重复的“{char}”请核对。"))
        if pos and char == "." and text[pos - 1] == "." and not (pos >= 2 and text[pos - 2] == ".") and not (pos + 1 < len(text) and text[pos + 1] == "."):
            findings.append(_finding("123.3", "标点重复", text, index, pos, "连续两个英文句点，请核对是否多余。"))
    stripped_start = len(text) - len(text.lstrip())
    stripped_end = len(text.rstrip())
    if stripped_start < len(text) and text[stripped_start] in ",，。;；:：?!？！":
        findings.append(_finding("123.3", "疑似多余标点", text, index, stripped_start, "段首出现标点，请核对是否多余。"))
    if stripped_end and text[stripped_end - 1] in ",，;；:：":
        findings.append(_finding("123.4", "疑似缺少句末标点", text, index, stripped_end - 1, "段末以逗号、分号或冒号结束，请核对是否缺少后续内容或句末标点。"))
    for char, pos in stack:
        findings.append(_finding(_pair_code(char), "标点成对", text, index, pos, f"“{char}”缺少匹配的右侧符号。"))
    stripped = text.rstrip()
    if len(stripped) >= 12 and re.search(r"[\u3400-\u9fff]", stripped) and re.search(r"[\u3400-\u9fff]$", stripped):
        findings.append(_finding("123.4", "疑似缺少句末标点", text, index, len(stripped) - 1,
                                 "正文段落以汉字结束，疑似缺少句号等句末标点；请人工核对。"))
    return findings
