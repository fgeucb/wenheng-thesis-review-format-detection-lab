"""只读的域结构扫描工具（页眉页脚、标题、目录等大类共用）。

Field results are saved caches, never evidence of current rendered page numbers.
No field is evaluated, refreshed or rewritten here.
"""
from __future__ import annotations

import re
from thesis_format_checker.core.xmlutil import qn


class FieldScan(list):
    def __init__(self):
        super().__init__()
        self.errors = []


def scan_fields(root):
    """Read simple and nested complex fields in document order.

    Each result has kind, instruction, element, paragraph_element,
    result_elements (w:t/w:tab/w:br leaves), result_paragraph_elements,
    result_text, nested, has_nested, malformed, cache_present, in_textbox.
    A FieldScan also exposes .errors for unmatched markers. The parser retains
    empty result paragraphs and does not depend on python-docx.paragraphs,
    which omits content controls and parts of complex TOC structures.
    """
    result, stack = FieldScan(), []

    def paragraph(node):
        return next((p for p in (node, *node.iterancestors()) if p.tag == qn("w:p")), None)

    def create(node, simple=False):
        record = {
            "kind": "", "instruction": node.get(qn("w:instr"), "") if simple else "",
            "element": node, "paragraph_element": paragraph(node),
            "result_elements": [], "result_paragraph_elements": [], "result_text": "",
            "nested": bool(stack), "has_nested": False, "malformed": False,
            "cache_present": simple,
            "in_textbox": any(a.tag == qn("w:txbxContent") for a in node.iterancestors()),
            "locked": node.get(qn("w:fldLock"), "0") in ("1", "true", "on"),
            "dirty": node.get(qn("w:dirty"), "0") in ("1", "true", "on"),
        }
        if stack:
            stack[-1]["record"]["has_nested"] = True
        result.append(record)
        frame = {"record": record, "phase": "result" if simple else "instruction", "simple": simple}
        stack.append(frame)
        return frame

    def mark_bad(frame, reason):
        frame["record"]["malformed"] = True
        result.errors.append(reason)

    def result_frames():
        # A nested field's code belongs to its instruction, not to outer cache.
        found = []
        for frame in reversed(stack):
            if frame["phase"] != "result":
                break
            found.append(frame)
        return found

    def add_paragraph(frame, node):
        p = paragraph(node)
        target = frame["record"]["result_paragraph_elements"]
        if p is not None and not any(p is old for old in target):
            target.append(p)

    def walk(node):
        if node.tag in (qn("w:del"), qn("w:moveFrom")):
            return
        if node.tag == qn("w:fldSimple"):
            frame = create(node, True)
            add_paragraph(frame, node)
            for child in node:
                walk(child)
            while stack and stack[-1] is not frame:
                mark_bad(stack.pop(), "简单域内存在未闭合的嵌套域")
            if stack:
                stack.pop()
            return
        if node.tag == qn("w:fldChar"):
            kind = node.get(qn("w:fldCharType"))
            if kind == "begin":
                create(node)
            elif kind == "separate":
                if not stack or stack[-1]["simple"]:
                    result.errors.append("存在未匹配的域结果分隔标记")
                elif stack[-1]["phase"] == "result":
                    mark_bad(stack[-1], "同一域存在重复结果分隔标记")
                else:
                    stack[-1]["phase"] = "result"
                    stack[-1]["record"]["cache_present"] = True
                    add_paragraph(stack[-1], node)
            elif kind == "end":
                if not stack or stack[-1]["simple"]:
                    result.errors.append("存在未匹配的域结束标记")
                else:
                    stack.pop()
            else:
                result.errors.append("域标记类型无法读取")
            return
        if node.tag == qn("w:instrText"):
            if stack and stack[-1]["phase"] == "instruction":
                stack[-1]["record"]["instruction"] += node.text or ""
            elif not stack:
                result.errors.append("存在没有起始标记的域指令")
            return
        if node.tag == qn("w:p"):
            for frame in result_frames():
                add_paragraph(frame, node)
        if node.tag in (qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")):
            text = node.text or "" if node.tag == qn("w:t") else "\t" if node.tag == qn("w:tab") else "\n"
            for frame in result_frames():
                frame["record"]["result_elements"].append(node)
                frame["record"]["result_text"] += text
                add_paragraph(frame, node)
        for child in node:
            walk(child)

    walk(root)
    for frame in stack:
        mark_bad(frame, "域未闭合")
    for record in result:
        record["instruction"] = record["instruction"].strip()
        match = re.match(r"\s*([A-Za-z]+)(?=\s|\\|$)", record["instruction"])
        record["kind"] = match.group(1).upper() if match else ""
        if not match:
            record["malformed"] = True
            result.errors.append("域指令缺失或无法识别")
    return result


def scan_body_fields(ctx):
    """整篇文档的域扫描：同一份文档只算一次（见 core.docx_context.memo）。"""

    from thesis_format_checker.core.docx_context import memo

    return memo(ctx, "body_fields", lambda: scan_fields(ctx.document.element.body))
