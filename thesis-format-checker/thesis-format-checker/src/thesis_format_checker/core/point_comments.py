"""字符级批注及页边界零长度批注，不改变正文内容与格式。

本机 Word 16.0 拒绝页眉页脚内批注，但支持正文零长度批注。
页级问题仅在 Word 已验证的本页正文首末边界写空范围，不圈选标题、图片或正文。
无法匹配边界时仅报告，不新增空格、段落、图形或回退位置。

实现要点：

* 只在“段落里除 pPr/批注标记外全是普通 run，且 run 内只有文字类子节点”时动手，
  含域、修订、嵌入对象的段落一律不碰（字符位置无法用文本长度推算）；
* 需要时按字符**无损拆分 run**：左右两半复制原 run 的全部属性，只切分 w:t 文本；
* 拆分后的文档用 :func:`content_tree` 校验——忽略批注并规范化相邻同格式 run，
  保证“除了批注，文字与格式没有变化”。
"""
from __future__ import annotations

from copy import deepcopy

from thesis_format_checker.core.xmlutil import qn
from docx.text.run import Run

from thesis_format_checker.core.word_session import WordAutomationError

MARKERS = {qn("w:commentRangeStart"), qn("w:commentRangeEnd"),
           qn("w:bookmarkStart"), qn("w:bookmarkEnd")}
TEXT = {qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr")}


def safe_paragraph(paragraph) -> bool:
    """段落是否可以安全地按字符定位：不能含域、修订、嵌套对象。"""

    for child in paragraph._p:
        if child.tag == qn("w:pPr") or child.tag in MARKERS:
            continue
        if child.tag != qn("w:r"):
            return False
        if any(node.tag not in TEXT | {qn("w:rPr"), qn("w:lastRenderedPageBreak"),
                                       qn("w:commentReference")}
               for node in child):
            return False
    return True


def safe_field_edge(paragraph) -> bool:
    """Only outer edges of self-contained text fields; never split their code."""
    if paragraph._p.xpath(".//w:drawing|.//w:pict|.//w:object|.//m:oMath|.//w:ins|.//w:del|.//w:br"):
        return False
    depth = 0
    fields = paragraph._p.xpath(".//w:fldChar")
    if not fields:
        return False
    for node in fields:
        kind = node.get(qn("w:fldCharType"))
        if kind == "begin":
            depth += 1
        elif kind == "end":
            depth -= 1
        if depth < 0:
            return False
    return depth == 0


def _length(child) -> int:
    return len(child.text or "") if child.tag == qn("w:t") else int(child.tag in TEXT)


def boundary_text(paragraph) -> str:
    """含显式分页符的 Word 字符流；python-docx 的 paragraph.text 会省略分页符。"""
    if not safe_paragraph(paragraph):
        return paragraph.text
    parts = []
    for run in paragraph._p.findall(qn("w:r")):
        for node in run:
            if node.tag == qn("w:t"):
                parts.append(node.text or "")
            elif node.tag == qn("w:tab"):
                parts.append("\t")
            elif node.tag == qn("w:br"):
                parts.append("\f" if node.get(qn("w:type")) == "page" else "\n")
            elif node.tag == qn("w:cr"):
                parts.append("\n")
    return "".join(parts)


def _split(run, cut: int):
    """在 run 内部第 ``cut`` 个字符处把它拆成两个格式完全相同的 run。"""

    left, right = deepcopy(run), deepcopy(run)
    for result in (left, right):
        for child in list(result):
            if child.tag != qn("w:rPr"):
                result.remove(child)
    position = 0
    for child in run:
        if child.tag == qn("w:rPr"):
            continue
        size = _length(child)
        if position < cut < position + size:
            if child.tag != qn("w:t"):
                raise WordAutomationError("不能在非文字对象中拆分批注范围")
            first, second = deepcopy(child), deepcopy(child)
            first.text, second.text = child.text[:cut - position], child.text[cut - position:]
            for node in (first, second):
                node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            left.append(first)
            right.append(second)
        elif position < cut:
            left.append(deepcopy(child))
        else:
            right.append(deepcopy(child))
        position += size
    parent = run.getparent()
    index = parent.index(run)
    parent.remove(run)
    parent.insert(index, left)
    parent.insert(index + 1, right)
    return left, right


def _boundary(paragraph, offset: int) -> None:
    """把段落里的 run 边界调整到 ``offset`` 处（必要时拆分）。"""

    position = 0
    for run in list(paragraph._p.findall(qn("w:r"))):
        size = sum(_length(child) for child in run)
        if position < offset < position + size:
            _split(run, offset - position)
            return
        position += size


def add_point_comment(document, paragraph, offset: int, message: str):
    """把批注锚定在段落第 ``offset`` 个字符上（单字符范围）。"""

    from thesis_format_checker.core.comments import comment_allowed
    if not comment_allowed(message=message):
        return None
    if not safe_paragraph(paragraph) or not 0 <= offset < len(paragraph.text):
        raise WordAutomationError("正文字符批注位置无法安全匹配")
    _boundary(paragraph, offset)
    _boundary(paragraph, offset + 1)
    position = 0
    for run in paragraph._p.findall(qn("w:r")):
        size = sum(_length(child) for child in run)
        if position == offset and size == 1:
            return document.add_comment([Run(run, paragraph)], text=message,
                                        author="论文格式检测", initials="TC")
        position += size
    raise WordAutomationError("未能定位批注指定的正文字符")


def add_range_comment(document, paragraph, start: int, end: int, message: str):
    """Exact existing text range, including multiple differently formatted runs."""
    from thesis_format_checker.core.comments import comment_allowed
    if not comment_allowed(message=message):
        return None
    # A standalone signature image outside the selected text does not make
    # adjacent label text ambiguous. Do not permit fields or mixed object runs.
    safe = safe_paragraph(paragraph)
    if not safe:
        copy = deepcopy(paragraph._p)
        for run in list(copy.findall(qn("w:r"))):
            children = [n for n in run if n.tag != qn("w:rPr")]
            if children and all(n.tag in {qn("w:drawing"), qn("w:pict"), qn("w:object")} for n in children):
                copy.remove(run)
        from docx.text.paragraph import Paragraph
        safe = safe_paragraph(Paragraph(copy, paragraph._parent))
    if (not safe or boundary_text(paragraph) != paragraph.text
            or not 0 <= start < end <= len(paragraph.text)):
        raise WordAutomationError("文字范围无法无损精确定位")
    _boundary(paragraph, start)
    _boundary(paragraph, end)
    runs, position = [], 0
    for run in paragraph._p.findall(qn("w:r")):
        size = sum(_length(child) for child in run)
        if size and start <= position and position + size <= end:
            runs.append(Run(run, paragraph))
        position += size
    if not runs:
        raise WordAutomationError("文字范围没有可标注的字符")
    return document.add_comment(runs, text=message, author="论文格式检测", initials="TC")


def add_boundary_comment(document, paragraph, offset: int, message: str, *, edge=None):
    """既有正文边界的零长度批注；不添加文字、不圈选图片或无关正文。"""
    from docx.oxml import OxmlElement
    from thesis_format_checker.core.comments import comment_allowed

    if not comment_allowed(message=message):
        return None
    field_edge = edge in ("start", "end") and safe_field_edge(paragraph)
    if field_edge:
        if offset != (0 if edge == "start" else len(paragraph.text)):
            raise WordAutomationError("域外边界与偏移不匹配")
    elif not safe_paragraph(paragraph) or not 0 <= offset <= len(boundary_text(paragraph)):
        raise WordAutomationError("页边界无法无损定位")
    if offset == 0 and boundary_text(paragraph).startswith("\f"):
        raise WordAutomationError("分页符前的批注会改变下一页段前距，不能作为页边界")
    if not field_edge:
        _boundary(paragraph, offset)
    position = 0
    insert_at = len(paragraph._p)
    for index, node in enumerate(paragraph._p) if not field_edge else []:
        if node.tag != qn("w:r") or node.xpath("./w:commentReference"):
            continue
        if position == offset:
            insert_at = index
            break
        position += sum(_length(child) for child in node)
    if field_edge:
        insert_at = (1 if paragraph._p.find(qn("w:pPr")) is not None else 0) if edge == "start" else len(paragraph._p)
    elif position != offset:
        raise WordAutomationError("页边界字符偏移不匹配")
    comment = document.comments.add_comment(message, author="论文格式检测", initials="TC")
    start, end, run, ref = [OxmlElement(tag) for tag in
                          ("w:commentRangeStart", "w:commentRangeEnd", "w:r", "w:commentReference")]
    for node in (start, end, ref):
        node.set(qn("w:id"), str(comment.comment_id))
    run.append(ref)
    for index, node in enumerate((start, end, run), insert_at):
        paragraph._p.insert(index, node)
    return comment


def marker_offset(paragraph, marker) -> int:
    """批注标记在段落内的字符偏移（用于输出后复核锚点）。"""

    offset = 0
    for node in paragraph.iter():
        if node is marker:
            return offset
        # w:tabs/w:tab in paragraph properties is a tab STOP, not body text.
        if not any(a.tag in {qn("w:pPr"), qn("w:rPr")} for a in node.iterancestors()):
            offset += _length(node)
    raise WordAutomationError("批注锚点不在指定段落中")


def _semantic_tree(element):
    """忽略批注标记的语义树（元素、属性、文本、子节点全比较）。"""

    tree = deepcopy(element)
    for node in list(tree.iter()):
        if node.tag in (qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")):
            parent = node.getparent()
            parent.remove(node)
            if parent.tag == qn("w:r") and not any(child.tag != qn("w:rPr") for child in parent):
                parent.getparent().remove(parent)

    def shape(node):
        return (node.tag, tuple(sorted(node.attrib.items())), node.text or "", node.tail or "",
                tuple(shape(child) for child in node))

    return shape(tree)


def content_tree(element):
    """忽略批注并规范化相邻同格式 run 的内容树；只允许“无损拆分”，不放宽校验。"""

    root = deepcopy(element)
    for node in list(root.iter()):
        if node.tag in {qn("w:commentRangeStart"), qn("w:commentRangeEnd"), qn("w:commentReference")}:
            parent = node.getparent()
            parent.remove(node)
            if parent.tag == qn("w:r") and not any(child.tag != qn("w:rPr") for child in parent):
                parent.getparent().remove(parent)
    for parent in root.iter():
        previous = None
        for run in list(parent):
            if run.tag != qn("w:r"):
                previous = None
                continue
            prop = run.find(qn("w:rPr"))
            signature = (tuple(sorted(run.attrib.items())),
                         _semantic_tree(prop) if prop is not None else None)
            if previous is not None and previous[1] == signature:
                for child in list(run):
                    if child.tag != qn("w:rPr"):
                        previous[0].append(child)
                parent.remove(run)
            else:
                previous = (run, signature)
    for run in root.iter(qn("w:r")):
        previous = None
        for child in list(run):
            if child.tag != qn("w:t"):
                previous = None
                continue
            child.attrib.pop("{http://www.w3.org/XML/1998/namespace}space", None)
            if previous is not None and dict(previous.attrib) == dict(child.attrib):
                previous.text = (previous.text or "") + (child.text or "")
                run.remove(child)
            else:
                previous = child
    return _semantic_tree(root)
