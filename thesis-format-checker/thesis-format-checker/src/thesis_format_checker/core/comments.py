from __future__ import annotations

from thesis_format_checker.core.models import Finding, Status


def comment_allowed(finding: Finding | None = None, *, message: str = "") -> bool:
    """所有写入入口共用的门禁：只有确定问题可以成为批注。"""
    if finding is not None:
        if finding.status != Status.FAIL:
            return False
        message = finding.message
    return "无法判定" not in message


def paragraph_comment_limitation(paragraph) -> str | None:
    # Even native Word Comments.Add converts these display equations to inline
    # layout when it inserts a comment-reference run. Preserve the source math.
    if paragraph._p.xpath(".//m:oMathPara"):
        return "显示公式段落插入批注会改变 Word 排版，仅保留报告"
    return None


def add_comment(document, finding: Finding) -> None:
    if not comment_allowed(finding):
        return
    # COM 旧入口没有可靠的逐页边界映射，页眉页脚问题只能由页级写入器处理。
    if finding.page_story or finding.header_kind is not None:
        return
    if finding.paragraph_index is not None:
        paragraph = document.Paragraphs(finding.paragraph_index)
        offset, length = finding.start, None if finding.end is None else finding.end - (finding.start or 0)
        if offset is None and finding.excerpt:
            # 只有段落号没有偏移时，用摘录文字在段落里定位，让批注贴住那一行/那几个字。
            text = str(paragraph.Range.Text).rstrip("\r\x07")
            position = text.find(finding.excerpt)
            if position >= 0:
                offset, length = position, len(finding.excerpt)
        text_end = max(paragraph.Range.Start, paragraph.Range.End - 1)
        start = min(paragraph.Range.Start + (offset or 0), text_end)
        requested_end = paragraph.Range.Start + (offset or 0) + max(length or 1, 1)
        end = min(max(start + 1, requested_end), paragraph.Range.End)
        anchor = document.Range(start, end)
    elif finding.section_index is not None:
        section = document.Sections(finding.section_index)
        anchor = _section_anchor(document, section)
    else:
        return
    document.Comments.Add(Range=anchor, Text=finding.message)


def _section_anchor(document, section):
    section_start = section.Range.Start
    section_end = section.Range.End
    for index in range(1, document.Paragraphs.Count + 1):
        paragraph = document.Paragraphs(index)
        if paragraph.Range.Start < section_start:
            continue
        if paragraph.Range.Start >= section_end:
            break
        text = paragraph.Range.Text.rstrip("\r\x07")
        if text.strip():
            return document.Range(paragraph.Range.Start, min(paragraph.Range.Start + 1, paragraph.Range.End))
    return document.Range(section_start, min(section_start + 1, section_end))


def add_word_comment(document, finding: Finding, paragraph=None) -> None:
    """Reuse Word's original writer or python-docx's native comment API.

    Explicit character ranges are split losslessly at both ends; unsupported
    complex ranges are reported rather than expanded to unrelated text.
    """
    if not comment_allowed(finding):
        return
    if finding.page_story or finding.header_kind is not None:
        return
    if not hasattr(document, "add_comment"):
        add_comment(document, finding)
        return
    from docx.oxml import OxmlElement
    from thesis_format_checker.core.xmlutil import qn
    from docx.text.run import Run

    if paragraph is None:
        from thesis_format_checker.core.docx_context import DocxContext
        candidates = DocxContext(document).paragraphs
        index = (finding.paragraph_index or 0) - 1
        if 0 <= index < len(candidates):
            candidate = candidates[index]
            if not finding.excerpt or finding.excerpt in candidate.text:
                paragraph = candidate
    if paragraph is None:
        return
    if paragraph_comment_limitation(paragraph):
        return
    start, end = finding.start, finding.end
    if start is None and finding.excerpt and paragraph.text.count(finding.excerpt) == 1:
        start = paragraph.text.index(finding.excerpt)
        end = start + len(finding.excerpt)
    if start is not None or end is not None:
        from thesis_format_checker.core.point_comments import add_range_comment
        start = int(start or 0)
        return add_range_comment(document, paragraph, start,
                                 int(end) if end is not None else start + 1, finding.message)
    runs = [Run(r, paragraph) for r in paragraph._p.iter(qn("w:r"))
            if not any(a.tag in (qn("w:txbxContent"), qn("w:del")) for a in r.iterancestors())
            and not r.xpath("./w:commentReference")]
    if runs:
        document.add_comment(runs, text=finding.message, author="论文格式检测", initials="TC")
        return
    comment = document.comments.add_comment(finding.message, author="论文格式检测", initials="TC")
    start, end, ref_run, ref = (OxmlElement(tag) for tag in
                                ("w:commentRangeStart", "w:commentRangeEnd", "w:r", "w:commentReference"))
    for node in (start, end, ref):
        node.set(qn("w:id"), str(comment.comment_id))
    ref_run.append(ref)
    paragraph._p.extend((start, end, ref_run))


