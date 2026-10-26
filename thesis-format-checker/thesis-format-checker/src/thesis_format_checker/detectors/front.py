"""封面与前置部分检测（012–022）的取证实现（不做语义推断）。

014: 本条依赖大模型语义，本阶段暂不实现
017: 本条依赖大模型语义，本阶段暂不实现
020: 本条依赖大模型语义，本阶段暂不实现

Only the template supplies formatting values.  Labels and placeholder syntax
below identify document roles; they do not prescribe a school's requirements.
"""
from __future__ import annotations

import re
from datetime import date
from typing import Any

from thesis_format_checker.core.xmlutil import qn

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import (
    DocxContext, describe_value, effective_paragraph, effective_runs,
    make_finding, unknown,
)


_LABELS = {
    "title_cn": ("中文论文题目", "论文中文题目", "论文题目", "论文题名", "中文题目", "题目", "题名"),
    "title_en": ("英文论文题目", "论文英文题目", "英文题目", "英文题名", "Thesis title", "Title"),
    "advisor_name": ("指导教师姓名", "指导教师", "导师姓名", "导师", "Supervisor"),
    "advisor_rank": ("指导教师职称", "导师职称", "职称"),
    "advisor_org": ("指导教师单位", "导师单位", "指导单位"),
    "submission_date": ("论文提交日期", "论文提交时间", "提交日期", "提交时间"),
    "defense_date": ("论文答辩日期", "答辩日期", "答辩时间"),
    "confer_date": ("学位授予日期", "学位授予时间", "授予日期"),
    "secrecy": ("密级",),
    "classification": ("中图分类号", "分类号"),
    "udc": ("UDC",),
    "author": ("作者姓名", "作者", "姓名"),
    "student_id": ("学号",),
    "college": ("学院", "院系"),
    "major": ("专业名称", "学科专业", "学科名称", "专业"),
    "class": ("班级",),
}
_LABEL_LOOKUP = {
    re.sub(r"\s+", "", label).lower(): role
    for role, labels in _LABELS.items() for label in labels
}
_CN = re.compile(r"[\u3400-\u9fff]")
_LATIN = re.compile(r"[A-Za-z0-9]")
_OPTIONAL = re.compile(r"选填|可选|可不填|无需填写|不必填写|如有|适用时|按需")
_REQUIRED = re.compile(r"必填|必须填写|不能为空|不得为空|应填写|须填写")
_PLACEHOLDER = re.compile(
    r"^(?:[_＿…\.．□○\-\s]+|[xX×]+|待填(?:写)?|请(?:输入|填写).+|"
    r"(?:YYYY|yyyy|年).*)$"
)
_FRONT_RULES = {
    "012": "论文题目缺失或不一致", "013": "封面题目格式不符合模板",
    "015": "指导教师信息缺失或格式不统一", "016": "日期缺失或格式错误",
    "018": "封面元素布局与模板不一致", "019": "密级、分类号或UDC缺失或位置错误",
    "021": "摘要排版格式不符合模板",
    "014": "作者姓名、学号、学院、专业、班级等基本信息缺失或填写错误",
    "020": "原创性声明中的签名、日期或必填信息缺失",
}

#: 封面基本信息字段 → 细粒度规则编号与中文名（014）。
_AUTHOR_FIELD_CODES = {"author": "014.1", "student_id": "014.2", "college": "014.3",
                       "major": "014.4", "class": "014.5"}
_AUTHOR_FIELD_LABELS = {"author": "作者姓名", "student_id": "学号", "college": "学院/院系",
                        "major": "专业", "class": "班级"}
#: 声明行：签名与日期（020）。
_DECLARATION_HEADING = re.compile(r"(原创性声明|独创性声明|^\s*声\s*明\s*$)")
#: 授权声明页标题：前置区里的下一页，原创性声明区域到此结束。
_AUTHORIZATION_HEADING = re.compile(r"(授权使用声明|使用授权声明|授权声明|使用授权书|学位论文使用授权)")
_SIGNATURE_LABEL = re.compile(r"(?:论文作者签名|作者签名|签名)\s*[:：]")
_DATE_LABEL = re.compile(r"日期\s*[:：]")
_DECLARATION_CODES = {"signature": "020.1", "date": "020.2"}
#: 文本框里出现这些词才怀疑前置字段被藏在文本框里（模板自带的说明文本框不算）。
_FIELD_HINTS = ("题目", "题名", "姓名", "学号", "学院", "专业", "班级",
                "指导教师", "导师", "签名", "日期", "摘要", "关键词")


def _textboxes_may_hide_fields(ctx: DocxContext) -> bool:
    """文本框内容是否可能包含前置字段；纯说明性文本框（如“宋体小五号”）不算。"""

    try:
        texts = [str(text) for text in ctx.document.element.xpath(".//w:txbxContent//w:t/text()")]
    except Exception:
        return False
    return any(any(hint in text for hint in _FIELD_HINTS) for text in texts)
_PARA_KEYS = (
    "alignment", "first_line_indent", "left_indent", "right_indent",
    "space_before", "space_after", "line_spacing", "line_spacing_rule",
)
_FONT_KEYS = ("east_asia_font", "ascii_font", "size_pt", "bold")
_UNREADABLE = "该对象docx库无法读取，本条暂无法执行检测"
_PROPERTY_LABELS = {
    "east_asia_font": "中文字体", "ascii_font": "英文及数字字体", "size_pt": "字号",
    "bold": "加粗", "alignment": "对齐方式", "first_line_indent": "首行缩进",
    "left_indent": "左缩进", "right_indent": "右缩进", "space_before": "段前间距",
    "space_after": "段后间距", "line_spacing": "行距数值", "line_spacing_rule": "行距类型",
    "container": "所在容器", "table_column": "表格网格列位置", "cell_width": "单元格宽度",
    "table_indent": "表格缩进", "table_alignment": "表格对齐方式",
    "cell_vertical_alignment": "单元格垂直对齐方式",
}

#: 封面题目格式：属性 → 细粒度规则编号
_COVER_TITLE_CODES = {
    "east_asia_font": "013.1",
    "ascii_font": "013.2",
    "size_pt": "013.3",
    "bold": "013.4",
    "line_spacing": "013.5",
    "line_spacing_rule": "013.5",
}
#: 封面元素布局：属性 → 细粒度规则编号
_LAYOUT_CODES = {
    "alignment": "018.1", "table_alignment": "018.1", "cell_vertical_alignment": "018.1",
    "space_before": "018.2", "space_after": "018.2",
    "line_spacing": "018.2", "line_spacing_rule": "018.2",
    "left_indent": "018.3", "right_indent": "018.3", "table_indent": "018.3",
    "cell_width": "018.3", "table_column": "018.3", "container": "018.3",
}
#: 指导教师字段：角色 → 缺失时使用的细粒度规则
_ADVISOR_MISSING_CODES = {"advisor_name": "015.1", "advisor_rank": "015.2", "advisor_org": "015.3"}
#: 需填字段（密级/分类号/UDC）：角色 → 缺失时使用的细粒度规则
_CLASSIFICATION_MISSING_CODES = {"secrecy": "019.1", "classification": "019.2", "udc": "019.3"}


def _title(code: str) -> str:
    """细粒度规则标题取自规则目录；传入族编号时用本模块的族标题。"""

    item = FINE_RULES.get(code)
    return item.title if item else _FRONT_RULES[code]


def _field_codes(role: str) -> tuple[str, str]:
    """按字段角色给出（缺失时规则, 格式/位置不符规则）。"""

    if role.startswith("advisor_"):
        return _ADVISOR_MISSING_CODES.get(role, "015.4"), "015.4"
    if role.endswith("_date"):
        return "016.1", "016.2"
    return _CLASSIFICATION_MISSING_CODES.get(role, "019.4"), "019.4"


def _unreadable(rule, reason, paragraph=None):
    # 标题统一经 _title 解析，既支持族编号也支持细粒度编号（015.4 等）
    return unknown(rule, _title(rule), _UNREADABLE + "；" + reason, paragraph)


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text).lower()


def _label(text: str):
    udc_only = re.fullmatch(r"\s*UDC(?:\s*注\s*\d+)?\s*", text, re.I)
    if udc_only:
        return "udc", "UDC", "", udc_only.end()
    # An exact label or an explicit colon/tab boundary avoids matching prose.
    m = re.match(r"^\s*(.+?)([:：\t])\s*(.*)$", text, flags=re.S)
    if m:
        role = _LABEL_LOOKUP.get(_compact(m.group(1)))
        if role:
            return role, m.group(1).strip(), m.group(3), m.start(3)
    role = _LABEL_LOOKUP.get(_compact(text))
    if role:
        return role, text.strip(), "", len(text)
    # 格式指南型模板常用空格排版而不用冒号，例如
    # “指导教师姓名            ××× 教授”。标签内部也可能逐字空开。
    aliases = sorted(
        ((role_name, alias) for role_name, labels in _LABELS.items() for alias in labels),
        key=lambda item: len(item[1]), reverse=True,
    )
    for role_name, alias in aliases:
        pattern = r"^\s*" + r"\s*".join(re.escape(char) for char in alias)
        match = re.match(pattern + r"(?:[:：\t]|\s{2,})(.*)$", text, re.I | re.S)
        if match:
            return role_name, text[:match.start(1)].strip(" ：:\t"), match.group(1), match.start(1)
    return None


def _compound_labels(text: str) -> list[tuple[str, str, str, int]]:
    """同一段里的“分类号 … 密级 …”拆成两个独立字段。"""

    matches = []
    for role in ("classification", "secrecy"):
        for alias in _LABELS[role]:
            pattern = r"\s*".join(re.escape(char) for char in alias)
            found = re.search(pattern, text, re.I)
            if found:
                matches.append((found.start(), found.end(), role, found.group(0)))
                break
    if len(matches) < 2:
        return []
    matches.sort()
    result = []
    for position, (start, end, role, label) in enumerate(matches):
        stop = matches[position + 1][0] if position + 1 < len(matches) else len(text)
        value_start = end
        value = text[value_start:stop].strip(" ：:\t\u3000")
        result.append((role, label.strip(), value, value_start))
    return result


def _boundary(paragraph) -> bool:
    text = _compact(paragraph.text)
    if text in {"摘要", "中文摘要", "英文摘要", "abstract", "目录", "contents", "题名页",
                "原创性声明", "独创性声明", "授权使用声明", "学位论文原创性声明"}:
        return True
    if re.match(r"^(第[一二三四五六七八九十百0-9]+章|\d+[.、\s])", paragraph.text.strip()):
        return True
    style = paragraph.style
    return bool(style and re.match(r"^(heading|标题)\s*\d+$", style.name, re.I))


def _placeholder(value: str, role: str = "") -> bool:
    value = value.strip()
    # Parentheses/brackets also occur in real titles and confidentiality values.
    # A bracket alone is not evidence that a populated field is a placeholder.
    bracketed = re.fullmatch(r"[〈<【\[（(](.*)[〉>】\]）)]", value)
    if bracketed:
        value = bracketed.group(1).strip()
    return not value or bool(_PLACEHOLDER.fullmatch(value)) or (
        _compact(value) in _LABEL_LOOKUP and _LABEL_LOOKUP[_compact(value)] == role
    ) or (
        role.endswith("_date") and bool(re.fullmatch(r"[YyMmDd_＿年月日./\-\s]+", value))
    )


def _requirement(record: dict, ctx: DocxContext):
    related = [record["label"] + record["value"]]
    aliases = _LABELS.get(record["role"], (record["label"],))
    for text in ctx.texts:
        if any(_compact(alias) in _compact(text) for alias in aliases) and (
            _REQUIRED.search(text) or _OPTIONAL.search(text)
        ):
            related.append(text)
    joined = "\n".join(related)
    required, optional = bool(_REQUIRED.search(joined)), bool(_OPTIONAL.search(joined))
    if required and optional:
        return "conflict"
    if optional:
        return False
    if required or _placeholder(record["value"], record["role"]):
        return True
    return None


def _unlabelled_title_records(ctx: DocxContext, first_boundary: int) -> list[dict]:
    """识别规范模板常见的无标签封面题目、封二题目和英文题名。

    学校模板经常只用“（题名和副题名）”说明封面题目位置，并未写
    “论文题目：”。这里使用明确的区域边界，不按字号或任意居中文字猜测题目。
    封面题目为空时仍返回一个空记录，使缺失批注能锚在题名占位说明处。
    """

    paragraphs = ctx.paragraphs
    compact = [_compact(paragraph.text) for paragraph in paragraphs]
    records: list[dict] = []

    def add(role: str, area: str, start: int, end: int, *, empty_anchor: int | None = None):
        candidates = [
            index for index in range(start + 1, end)
            if paragraphs[index].text.strip()
            and _label(paragraphs[index].text) is None
            and not re.fullmatch(r"[（(].{0,30}[）)]", paragraphs[index].text.strip())
        ]
        if role == "title_cn":
            value = "".join(paragraphs[index].text.strip() for index in candidates)
        else:
            value = " ".join(paragraphs[index].text.strip() for index in candidates)
        anchor_index = candidates[0] if candidates else (empty_anchor if empty_anchor is not None else end)
        paragraph = paragraphs[anchor_index]
        records.append({
            "role": role, "label": "论文题目" if role == "title_cn" else "英文论文题目",
            "value": value, "paragraph": paragraph, "label_paragraph": paragraph,
            "offset": 0, "index": anchor_index, "cover": anchor_index < first_boundary,
            "area": area,
        })

    marker = next((index for index, value in enumerate(compact)
                   if "题名和副题名" in value), None)
    if marker is not None:
        start = next((index for index in range(marker - 1, -1, -1)
                      if compact[index] == "硕士学位论文" or compact[index] == "博士学位论文"), None)
        if start is not None:
            add("title_cn", "cover", start, marker, empty_anchor=marker)

        author_marker = next((index for index in range(marker + 1, min(first_boundary, len(paragraphs)))
                              if "作者姓名" in compact[index]), None)
        if author_marker is not None:
            author_candidates = [
                index for index in range(marker + 1, author_marker)
                if paragraphs[index].text.strip() and _label(paragraphs[index].text) is None
            ]
            author_index = author_candidates[-1] if author_candidates else author_marker
            author_paragraph = paragraphs[author_index]
            records.append({
                "role": "author", "label": "作者姓名",
                "value": author_paragraph.text.strip() if author_candidates else "",
                "paragraph": author_paragraph, "label_paragraph": paragraphs[author_marker],
                "offset": 0, "index": author_index, "cover": author_index < first_boundary,
                "area": "cover",
            })

        title_page_start = next((index for index in range(marker + 1, min(first_boundary, len(paragraphs)))
                                 if compact[index] in {"硕士学位论文", "博士学位论文"}), None)
        if title_page_start is not None:
            title_page_end = next((index for index in range(title_page_start + 1, first_boundary)
                                   if (_label(paragraphs[index].text) or (None,))[0] == "author"), None)
            if title_page_end is not None:
                add("title_cn", "title_page", title_page_start, title_page_end)

    english_start = next((index for index, value in enumerate(compact[:first_boundary])
                          if value in {"masterdissertation", "doctoraldissertation", "phddissertation"}), None)
    if english_start is not None:
        english_end = next((index for index in range(english_start + 1, first_boundary)
                            if compact[index] == "by"), None)
        if english_end is not None:
            add("title_en", "english_title_page", english_start, english_end)
    return records


def _fields(ctx: DocxContext) -> list[dict]:
    records = []
    first_boundary = next((i for i, p in enumerate(ctx.paragraphs) if _boundary(p)), len(ctx.paragraphs))
    consumed = set()
    for i, paragraph in enumerate(ctx.paragraphs):
        compound = _compound_labels(paragraph.text)
        if compound:
            for role, label, value, offset in compound:
                records.append({"role": role, "label": label, "value": value,
                                "paragraph": paragraph, "label_paragraph": paragraph,
                                "offset": offset, "index": i, "cover": i < first_boundary})
            continue
        item = _label(paragraph.text)
        if item:
            role, label, value, offset = item
            value_p = paragraph
            if not value.strip() and i + 1 < len(ctx.paragraphs):
                candidate = ctx.paragraphs[i + 1]
                if (candidate.text.strip() and not _label(candidate.text) and not _boundary(candidate)
                        and not _REQUIRED.search(candidate.text) and not _OPTIONAL.search(candidate.text)):
                    # In a table only use the same cell or the next cell in this
                    # row; do not consume an unrelated row as the field value.
                    tc = _ancestor(paragraph._p, "tc")
                    other_tc = _ancestor(candidate._p, "tc")
                    adjacent = tc is None and other_tc is None
                    if tc is not None and other_tc is not None:
                        adjacent = tc is other_tc or (
                            tc.getparent() is other_tc.getparent() and tc.getnext() is other_tc
                        )
                    if adjacent:
                        value_p, value, offset = candidate, candidate.text, 0
                        consumed.add(i + 1)
            records.append({"role": role, "label": label, "value": value, "paragraph": value_p,
                            "label_paragraph": paragraph, "offset": offset, "index": i,
                            "cover": i < first_boundary})
        elif i not in consumed and i < first_boundary and paragraph.text.strip():
            style_name = paragraph.style.name if paragraph.style else ""
            if style_name.lower() in {"title", "标题"}:
                role = "title_cn" if _CN.search(paragraph.text) else "title_en"
                records.append({"role": role, "label": "论文题目", "value": paragraph.text,
                                "paragraph": paragraph, "label_paragraph": paragraph, "offset": 0,
                                "index": i, "cover": True})
    known = {(record["role"], id(record["paragraph"]._p)) for record in records}
    for record in _unlabelled_title_records(ctx, first_boundary):
        token = (record["role"], id(record["paragraph"]._p))
        if token not in known:
            records.append(record)
            known.add(token)
    compact = [_compact(paragraph.text) for paragraph in ctx.paragraphs]
    marker = next((index for index, value in enumerate(compact[:first_boundary])
                   if "题名和副题名" in value), None)
    title_page_start = None if marker is None else next((
        index for index in range(marker + 1, first_boundary)
        if compact[index] in {"硕士学位论文", "博士学位论文"}
    ), None)
    english_start = next((index for index, value in enumerate(compact[:first_boundary])
                          if value in {"masterdissertation", "doctoraldissertation", "phddissertation"}), None)
    for record in records:
        if record.get("area") or title_page_start is None:
            continue
        index = record["index"]
        if index < title_page_start:
            record["area"] = "cover"
        elif english_start is None or index < english_start:
            record["area"] = "title_page"
        else:
            record["area"] = "english_title_page"
    return records


def _ancestor(element, local_name: str):
    return next((item for item in element.iterancestors() if item.tag == qn("w:" + local_name)), None)


def _unsupported(paragraph) -> bool:
    return bool(paragraph._p.xpath(".//w:txbxContent | .//w:object | .//w:fldSimple | .//w:instrText"))


def _space_positioned(paragraph) -> bool:
    """段落是否靠前导空白排版（封面题目/字段常用空格代替居中或缩进）。"""

    text = paragraph.text
    return len(text) - len(text.lstrip(" \u3000\t")) >= 2


def _runs(record: dict):
    end = 0
    result = []
    for text, fmt in effective_runs(record["paragraph"]):
        begin, end = end, end + len(text)
        if end <= record.get("offset", 0):
            continue
        text = text[max(0, record.get("offset", 0) - begin):]
        if text.strip():
            result.append((text, fmt))
    return result


def _runs_filled(record: dict):
    """run 级格式：个别 run 读不到字体/字号时用整段的主导值补齐。

    封面题目、字段值常由多个 run 拼接，其中一个 run 未显式指定字体（回退到样式或
    文档默认）时，不应该把整条规则判成“无法判定”；同一行文字的有效字体本来就是
    一致的，用同段主导值补齐后照常与模板比对。
    """

    runs = _runs(record)
    if not runs:
        # 模板里字段值常是空白（只有“题 目:”这样的标签），按值取字体必然为空；
        # 这时回退到整段文字（标签与填写值本来就用同一种字体），照常与论文比对。
        runs = [(text, dict(fmt)) for text, fmt in effective_runs(record["paragraph"]) if text.strip()]
    for key in _FONT_KEYS:
        values = [fmt.get(key) for _, fmt in runs if fmt.get(key) is not None]
        if not values:
            continue
        dominant = max(set(values), key=values.count)
        for _, fmt in runs:
            if fmt.get(key) is None:
                fmt[key] = dominant
    return runs


def _profile(record: dict) -> dict:
    result = effective_paragraph(record["paragraph"])
    runs = _runs_filled(record)
    for key in _FONT_KEYS:
        selected = runs
        if key == "east_asia_font":
            selected = [item for item in runs if _CN.search(item[0])] or runs
        elif key == "ascii_font":
            selected = [item for item in runs if _LATIN.search(item[0])] or runs
        values = {item[1].get(key) for item in selected}
        result[key] = values.pop() if len(values) == 1 else None
    return result


def _geometry(paragraph) -> dict:
    result: dict[str, Any] = {}
    tc = _ancestor(paragraph._p, "tc")
    result["container"] = "表格单元格" if tc is not None else "正文段落"
    if tc is not None:
        grid_column = 0
        for sibling in tc.itersiblings(preceding=True):
            if sibling.tag == qn("w:tc"):
                span = sibling.xpath("./w:tcPr/w:gridSpan")
                grid_column += int(span[0].get(qn("w:val"))) if span else 1
        result["table_column"] = grid_column + 1
        tbl = _ancestor(tc, "tbl")
        for key, element, path in (
            ("cell_width", tc, "./w:tcPr/w:tcW"),
            ("table_indent", tbl, "./w:tblPr/w:tblInd"),
            ("table_alignment", tbl, "./w:tblPr/w:jc"),
            ("cell_vertical_alignment", tc, "./w:tcPr/w:vAlign"),
        ):
            nodes = element.xpath(path)
            if nodes:
                attrs = nodes[0].attrib
                result[key] = tuple(sorted((name.split("}")[-1], val) for name, val in attrs.items()))
    return result


def _record(paragraph, offset: int = 0) -> dict:
    return {"paragraph": paragraph, "offset": offset, "value": paragraph.text[offset:]}


def _abstracts(ctx: DocxContext) -> dict:
    regions = {}
    for index, paragraph in enumerate(ctx.paragraphs):
        title = _compact(paragraph.text)
        language = "中文" if title in {"摘要", "中文摘要"} else "英文" if title in {"abstract", "英文摘要"} else None
        if not language:
            continue
        if language in regions:
            regions[language]["ambiguous"] = True
            continue
        body, keywords = [], []
        for following in ctx.paragraphs[index + 1:]:
            keyword = re.match(r"^\s*(关键词|关键字|key\s*words?)\s*[:：]\s*", following.text, re.I)
            if keyword:
                keywords.append(_record(following, keyword.end()))
                continue
            if _boundary(following):
                break
            if following.text.strip() and not keywords:
                body.append(_record(following))
        regions[language] = {"title": _record(paragraph), "body": body, "keywords": keywords}
    return regions


def extract_front_baseline(ctx: DocxContext) -> dict:
    """Return this module's namespace; never replace existing template keys."""
    records = _fields(ctx)
    guide_cover = any("封面格式" in _compact(text) for text in ctx.texts[:100])
    for record in records:
        record["required"] = (True if record.get("area") == "cover"
                              and (record["role"] in {"title_cn", "title_en"} or guide_cover)
                              else _requirement(record, ctx))
        record["profile"] = _profile(record)
        record["geometry"] = _geometry(record["paragraph"])
    declaration = _declaration_lines(ctx)
    return {"fields": records, "abstracts": _abstracts(ctx), "has_textboxes": ctx.has_textboxes,
            "textbox_fields": _textboxes_may_hide_fields(ctx),
            # 模板声明页里实际出现过哪些行（签名/日期）；没出现过的行不检查。
            "declaration": {kind: True for kind, record in declaration.items() if record is not None}}


#: 声明页里的填写行：原创性声明（作者签名/日期）与授权声明（作者签名、导师签名、两个日期）。
_DECLARATION_SIGNATURE = re.compile(r"(?:论文作者签名|作者签名|研究生签名)\s*[:：]?")
_DECLARATION_ADVISOR = re.compile(r"(?:导师签名|指导教师签名|导师签字)\s*[:：]?")
_DECLARATION_DATE = re.compile(r"日期\s*[:：]?")
_DECLARATION_INLINE_DATE = re.compile(r"(?:\d{1,4}\s*)?年\s*(?:\d{1,2}\s*)?月\s*(?:\d{1,2}\s*)?日")
_DECLARATION_PAGE_KEYS = {
    "originality": {"signature": "signature", "date": "date"},
    "authorization": {"signature": ("authorization_signature", "authorization_advisor_signature"),
                      "date": ("authorization_date", "authorization_advisor_date")},
}
_DECLARATION_LABELS = {"signature": "作者签名", "date": "作者日期",
                       "authorization_signature": "作者签名",
                       "authorization_advisor_signature": "导师签名",
                       "authorization_date": "作者日期",
                       "authorization_advisor_date": "导师日期"}
_DECLARATION_CODES = {"signature": "020.1", "authorization_signature": "020.1",
                      "authorization_advisor_signature": "020.1",
                      "date": "020.2", "authorization_date": "020.2",
                      "authorization_advisor_date": "020.2"}


def _declaration_lines(ctx: DocxContext) -> dict:
    """定位声明页（原创性声明 + 授权声明）里的签名行与日期行。

    返回 ``{类别: {"paragraph", "value", "object", "page"}}``：

    * 一页里可能有多个标签（授权声明页是“作者签名 + 导师签名”，日期行有两个“日期：”），
      按出现顺序把每个标签的值限定在“本标签之后、下一个标签之前”；
    * 值区间里出现图片/绘图（手写签名扫描件）视为**已签名**，不再误报“签名未填写”；
    * ``page`` 取 ``originality`` / ``authorization``，用于文案里写清是哪一页。
    """

    from thesis_format_checker.core.scope import front_end_of
    from thesis_format_checker.detectors.headings import _heading_level

    front_end = front_end_of(ctx)
    active_page = None
    active_at = None
    pending: dict[str, dict] = {}
    used: dict[str, int] = {}

    def take(kind: str, paragraph, value: str, has_object: bool, start: int, end: int) -> None:
        """按出现顺序把同一页的第 1/2 个标签分给作者/导师。"""

        keys = _DECLARATION_PAGE_KEYS.get(active_page, {}).get(kind)
        if keys is None:
            return
        names = (keys,) if isinstance(keys, str) else keys
        position = used.get(kind, 0)
        if position >= len(names):
            return
        used[kind] = position + 1
        pending[names[position]] = {"paragraph": paragraph, "value": value.strip(),
                                    "object": has_object, "page": active_page,
                                    "start": start, "end": end}

    for index, paragraph in enumerate(ctx.paragraphs):
        text = paragraph.text
        if _DECLARATION_HEADING.search(text):
            active_page, active_at, used = "originality", index, {}
            continue
        if len(re.sub(r"\s+", "", text)) <= 30 and _AUTHORIZATION_HEADING.search(text):
            active_page, active_at, used = "authorization", index, {}
            continue
        if active_page is None:
            continue
        if 0 < front_end and index >= front_end and (active_at is None or active_at < front_end):
            break  # 摘要/目录开始：声明页到此结束
        if _heading_level(paragraph) and not _SIGNATURE_LABEL.search(text) and not _DATE_LABEL.search(text):
            break  # 进入下一个章节
        patterns = [
            ("signature", _DECLARATION_SIGNATURE, False),
            ("signature", _DECLARATION_ADVISOR, False),
            ("date", _DECLARATION_DATE, False),
        ]
        if not _DECLARATION_DATE.search(text):
            patterns.append(("date", _DECLARATION_INLINE_DATE, True))
        spans = sorted(
            [(match.start(), match.end(), kind, inline)
             for kind, pattern, inline in patterns for match in pattern.finditer(text)],
            key=lambda item: item[0])
        if not spans:
            continue
        has_object = bool(paragraph._p.xpath(".//w:drawing|.//w:pict|.//w:object"))
        for position, (start, end, kind, inline) in enumerate(spans):
            stop = spans[position + 1][0] if position + 1 < len(spans) else len(text)
            # 图片只算它旁边那一个标签（一行里有“作者签名 + 导师签名”时，
            # 作者签名扫描件不能让导师签名也算已填写）。
            value = text[start:stop] if inline else text[end:stop].lstrip("：: \u3000\t")
            take(kind, paragraph, value,
                 has_object and position == 0, start, stop)
    return pending


def _filled_declaration(value: str, kind: str, has_object: bool = False) -> bool:
    """签名/日期是否已填写：去掉空格与占位符后，日期还要含年份数字。"""

    if has_object:
        return True  # 手写签名扫描件（图片）视为已签名
    text = re.sub(r"[\s\u3000_＿.．·—\-]+", "", value or "")
    if not text:
        return False
    if kind == "date":
        return bool(re.search(r"\d", text))
    return True


def _author_field_checks(references: list[dict], actual: list[dict], ctx: DocxContext) -> list:
    """014：模板封面出现的字段，论文封面必须填写（空值或占位符都算缺失）。"""

    findings = []
    def cover_roles(records):
        selected = {}
        for record in records:
            if not record.get("cover"):
                continue
            role = record["role"]
            if record.get("area") == "cover" or role not in selected:
                selected[role] = record
        return selected

    template_roles = cover_roles(references)
    paper_roles = cover_roles(actual)
    fallback = next(iter(ctx.paragraphs), None)
    for role, code in _AUTHOR_FIELD_CODES.items():
        if role not in template_roles:
            continue  # 模板封面没有该字段：本条对这份模板不适用
        record = paper_roles.get(role)
        value = "" if record is None else record.get("value", "")
        if record is not None and value.strip() and not _placeholder(value, role):
            continue
        findings.append(make_finding(
            code, _title(code), f"封面应填写{_AUTHOR_FIELD_LABELS[role]}",
            f"封面未填写{_AUTHOR_FIELD_LABELS[role]}"
            + ("（缺失对应字段）" if record is None else "（字段为空或仍是占位符）"),
            (record or {}).get("paragraph") or fallback))
    return findings


def _declaration_checks(baseline: dict, ctx: DocxContext) -> list:
    """020：模板声明页有签名/日期行时，论文对应行必须填写（含授权声明的导师签名/日期）。"""

    required = baseline.get("declaration") or {}
    if not required:
        return []
    lines = _declaration_lines(ctx)
    findings = []
    fallback = next(iter(ctx.paragraphs), None)
    for key, code in _DECLARATION_CODES.items():
        if not required.get(key):
            continue
        record = lines.get(key)
        filled_kind = "signature" if key.endswith("signature") else "date"
        if record is not None and _filled_declaration(record["value"], filled_kind, record.get("object", False)):
            continue
        label = _DECLARATION_LABELS.get(key, key)
        page_name = "使用授权声明" if key.startswith("authorization") else "原创性声明"
        actual = (page_name + "缺少" + label + "行" if record is None
                  else label + "未填写（" + (record["value"].strip()[:20] or "空白") + "）")
        finding = make_finding(code, page_name + ("缺少签名" if filled_kind == "signature" else "缺少日期"),
                                     f"{page_name}应填写{label}",
                                     actual,
                                     (record or {}).get("paragraph") or fallback)
        if record:
            finding.start, finding.end = record["start"], record["end"]
        findings.append(finding)
    return findings


def _equal(left, right) -> bool:
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return abs(left - right) <= 0.02  # Word point serialization precision.
    return left == right


def _display(key, value):
    return f"{_PROPERTY_LABELS.get(key, key)}：{describe_value(key, value)}"


def _geometry_display(key, value):
    label = _PROPERTY_LABELS.get(key, "布局属性")
    if key == "table_column":
        return f"{label}：第{value}列"
    if key == "container":
        return f"{label}：{value}"
    try:
        attributes = dict(value)
        if key in {"table_alignment", "cell_vertical_alignment"}:
            translations = {"left": "左对齐", "right": "右对齐", "center": "居中", "start": "起始端对齐",
                            "end": "末端对齐", "top": "顶端对齐", "bottom": "底端对齐", "both": "两端对齐"}
            display = translations.get(attributes.get("val"), "无法解析的对齐值")
        elif attributes.get("type") == "dxa":
            display = f"{int(attributes['w']) / 20:g}磅"
        elif attributes.get("type") == "pct":
            display = f"{int(attributes['w']) / 50:g}%"
        elif attributes.get("type") == "auto":
            display = "自动"
        elif attributes.get("type") == "nil":
            display = "零宽度"
        else:
            display = "无法解析有效值"
    except (TypeError, ValueError, KeyError):
        display = "无法解析有效值"
    return f"{label}：{display}"


def _compare_format(rule: str, label: str, reference: dict, actual: dict, para_keys, font_keys,
                    codes: dict[str, str] | None = None):
    """比较段落/字体属性：**一个属性一条结论**（不再合并成一条批注）。

    ``codes`` 把属性名映射到细粒度规则编号；未映射的属性沿用 ``rule``。
    """

    paragraph = actual["paragraph"]
    if _unsupported(reference["paragraph"]) or _unsupported(paragraph):
        return [_unreadable(rule, f"{label}包含域、文本框或嵌入对象，无法读取有效格式", paragraph)]
    expected = reference.get("profile") or _profile(reference)
    current = effective_paragraph(paragraph)
    differences, missing = [], []
    for key in para_keys:
        a, b = expected.get(key), current.get(key)
        if key == "alignment" and _space_positioned(reference["paragraph"]) and _space_positioned(paragraph):
            # 靠空格排版（封面题目常见）：对齐方式不具可比性，改比首行前导空白。
            lead_a = len(reference["paragraph"].text) - len(reference["paragraph"].text.lstrip(" \u3000"))
            lead_b = len(paragraph.text) - len(paragraph.text.lstrip(" \u3000"))
            if lead_a != lead_b:
                differences.append(("alignment", f"前导空白 {lead_a} 个", f"前导空白 {lead_b} 个"))
            continue
        if a is None or b is None:
            missing.append(key)
        elif not _equal(a, b):
            differences.append((key, a, b))
    for key in font_keys:
        a = expected.get(key)
        applicable = [fmt.get(key) for text, fmt in _runs_filled(actual) if (
            _CN.search(text) if key == "east_asia_font" else _LATIN.search(text) if key == "ascii_font" else text.strip()
        )]
        if not applicable:
            continue
        if a is None or any(value is None for value in applicable):
            missing.append(key)
        for b in set(value for value in applicable if value is not None):
            if a is not None and not _equal(a, b):
                differences.append((key, a, b))
    findings = []
    seen = set()
    for key, a, b in differences:
        if key in seen:  # 同一属性（如字体在多段文字上取值不同）只报一次
            continue
        seen.add(key)
        code = (codes or {}).get(key, rule)
        # 保留“哪一部分”（如“英文摘要正文”）信息，便于定位到具体对象
        findings.append(make_finding(code, _title(code),
                                     f"{label}的{_display(key, a)}", f"{label}的{_display(key, b)}",
                                     paragraph))
    for key in dict.fromkeys(missing):
        code = (codes or {}).get(key, rule)
        findings.append(_unreadable(code, f"{label}的{_PROPERTY_LABELS.get(key, key)}无法完整确认", paragraph))
    return findings


def _missing(code: str, reference: dict, ctx: DocxContext, actual=None):
    """字段缺失：传入的应当是该字段对应的细粒度规则编号。"""

    paragraph = actual["label_paragraph"] if actual else next(iter(ctx.paragraphs), None)
    label = reference["label"]
    if reference["required"] is False:
        return []
    # 已找到明确的空字段/空题名区域时，可以直接判定缺失；文档里其他无关文本框
    # 不能推翻这个局部证据。只有整项字段都未找到且文本框可能藏有该字段时降级。
    if _unsupported(reference["paragraph"]) or (
            actual is None and _textboxes_may_hide_fields(ctx)):
        return [_unreadable(code, f"{label}可能位于不可读取的文本框或域内，不能认定缺失", paragraph)]
    if reference["required"] is True:
        return [make_finding(code, _title(code), f"模板要求填写{label}", "缺失、空白或仍为占位内容", paragraph)]
    reason = "模板必填与选填说明冲突" if reference["required"] == "conflict" else "模板未明确必填要求"
    return [unknown(code, _title(code), f"{label}未填写，{reason}", paragraph)]


def _date_pattern(value: str):
    text = value.strip().strip("〈〉<>【】[]（）()")
    tokens = r"(?:[0-9]{1,4}|[Yy]{4}|[MmDd]{1,2}|[_＿]{1,4})"
    m = re.fullmatch(rf"({tokens})(年|[-./])({tokens})(月|[-./])?({tokens})?(日)?", text)
    if not m:
        return None
    if bool(m.group(5)) != bool(m.group(4) and m.group(4) != "月") and m.group(2) != "年":
        return None
    segments = []
    for i, token in enumerate(m.groups()):
        if token is None:
            continue
        if i in (0, 2, 4):
            width = len(token) if re.fullmatch(r"[YyMmDd]+", token) else (4 if i == 0 else 2 if token.startswith("0") else None)
            segments.append(r"\d{" + str(width) + "}" if width else r"\d{1,2}")
        else:
            segments.append(re.escape(token))
    return "".join(segments)


def _date_valid(value: str, pattern: str) -> bool:
    if not re.fullmatch(pattern, value.strip()):
        return False
    numbers = [int(item) for item in re.findall(r"\d+", value)]
    try:
        date(numbers[0], numbers[1], numbers[2] if len(numbers) > 2 else 1)
    except (ValueError, IndexError):
        return False
    return True


def _title_checks(reference_fields, actual_fields, ctx):
    findings = []
    for role in ("title_cn", "title_en"):
        references = [f for f in reference_fields if f["role"] == role and f["cover"]]
        if not references:
            continue
        targets = [f for f in actual_fields if f["role"] == role]
        area_aware_reference = any("area" in item for item in references)
        area_aware_actual = any("area" in item for item in targets)
        reference = (next((f for f in references if f.get("area") == "cover"), None)
                     if area_aware_reference else None) or references[0]
        cover = (next((f for f in targets if f.get("area") == "cover"), None)
                 if area_aware_actual else next((f for f in targets if f["cover"]), None))
        if cover is None or _placeholder(cover["value"], role):
            findings += _missing("012.1", reference, ctx, cover)
            continue
        findings += _compare_format("013.3", reference["label"], reference, cover,
                                    ("line_spacing", "line_spacing_rule"), _FONT_KEYS,
                                    codes=_COVER_TITLE_CODES)
        for target in targets:
            if target is cover or _placeholder(target["value"], role):
                continue
            # 封面区域内的重复题目 → 012.2；正文/摘要中的题目 → 012.3
            code = "012.2" if target["cover"] else "012.3"
            if _unsupported(target["paragraph"]) or _unsupported(cover["paragraph"]):
                findings.append(_unreadable(code, "题目含无法可靠读取的域或嵌入对象", target["paragraph"]))
            # Collapse layout whitespace only.  Preserve meaningful internal
            # spaces in English titles and all punctuation/subtitle text.
            elif _title_value(target["value"], role) != _title_value(cover["value"], role):
                findings.append(make_finding(code, _title(code),
                    f"模板中同语种题目字段应一致；本论文封面题目为“{cover['value']}”",
                    f"本处题目为“{target['value']}”", target["paragraph"]))
    return findings


def _title_value(value: str, role: str) -> str:
    if role == "title_cn":
        value = re.sub(r"[\r\n\v]+", "", value)
    return " ".join(value.split())


def _field_checks(reference_fields, actual_fields, ctx):
    findings = []
    covered = set()
    for reference in reference_fields:
        role = reference["role"]
        rule = "015" if role.startswith("advisor_") else "016" if role.endswith("_date") else "019" if role in {"secrecy", "classification", "udc"} else None
        if not rule or not reference["cover"] or role in covered:
            continue
        covered.add(role)
        missing_code, format_code = _field_codes(role)
        targets = [f for f in actual_fields if f["role"] == role]
        if reference.get("area") == "cover" and any("area" in f for f in targets):
            targets = [f for f in targets if f.get("area") == "cover"]
        if not targets:
            findings += _missing(missing_code, reference, ctx)
            continue
        comparable_references = [f for f in reference_fields if f["role"] == role and f["cover"]]
        if reference.get("area"):
            comparable_references = [f for f in comparable_references
                                     if f.get("area") == reference.get("area")]
        if len(comparable_references) > 1:
            findings.append(unknown(missing_code, _title(missing_code),
                f"模板中存在多个{reference['label']}字段，无法确认导师角色或日期/位置的逐一对应关系",
                targets[0]["paragraph"]))
            continue
        for actual in targets:
            if _placeholder(actual["value"], role):
                findings += _missing(missing_code, reference, ctx, actual)
                continue
            if _unsupported(reference["paragraph"]) or _unsupported(actual["paragraph"]):
                findings.append(_unreadable(format_code, "字段包含无法读取的域、文本框或OLE对象", actual["paragraph"]))
                continue
            if rule == "015":
                findings += _compare_format(format_code, reference["label"], reference, actual, _PARA_KEYS, _FONT_KEYS)
            elif rule == "016":
                pattern = _date_pattern(reference["value"])
                if pattern is None:
                    findings.append(unknown(format_code, _title(format_code), f"模板的{reference['label']}未提供可解析的日期形式", actual["paragraph"]))
                elif not _date_valid(actual["value"], pattern):
                    findings.append(make_finding(format_code, _title(format_code),
                        f"{reference['label']}采用“{reference['value']}”所示日期形式（不比较具体日期）",
                        actual["value"], actual["paragraph"]))
            else:
                if not actual["cover"]:
                    findings.append(make_finding(format_code, _title(format_code), f"{reference['label']}位于模板封面区域",
                                                  "当前字段位于封面之后", actual["paragraph"]))
                findings += _compare_geometry(format_code, reference, actual)
    # 同段并排字段（常见“分类号 … 密级 …”）的左右顺序也属于位置要求。
    def cover_position(records, role):
        candidates = [record for record in records if record["role"] == role and record.get("cover")]
        if any(record.get("area") == "cover" for record in candidates):
            candidates = [record for record in candidates if record.get("area") == "cover"]
        if not candidates:
            return None
        record = candidates[0]
        return (record["index"], record.get("offset", 0), record)

    order_roles = ("classification", "secrecy", "udc")
    expected_positions = {role: cover_position(reference_fields, role) for role in order_roles}
    actual_positions = {role: cover_position(actual_fields, role) for role in order_roles}
    emitted = False
    for left_index, left in enumerate(order_roles):
        for right in order_roles[left_index + 1:]:
            expected_left, expected_right = expected_positions[left], expected_positions[right]
            actual_left, actual_right = actual_positions[left], actual_positions[right]
            if not all((expected_left, expected_right, actual_left, actual_right)):
                continue
            expected_before = expected_left[:2] < expected_right[:2]
            actual_before = actual_left[:2] < actual_right[:2]
            if expected_before != actual_before:
                label_left = _LABELS[left][0]
                label_right = _LABELS[right][0]
                findings.append(make_finding(
                    "019.4", _title("019.4"),
                    f"{label_left}与{label_right}的左右顺序应与模板一致",
                    f"当前{label_left}与{label_right}的左右顺序相反",
                    actual_left[2]["paragraph"],
                ))
                emitted = True
                break
        if emitted:
            break
    return findings


def _compare_geometry(rule, reference, actual, codes: dict[str, str] | None = None):
    """表格/段落位置属性比较：一个属性一条结论。"""

    findings = []
    expected, current = reference["geometry"], _geometry(actual["paragraph"])
    changes = [(key, value, current[key]) for key, value in expected.items()
               if key in current and current[key] != value]
    unavailable = [key for key in expected if key not in current]
    if changes:
        for key, a, b in changes:
            code = (codes or {}).get(key, rule)
            findings.append(make_finding(code, _title(code), _geometry_display(key, a),
                                         _geometry_display(key, b), actual["paragraph"]))
    if unavailable:
        findings.append(_unreadable(rule,
            "表格位置属性无法解析或依赖未解析的表格样式：" + "、".join(_PROPERTY_LABELS.get(key, key) for key in unavailable), actual["paragraph"]))
    if rule.startswith("019"):
        findings.append(_unreadable("019",
            f"{reference['label']}的区域和表格位置属性已检查；python-docx无法读取字段在页面上的绝对位置",
            actual["paragraph"]))
    return findings


def _layout_checks(reference_fields, actual_fields):
    findings = []
    for reference in reference_fields:
        if not reference["cover"]:
            continue
        if sum(1 for item in reference_fields if item["cover"] and item["role"] == reference["role"]) > 1:
            continue
        targets = [f for f in actual_fields if f["role"] == reference["role"] and f["cover"]]
        if len(targets) != 1:
            if len(targets) > 1:
                findings.append(unknown("018.1", _title("018.1"),
                    f"{reference['label']}有多个候选，无法确定布局对应关系", targets[0]["paragraph"]))
            continue
        actual = targets[0]
        findings += _compare_format("018.1", reference["label"], reference, actual,
                                    ("alignment", "left_indent", "right_indent", "space_before", "space_after"), (),
                                    codes=_LAYOUT_CODES)
        findings += _compare_geometry("018.3", reference, actual, codes=_LAYOUT_CODES)
    if any(f["cover"] for f in reference_fields):
        anchor = next((f["paragraph"] for f in actual_fields if f["cover"]), None)
        findings.append(_unreadable("018",
            "段落和表格布局属性已检查；python-docx无法排版，无法读取封面元素的实际页面坐标和渲染后间距", anchor))
    return findings


def _abstract_checks(baseline, ctx):
    findings = []
    actual = _abstracts(ctx)
    for language, reference in baseline["abstracts"].items():
        current = actual.get(language)
        if current is None:
            findings.append(unknown("021.1", _title("021.1"), f"未找到可读取的{language}摘要区域，无法比较格式", next(iter(ctx.paragraphs), None)))
            continue
        if reference.get("ambiguous") or current.get("ambiguous"):
            findings.append(unknown("021.1", _title("021.1"), f"存在多个{language}摘要候选，无法确定对应关系", current["title"]["paragraph"]))
            continue
        # 摘要标题 / 摘要正文 / 关键词 各自成一条规则
        findings += _compare_format("021.1", language + "摘要标题", reference["title"], current["title"],
                                    _PARA_KEYS, _FONT_KEYS)
        for area in ("body", "keywords"):
            expected_records, current_records = reference[area], current[area]
            label = language + ("摘要正文" if area == "body" else "关键词")
            code = "021.2" if area == "body" else "021.3"
            if not expected_records or not current_records:
                findings.append(unknown(code, _title(code), f"{label}缺少模板或文档有效内容，无法比较格式", current["title"]["paragraph"]))
                continue
            # Multiple template paragraphs may legitimately have distinct
            # formats.  Only common properties are safe to generalize.
            profiles = [_profile(record) for record in expected_records]
            profile = {key: profiles[0].get(key) if all(p.get(key) == profiles[0].get(key) for p in profiles) else None
                       for key in _PARA_KEYS + _FONT_KEYS}
            template_record = dict(expected_records[0], profile=profile)
            for record in current_records:
                findings += _compare_format(code, label, template_record, record, _PARA_KEYS, _FONT_KEYS)
    return findings


def check_front_errors(baseline: dict, ctx: DocxContext):
    """Check only 012/013/015/016/018/019/021; never alter either document."""
    baseline = baseline.get("front", baseline)
    findings = []
    try:
        actual = _fields(ctx)
        references = baseline.get("fields", [])
        operations = (
            (("012", "013"), lambda: _title_checks(references, actual, ctx)),
            (("015", "016", "019"), lambda: _field_checks(references, actual, ctx)),
            (("014",), lambda: _author_field_checks(references, actual, ctx)),
            (("020",), lambda: _declaration_checks(baseline, ctx)),
            (("018",), lambda: _layout_checks(references, actual)),
            (("021",), lambda: _abstract_checks(baseline, ctx)),
        )
        for rules, operation in operations:
            try:
                findings.extend(operation())
            except Exception as exc:
                for rule in rules:
                    findings.append(_unreadable(rule, f"读取前置部分对象失败（{type(exc).__name__}）", next(iter(ctx.paragraphs), None)))
        roles = {f["role"] for f in references if f["cover"]}
        supported = {
            "012": bool(roles & {"title_cn", "title_en"}), "013": bool(roles & {"title_cn", "title_en"}),
            "015": any(r.startswith("advisor_") for r in roles), "016": any(r.endswith("_date") for r in roles),
            "018": bool(roles), "019": bool(roles & {"secrecy", "classification", "udc"}),
            "021": bool(baseline.get("abstracts")),
            "014": any(role in roles for role in _AUTHOR_FIELD_CODES),
            "020": bool(baseline.get("declaration")),
        }
        for rule, exists in supported.items():
            if not exists:
                findings.append(unknown(rule, _FRONT_RULES[rule], "模板未提供可可靠识别的对应字段或区域基准", next(iter(ctx.paragraphs), None)))
        if _textboxes_may_hide_fields(ctx) or (
            baseline.get("has_textboxes") and baseline.get("textbox_fields")
        ):
            for rule in _FRONT_RULES:
                findings.append(_unreadable(rule, "模板或论文含文本框，其内部前置字段、位置及格式无法由python-docx可靠读取", next(iter(ctx.paragraphs), None)))
    except Exception as exc:
        for rule in _FRONT_RULES:
            findings.append(_unreadable(rule, f"前置部分读取异常（{type(exc).__name__}）", next(iter(ctx.paragraphs), None)))
    return findings
