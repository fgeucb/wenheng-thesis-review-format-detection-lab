"""需要 Word 排版几何才能确定的规则：078 题注位置、087 跨页表格、033/098 越界。

设计方案来自同门的 ``readers/layout_snapshot.py``：用 Word 读段落/图形/表格的真实
坐标与分页，**不以 XML 存储顺序冒充视觉位置**；坐标读不到时保留“无法判定”，
绝不猜一个结论。
"""
from __future__ import annotations

import re

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, make_finding, unknown
from thesis_format_checker.core.models import Status

_TITLES = {"033": "内容超出可打印区域存在裁切风险",
           "078": "图表题注相对位置不符合模板",
           "087": "跨页表格未重复标题行",
           "098": "长公式未换行超出页边距"}

#: 需要 Word **排版几何**（坐标/分页）才能判定的规则族。
GEO_LAYOUT_FAMILIES = frozenset({"078", "087", "033", "098", "059"})

#: 只需要 Word **分页事实**（render_facts）、不需要几何坐标的规则族。
FACT_ONLY_FAMILIES = frozenset({"032", "165"})


def _title(code: str) -> str:
    item = FINE_RULES.get(code)
    return item.title if item else _TITLES[code.split(".")[0]]


def _notice(code: str, reason: str, paragraph=None):
    return unknown(code, _title(code), UNREADABLE + "；" + reason, paragraph=paragraph)


# ---------------------------------------------------------------------------
# 078 题注上下位置：用坐标判定（浮动/嵌入对象也能判）
# ---------------------------------------------------------------------------

def _box_of(layout: dict, index: int):
    item = (layout.get("paragraphs") or {}).get(index) or {}
    return item.get("box")


def _object_box(ctx, layout: dict, record: dict):
    """题注对应对象的矩形：表格取表格框，图片取该段内图片/浮动图形的并集。"""

    obj = record.get("object")
    if obj is None:
        return None
    if record.get("kind") == "table":
        caption_index = ctx.index(record["paragraph"])
        table = (layout.get("tables") or {}).get(caption_index)
        if table and table.get("box"):
            return table["box"]
        # 表格宿主段落可能是紧邻的上一段（表题在表格上方）
        for index in (caption_index, caption_index + 1):
            table = (layout.get("tables") or {}).get(index)
            if table and table.get("box"):
                return table["box"]
        return None
    boxes = []
    for index, facts in (layout.get("figures") or {}).items():
        element = ctx.paragraphs[index]._p
        if element is obj or obj in list(element.iter()):
            boxes.extend(box for box in facts.get("boxes", []) if box)
    if not boxes:
        return None
    first = min(boxes, key=lambda box: (box["page"], box["y"]))
    last = max(boxes, key=lambda box: (box["end_page"], box["bottom"]))
    return {"page": first["page"], "end_page": last["end_page"], "y": first["y"],
            "bottom": last["bottom"], "x": first["x"],
            "width": max(box.get("width", 0.0) for box in boxes)}


def _actual_position(ctx, layout: dict, record: dict):
    """坐标判定题注在对象上方还是下方；任一坐标缺失返回 None。"""

    caption = _box_of(layout, ctx.index(record["paragraph"]))
    target = _object_box(ctx, layout, record)
    if not caption or not target:
        return None
    if caption["page"] != target["page"]:
        return "below" if caption["page"] > target["page"] else "above"
    return "below" if caption["y"] > target["y"] else "above"


def _caption_positions(baseline: dict, ctx, layout: dict) -> list:
    from thesis_format_checker.detectors import figures as figures_module

    data = baseline.get("figures") or {}
    findings = []
    groups = (data.get("figures") or {}).get("groups") or {}
    for record in figures_module.collect_caption_records(ctx):
        candidates = groups.get(figures_module._key(record), [])
        static = figures_module._position_check(record, candidates)
        if not static:
            continue
        if all(item.status != Status.UNKNOWN for item in static):
            findings.extend(static)          # 静态已能确定：原样保留
            continue
        position = _actual_position(ctx, layout, record)
        expected = next((item["position"] for item in candidates
                         if item.get("position") is not None), None)
        if position is None or expected is None:
            findings.extend(static)          # 坐标也读不到：保留“无法判定”
            continue
        if position == expected:
            continue                          # 坐标核对通过：不再报“无法判定”
        code = "078.2" if re.match(r"^\s*(?:表|Table)", record["text"], re.I) else "078.1"
        names = {"above": "图表对象上方", "below": "图表对象下方"}
        findings.append(make_finding(
            code, _title(code), names[expected], names[position], record["paragraph"]))
    return findings


# ---------------------------------------------------------------------------
# 087 跨页表格：同一原生表格跨页时标题行必须重复
# ---------------------------------------------------------------------------

def _table_pagination(ctx, layout: dict) -> list:
    findings = []
    for index, fact in sorted((layout.get("tables") or {}).items()):
        pages = fact.get("pages") or (None, None)
        if pages[0] is None or pages[1] is None or pages[0] == pages[1]:
            continue
        paragraph = ctx.paragraphs[index] if 0 <= index < len(ctx.paragraphs) else None
        if fact.get("floating"):
            findings.append(_notice("087.1", "表格为浮动定位，跨页行为需人工核对", paragraph))
            continue
        headers = [value for value in (fact.get("headers") or []) if value is not None]
        if len(headers) <= 1 or not any(headers[1:]):
            findings.append(make_finding(
                "087.1", _title("087.1"),
                "跨页表格应重复标题行（首行设为“重复标题行”）",
                f"该表跨第 {pages[0]}–{pages[1]} 页，但后续行没有设置为重复标题行",
                paragraph))
    return findings


# ---------------------------------------------------------------------------
# 033 / 098 越界：用坐标核对可打印区域
# ---------------------------------------------------------------------------

def _overflow(ctx, layout: dict) -> list:
    """图形/表格/公式越出可打印区域。坐标或页面设置读不到时不报。"""

    findings = []
    setup = None
    section = ctx.document.sections[0] if ctx.document.sections else None
    if section is not None:
        setup = section
    if setup is None:
        return findings
    # python-docx 的长度是 EMU，Word 坐标是磅：统一换算成磅再比较。
    emu_to_point = 1 / 12700
    width = (float(setup.page_width) - float(setup.left_margin)
             - float(setup.right_margin)) * emu_to_point
    right_edge = (float(setup.page_width) - float(setup.right_margin)) * emu_to_point
    for index, facts in sorted((layout.get("figures") or {}).items()):
        paragraph = ctx.paragraphs[index] if 0 <= index < len(ctx.paragraphs) else None
        for box in facts.get("boxes", []):
            if not box or not box.get("width"):
                continue
            if box["x"] + box["width"] > right_edge + 1.0 or box.get("width") > width + 1.0:
                findings.append(make_finding(
                    "033.1", _title("033.1"),
                    f"内容应完全落在可打印区域内（版心宽 {width / 28.35:.1f} 毫米）",
                    f"图片右边缘超出可打印区（宽 {box['width'] / 28.35:.1f} 毫米，"
                    f"起点 x={box['x'] / 28.35:.1f} 毫米）", paragraph))
                break
    return findings


def check_layout_errors(baseline: dict, ctx, layout: dict) -> list:
    """用排版几何复核 078 / 087 / 033：能确定就给结论，读不到就保留无法判定。"""

    findings = []
    findings.extend(_caption_positions(baseline, ctx, layout))
    findings.extend(_table_pagination(ctx, layout))
    findings.extend(_overflow(ctx, layout))
    findings.extend(_toc_pages(ctx, layout))
    findings.extend(_equation_overflow(ctx, layout))
    return findings


def check_fact_errors(baseline: dict, ctx) -> list:
    """只用 Word 分页事实（``render_facts``）就能判定的规则：032/165.1 空白页。

    这类规则不需要几何坐标，因此**不必打开 Word 读排版**——分页事实在 Word
    有效格式流程里已经顺手取到（见 ``core/word_pages.read_page_facts``）。
    """

    return _blank_pages(ctx, getattr(ctx, "render_facts", None) or {})


# ---------------------------------------------------------------------------
# 静态先筛（性能优化第 3 条）：先判断“哪些规则真的需要 Word 排版”
# ---------------------------------------------------------------------------

def _figure_indices(ctx) -> set:
    """含图片/浮动图形/OLE 对象的段落（033 越界检测对象）。"""

    return {index for index, paragraph in enumerate(ctx.paragraphs)
            if paragraph._p.xpath(".//w:drawing|.//w:pict|.//w:object")}


def _caption_indices(ctx) -> set:
    """图题/表题所在段落；表题还要看紧邻的下一段（表题在表格上方的排版）。"""

    try:
        from thesis_format_checker.detectors.figures import collect_caption_records
    except Exception:
        return set()
    indices: set = set()
    try:
        records = collect_caption_records(ctx)
    except Exception:
        return set()
    for record in records:
        index = ctx.index(record["paragraph"])
        if index < 0:
            continue
        indices.add(index)
        if record.get("kind") == "table":
            indices.add(index + 1)
    return indices


def _equation_indices(ctx) -> set:
    """公式段落（098 越界检测对象）。"""

    try:
        from thesis_format_checker.detectors.equations import _read_records

        records, _ = _read_records(ctx)
    except Exception:
        return set()
    return {ctx.index(record["paragraph"]) for record in records}


def _heading_indices(ctx) -> set:
    """标题段落（059 目录页码比对需要它们所在页的打印页码）。"""

    try:
        from thesis_format_checker.detectors.headings import collect_heading_records

        records = collect_heading_records(ctx)
    except Exception:
        return set()
    return {ctx.index(record["paragraph"]) for record in records}


def _table_host_indices(ctx) -> set:
    """紧邻原生表格的段落（087 跨页表格检测的入口）。"""

    from thesis_format_checker.core.xmlutil import qn

    hosts: set = set()
    for index, paragraph in enumerate(ctx.paragraphs):
        parent = paragraph._p.getparent()
        if parent is not None and any(node.tag == qn("w:tbl") for node in parent):
            hosts.add(index)
    return hosts


def layout_requirements(ctx, findings) -> dict | None:
    """静态先筛：哪些段落/几何真的会被用到；都不用时返回 ``None``（不读排版）。

    只有“静态阶段给出无法判定、且结论确实依赖排版几何”的规则族才需要读 Word
    坐标；其余情况（例如全文只有空白页结论）完全不打开 Word，从而省掉整趟排版遍历。

    087（跨页表格）与 098（公式越界）没有静态取证来源——OOXML 里既读不出“跨了几页”，
    也读不出“右边缘在哪”，所以只要文档里**真的存在**表格/公式，就按“需要量一次坐标”
    处理；文档里没有这些对象时不请求，避免为了不存在的对象打开 Word。

    Returns:
        ``{"indices": {...}, "want_figures": bool, "want_tables": bool}`` 或 ``None``。
    """

    unknown_families = {finding.rule_id.split(".")[0] for finding in findings
                        if finding.status == Status.UNKNOWN}
    wanted = set(unknown_families & GEO_LAYOUT_FAMILIES)
    table_hosts = _table_host_indices(ctx)
    if table_hosts:
        wanted.add("087")
    equation_indices = _equation_indices(ctx)
    if equation_indices:
        wanted.add("098")
    if not wanted:
        return None
    indices: set = set(table_hosts) | set(equation_indices)
    want_figures = False
    want_tables = False
    if "078" in wanted:
        # 题注位置：题注段落自身的框 + 图片/表格对象的框。
        indices |= _caption_indices(ctx)
        want_figures = True
        want_tables = True
    if "033" in wanted:
        # 越界：图片/浮动对象矩形。
        indices |= _figure_indices(ctx)
        want_figures = True
    if "098" in wanted:
        # 公式越界：公式段落的行框。
        indices |= equation_indices
    if "059" in wanted:
        # 目录页码：标题段落的实际打印页码。
        indices |= _heading_indices(ctx)
    if "087" in wanted:
        # 跨页表格：表格分页与标题行设置。
        want_tables = True
    return {"indices": indices, "want_figures": want_figures, "want_tables": want_tables}


def content_layout_request(ctx) -> dict | None:
    """**只看文档对象**判断"可能要读哪些排版几何"（不看规则结论）。

    用途：把几何读取合并进旧版 Word 流程的同一次会话。调度层在打开 Word 之前
    先问一次本函数，只要文档里存在题注/图片/公式/表格/标题，就顺手把坐标读出来，
    之后不必再为几何单独打开一次 Word（省一次打开 + 重排，样例论文上约 1–3 分钟）。

    返回结构与 :func:`layout_requirements` 一致，可直接传给
    ``core.word_layout.read_layout_in_session``。
    """

    indices: set = set()
    want_figures = False
    want_tables = False
    captions = _caption_indices(ctx)
    if captions:
        indices |= captions
        want_figures = True
        want_tables = True
    figures = _figure_indices(ctx)
    if figures:
        indices |= figures
        want_figures = True
    equations = _equation_indices(ctx)
    if equations:
        indices |= equations
    tables = _table_host_indices(ctx)
    if tables:
        indices |= tables
        want_tables = True
    if _heading_indices(ctx):
        # 059（目录页码）需要标题所在页的打印页码。
        indices |= _heading_indices(ctx)
    if not indices:
        return None
    return {"indices": indices, "want_figures": want_figures, "want_tables": want_tables}


def _equation_overflow(ctx, layout: dict) -> list:
    """098：公式段落右边缘超出可打印区（用排版坐标，不用字符估算）。"""

    section = ctx.document.sections[0] if ctx.document.sections else None
    if section is None:
        return []
    right_edge = (float(section.page_width) - float(section.right_margin)) / 12700
    try:
        from thesis_format_checker.detectors.equations import _read_records

        records, _ = _read_records(ctx)
    except Exception:
        return []
    findings = []
    for record in records:
        index = ctx.index(record["paragraph"])
        box = ((layout.get("paragraphs") or {}).get(index) or {}).get("box")
        if not box or box.get("right") is None:
            continue
        if box["right"] > right_edge + 1.0:
            findings.append(make_finding(
                "098.1", _title("098.1"),
                f"公式应完全落在可打印区域内（版心右边缘 {right_edge / 28.35:.1f} 毫米处）",
                f"该公式右边缘在 {box['right'] / 28.35:.1f} 毫米处，超出可打印区",
                record["paragraph"]))
    return findings


def _blank_pages(ctx, facts: dict) -> list:
    """032/165.1：没有正文内容的空白页；文末连续空白页按 165.1 单独报。"""

    total = (facts or {}).get("total_pages")
    if not total:
        return []
    # 新版分页事实直接给出"空白页/有内容的页"（按页范围判定，跨页长段落不再误判）；
    # 旧格式（只有 pages 映射）按"该页有没有段落起始"推断，保持兼容。
    blank = (facts or {}).get("blank_pages")
    content = list((facts or {}).get("content_pages") or [])
    if blank is None:
        pages = {page for page in ((facts or {}).get("pages") or {}).values() if page}
        if not pages:
            return []
        content = sorted(pages)
        blank = [number for number in range(1, int(total) + 1) if number not in pages]
    if not blank:
        return []
    last_content = max(content) if content else 0
    anchor = ctx.paragraphs[-1] if ctx.paragraphs else None
    findings = []
    trailing = [number for number in blank if number > last_content]
    middle = [number for number in blank if number <= last_content]
    if middle:
        findings.append(make_finding(
            "032.1", _title("032.1"), "文档中不应出现非必要的空白页",
            "第 " + "、".join(str(number) for number in middle[:8])
            + " 页没有任何正文内容，请确认是否为多余空白页", anchor))
    if trailing:
        findings.append(make_finding(
            "165.1", _title("165.1"), "文档末尾不应保留多余空白页",
            "第 " + "、".join(str(number) for number in trailing[:8])
            + " 页在正文结束之后没有任何内容，请删除末尾空白页", anchor))
    return findings


# ---------------------------------------------------------------------------
# 059 目录页码：条目缓存的页码 vs 正文标题的实际打印页码
# ---------------------------------------------------------------------------

def _printed_number(text: str):
    """把目录里的页码文字（1 / I / 一）解析成数字；解析不了返回 None。"""

    from thesis_format_checker.detectors.headings import _token_value

    token = re.sub(r"[^0-9０-９IVXLCDMivxlcdm零〇一二两三四五六七八九十百千]", "", text or "")
    if not token:
        return None
    token = token.translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    value, _ = _token_value(token)
    return value


def _toc_pages(ctx, layout: dict) -> list:
    """059：目录条目缓存页码与正文标题实际页码不一致时报告。"""

    from thesis_format_checker.detectors import toc as toc_module
    from thesis_format_checker.detectors.headings import collect_heading_records

    try:
        groups = toc_module._snapshot(ctx)["groups"]
    except Exception:
        return []
    if not groups:
        return []
    by_title: dict[str, list] = {}
    for record in collect_heading_records(ctx):
        by_title.setdefault(toc_module._norm(record["title"]), []).append(record)
    findings = []
    for group in groups:
        for entry in group["entries"]:
            cached = _printed_number(entry.get("page"))
            if cached is None:
                continue
            matches = by_title.get(entry["label"]) or by_title.get(toc_module._norm(
                re.sub(r"^(?:第[0-9一二三四五六七八九十]+[章节]|[0-9.]+)", "", entry["label"])))
            if not matches or len(matches) > 1:
                continue
            index = ctx.index(matches[0]["paragraph"])
            facts = (layout.get("paragraphs") or {}).get(index) or {}
            actual = facts.get("printed_number")
            if actual is None or actual == cached:
                continue
            findings.append(make_finding(
                "059", _title("059"),
                f"目录中“{entry['label']}”的页码应为 {cached}",
                f"正文该标题实际在第 {actual} 页（目录缓存 {entry.get('page')}）",
                matches[0]["paragraph"]))
    return findings
