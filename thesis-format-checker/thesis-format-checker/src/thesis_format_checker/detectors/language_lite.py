"""不需要大模型的文字规范检查（本阶段实现）：

* **144.1 常见错别字/易混词**：固定词表（“做为→作为”“必竟→毕竟”这类无歧义错写）；
* **144.2 重复字词**：功能词叠字（“的的”“了了”“是是”…），白名单内的正常叠词不报；
* **149.1 缩略语首次出现未给全称**：首次出现的英文缩略语附近没有中文全称/英文全称（提醒）；
* **151.1 年份或序数的数字写法混用**：同一文档里“2026年”与“二〇二六年”等写法并存（提醒）；
* **119.1 英文文献作者姓名写法不统一**：参考文献里有的全大写、有的首字母大写（警告）。

设计原则与全项目一致：**宁可漏报，不误报**。词表只收“无歧义错写”；缩略语、数字
写法两类只给“提醒”，并在文案里说明是“写法统一”问题，不判定对错；证据不足时不输出。
"""
from __future__ import annotations

import re
from collections import Counter, defaultdict

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import UNREADABLE, make_finding, unknown
from thesis_format_checker.core.models import Status

_TITLES = {"119": "英文文献大小写或期刊缩写不统一",
           "144": "存在错别字、形近字误用、漏字或重复字",
           "149": "缩略语首次出现未给出全称或说明",
           "151": "数字、量词、年代、世纪或序数词表达不规范"}

#: 无歧义的常见错写 → 正确写法（学术写作里左边写法基本可以判定是笔误）。
_COMMON_TYPOS = {
    "做为": "作为", "按排": "安排", "按装": "安装", "报歉": "抱歉", "必竟": "毕竟",
    "辩认": "辨认", "辩别": "辨别", "部份": "部分", "材华": "才华", "粗旷": "粗犷",
    "措败": "挫败", "幅射": "辐射", "供献": "贡献", "关健": "关键", "广乏": "广泛",
    "急待": "亟待", "既使": "即使", "坚苦": "艰苦", "竞然": "竟然", "决对": "绝对",
    "拉笼": "拉拢", "伶听": "聆听", "漏习": "陋习", "媒界": "媒介", "明片": "名片",
    "年令": "年龄", "披漏": "披露", "却实": "确实", "溶入": "融入", "上叙": "上述",
    "审请": "申请", "事绩": "事迹", "书藉": "书籍", "题纲": "提纲", "通霄": "通宵",
    "统畴": "统筹", "透澈": "透彻", "图径": "途径", "推委": "推诿", "无耐": "无奈",
    "显署": "显著", "消毁": "销毁", "严竣": "严峻", "演译": "演绎", "涌跃": "踊跃",
    "专研": "钻研", "座落": "坐落", "作梦": "做梦", "帐号": "账号",
    "旁证博引": "旁征博引", "奇型怪状": "奇形怪状", "轻而意举": "轻而易举",
    "完壁归赵": "完璧归赵", "妄自匪薄": "妄自菲薄", "相形见拙": "相形见绌",
    "一笔勾消": "一笔勾销", "一慨而论": "一概而论", "一如继往": "一如既往",
    "以逸代劳": "以逸待劳", "应接不遐": "应接不暇", "有条不稳": "有条不紊",
    "与日具增": "与日俱增", "欲盖弥障": "欲盖弥彰", "再接再励": "再接再厉",
    "责无旁代": "责无旁贷", "支离破粹": "支离破碎", "执迷不误": "执迷不悟",
    "中流抵柱": "中流砥柱", "自抱自弃": "自暴自弃", "走头无路": "走投无路",
    "拔山涉水": "跋山涉水", "百尺杆头": "百尺竿头", "并行不背": "并行不悖",
    "不径而走": "不胫而走", "不加思索": "不假思索", "不落巢臼": "不落窠臼",
    "出奇不意": "出其不意", "度过难关": "渡过难关", "登陆系统": "登录系统",
}

#: 功能词叠字：这些字重复出现基本是笔误；正常叠词（看看、常常、渐渐…）不在表内。
_BAD_DOUBLES = ("的的", "了了", "是是", "和和", "在在", "我我", "你你", "他他", "她她",
                "它它", "有有", "就就", "也也", "很很", "都都", "而而", "与与", "及及",
                "为为", "以以", "从从", "并并", "或或", "但但", "则则", "用用", "把把",
                "被被", "让让", "使使", "得得", "地地", "们们", "向向", "于于", "由由",
                "对对于", "关关于")

_ACRONYM = re.compile(r"(?<![A-Za-z0-9/])([A-Z][A-Z0-9]{1,5})(?![A-Za-z0-9/])")
#: 纯罗马数字（I/II/III/IV…）：多为图表编号、章节序号或页码，不按缩略语提醒。
_ROMAN_NUMERAL = re.compile(r"M{0,4}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3})")
#: 缩略语旁边出现中文全称或英文全称的写法（含中英文括号、冒号、逗号等分隔）。
_FULL_NAME = re.compile(r"[\u4e00-\u9fff]{2,}\s*[（(]\s*$|^[\u4e00-\u9fff]{2,}[（(]|"
                        r"[（(][A-Za-z][A-Za-z\-\s]{3,}[）)]")
_YEAR_CN = re.compile(r"[〇零一二三四五六七八九]{4}年")
_YEAR_DIGIT = re.compile(r"(?<!\d)(?:1[89]\d{2}|20\d{2})年")
_ORDINAL_CN = re.compile(r"第[〇零一二三四五六七八九十]{1,3}[章节条款]")
_ORDINAL_DIGIT = re.compile(r"第\d{1,3}[章节条款]")
_REFERENCE_HEADING = re.compile(r"^参考文献$|^references?$", re.I)
_AUTHOR_CAPS = re.compile(r"\b[A-Z][A-Z'\-]{1,}(?:\s+[A-Z]){1,3}\b")
_AUTHOR_TITLE = re.compile(r"\b[A-Z][a-z'\-]{1,}(?:\s+[A-Z]\.?){1,3}\b")


def extract_language_baseline(ctx):
    """文字规范检查不依赖模板基准；这里只登记一条“已读取正文”的说明。"""

    return {"language_lite": {"paragraphs": len(ctx.paragraphs)}}


def _body_records(ctx):
    """正文段落（复用正文范围与前置/后置分区口径）。"""

    from thesis_format_checker.core.scope import front_end_of
    from thesis_format_checker.detectors.headings import _heading_level

    front_end = front_end_of(ctx)
    records = []
    for index, paragraph in enumerate(ctx.paragraphs):
        text = paragraph.text
        if not text.strip():
            continue
        if 0 < front_end and index < front_end:
            continue                     # 封面/声明只按模板同位置比对
        if _heading_level(paragraph):
            continue                     # 标题段单独由标题规则负责
        records.append((index, paragraph, text))
    return records


def _typos(records) -> list:
    findings = []
    for _, paragraph, text in records:
        for wrong, right in _COMMON_TYPOS.items():
            position = text.find(wrong)
            if position < 0:
                continue
            finding = make_finding(
                "144.1", _title("144.1"), f"应写作“{right}”",
                f"“{wrong}”是常见错写（第 {position + 1} 个字符起）", paragraph)
            finding.start, finding.end = position, position + len(wrong)
            finding.excerpt = text[position:position + len(wrong)]
            findings.append(finding)
    return findings


def _doubles(records) -> list:
    findings = []
    for _, paragraph, text in records:
        for token in _BAD_DOUBLES:
            position = text.find(token)
            if position < 0:
                continue
            finding = make_finding(
                "144.2", _title("144.2"), "不应出现重复的功能词",
                f"“{token}”出现重复字（第 {position + 1} 个字符起）", paragraph)
            finding.start, finding.end = position, position + len(token)
            finding.excerpt = token
            findings.append(finding)
    return findings


def _acronyms(records) -> list:
    """首次出现的英文缩略语附近没有全称时提醒（提醒级，不判定对错）。

    精度控制：只提醒**全文出现过两次及以上**的缩略语——只在文中出现一次的英文大写
    词更可能是单位、型号或图表编号，提醒"给全称"的收益低、噪声大；纯罗马数字
    （I/II/III…，常见于图表编号与页码）一律排除。
    """

    seen: dict[str, tuple] = {}
    counts: Counter = Counter()
    for _, _, text in records:
        counts.update(match.group(1) for match in _ACRONYM.finditer(text))
    findings: list = []
    for _, paragraph, text in records:
        for match in _ACRONYM.finditer(text):
            token = match.group(1)
            if token in seen or len(token) < 2 or token.isdigit():
                continue
            if counts[token] < 2 or _ROMAN_NUMERAL.fullmatch(token):
                seen[token] = (paragraph, match.start())
                continue
            # Bind the explanation to THIS token, not another acronym nearby.
            before, after = text[:match.start()], text[match.end():]
            if (re.search(r"[\u4e00-\u9fff]{2,}\s*[（(]\s*$", before)
                    and re.match(r"\s*[）)]", after)) or re.match(
                        r"\s*[（(](?:[\u4e00-\u9fff]{2,}|[A-Za-z][A-Za-z\-\s]{3,})[）)]", after):
                seen[token] = (paragraph, match.start())
                continue
            seen[token] = (paragraph, match.start())
            finding = make_finding(
                "149.1", _title("149.1"), "缩略语首次出现时建议给出中文全称（如“长短期记忆网络（LSTM）”）",
                f"“{token}”在本段首次出现，附近没有全称说明", paragraph,
                status=Status.FAIL)
            finding.start, finding.end = match.start(), match.end()
            finding.excerpt = token
            findings.append(finding)
    return findings


def _number_style(records) -> list:
    """年份/序数写法混用时提醒（同一文档两种写法并存）。"""

    year_cn = year_digit = ordinal_cn = ordinal_digit = None
    for _, paragraph, text in records:
        year_cn = year_cn or _YEAR_CN.search(text)
        year_digit = year_digit or _YEAR_DIGIT.search(text)
        ordinal_cn = ordinal_cn or _ORDINAL_CN.search(text)
        ordinal_digit = ordinal_digit or _ORDINAL_DIGIT.search(text)
    findings = []
    if year_cn and year_digit:
        findings.append(make_finding(
            "151.1", _title("151.1"), "同一文档中年份写法应统一",
            f"同时出现“{year_digit.group()}”与“{year_cn.group()}”两种年份写法", None))
    if ordinal_cn and ordinal_digit:
        findings.append(make_finding(
            "151.1", _title("151.1"), "同一文档中序数词写法应统一",
            f"同时出现“{ordinal_digit.group()}”与“{ordinal_cn.group()}”两种序数写法", None))
    return findings


def _reference_entries(ctx):
    """参考文献区域的条目段落（按标题角色定位，遇到下一个同级/更高级标题停止）。"""

    from thesis_format_checker.detectors.headings import collect_heading_records

    try:
        headings = collect_heading_records(ctx)
    except Exception:
        return []
    start = next((record for record in headings if record["role"] == "references"), None)
    if start is None:
        return []
    level = start["level"]
    begin = ctx.index(start["paragraph"]) + 1
    entries = []
    for index in range(begin, len(ctx.paragraphs)):
        paragraph = ctx.paragraphs[index]
        record = next((item for item in headings if item["paragraph"] is paragraph), None)
        if record is not None and record["level"] <= level:
            break
        if paragraph.text.strip():
            entries.append(paragraph)
    return entries


def _bibliography_case(entries) -> list:
    """英文文献作者姓名写法混用（全大写 vs 首字母大写）。"""

    styles = defaultdict(list)
    for paragraph in entries:
        text = paragraph.text.strip()
        letters = [char for char in text if char.isalpha()]
        if len(letters) < 10 or sum(char.isascii() for char in letters) / len(letters) < 0.8:
            continue                      # 只看英文条目
        head = text[:80]
        if _AUTHOR_CAPS.search(head):
            styles["全大写"].append(paragraph)
        elif _AUTHOR_TITLE.search(head):
            styles["首字母大写"].append(paragraph)
    if len(styles) < 2 or sum(len(items) for items in styles.values()) < 3:
        return []
    detail = "；".join(f"{name} {len(items)} 条" for name, items in sorted(styles.items()))
    anchor = next(iter(styles[sorted(styles)[0]]))
    return [make_finding(
        "119.1", _title("119.1"),
        "英文文献作者姓名的大小写写法应统一（全大写或首字母大写，二者取一）",
        f"参考文献中同时存在不同写法：{detail}", anchor)]


def _title(code: str) -> str:
    item = FINE_RULES.get(code)
    return item.title if item else _TITLES[code.split(".")[0]]


def check_language_errors(baseline, ctx, mapping=None):
    """文字规范检查主入口：一个细粒度规则一条结论，定位到具体文字范围。"""

    del baseline, mapping
    records = _body_records(ctx)
    findings = []
    findings.extend(_typos(records))
    findings.extend(_doubles(records))
    findings.extend(_acronyms(records))
    findings.extend(_number_style(records))
    findings.extend(_bibliography_case(_reference_entries(ctx)))
    return findings
