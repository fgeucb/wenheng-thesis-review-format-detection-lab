"""封面图片与图形检测（规则 017：校徽、学校名称图片/艺术字等）。

为什么需要单独一套逻辑
======================

python-docx 的 ``paragraph.text`` 只包含普通文字：**图片、艺术字、文本框里的
文字都读不到**，所以校徽、图片版校名这类内容以前完全没被检测。本模块直接读
OOXML：

* 图片：``a:blip/@r:embed`` → 关系 → 图片部件字节 → **sha256 指纹**；
  尺寸取 ``wp:extent/@cx,@cy``（EMU）；浮动图还读 ``wp:anchor`` 的绝对偏移；
* 艺术字/文本框：``w:txbxContent`` / ``wps:txbx`` / ``v:textbox`` 里的 ``w:t``；
* 位置：用“段落先后顺序”作为封面内的相对位置（存储顺序 = 视觉顺序）。

判定规则（按用户要求）
======================

1. **图片数据一致**（sha256 相同）+ **位置接近**（段落序号相差不超过容差，
   浮动图还比较绝对偏移）→ 判定无误；
2. 模板有、论文没有（或内容不同）→ 判定**缺失/不一致**；
3. **论文多出来的图片不报错**（例如论文里贴了作者签名图片，模板里是空行）；
4. 艺术字/文本框里的文字：模板封面里有，就要求论文也有同样的文字。
"""

from __future__ import annotations

import hashlib
import re
from typing import Any

from thesis_format_checker.core.xmlutil import qn

from thesis_format_checker.checks.catalog import FINE_RULES
from thesis_format_checker.core.docx_context import make_finding

#: 本模块负责的规则族（供规则目录自检识别）
RULE_TITLES = {"017": "封面校徽、学校名称图片或艺术字不符合模板"}

#: 组成部分 → 细粒度规则编号
CODE_MISSING_LOGO = "017.1"      # 缺少校徽
CODE_DIFF_LOGO = "017.2"         # 校徽图片内容与模板不一致
CODE_POS_LOGO = "017.3"          # 校徽位置或尺寸与模板不一致
CODE_MISSING_NAME = "017.4"      # 缺少学校名称（图片/艺术字）
CODE_DIFF_NAME = "017.5"         # 学校名称图片/艺术字与模板不一致
CODE_POS_OTHER = "017.6"         # 其他封面图片位置或尺寸与模板不一致

EMU_PER_MM = 36000.0
#: 位置容差：段落序号相差不超过 3 段、浮动偏移相差不超过 10 毫米
INDEX_TOLERANCE = 3
OFFSET_TOLERANCE_MM = 10.0
#: 尺寸容差：相差不超过 2 毫米或 15%
SIZE_TOLERANCE_MM = 2.0
SIZE_TOLERANCE_RATE = 0.15
#: 封面结束标志：出现这些标题即认为封面区结束
_COVER_END = ("摘要", "abstract", "目录", "关键词", "第一章", "第1章", "绪论", "引言", "第1章")
#: 文本框里的模板提示语（“宋体小五号，垂直居中”之类）不参与比对
_HINT = re.compile(r"(宋体|黑体|楷体|仿宋|字号|磅|居中|小五|三号|四号|五号)")


def _title(code: str) -> str:
    item = FINE_RULES.get(code)
    return item.title if item else code


def _normalize(text: str) -> str:
    return re.sub(r"[\s\u3000]+", "", text or "").casefold()


def _extent_mm(paragraph) -> tuple[float | None, float | None]:
    """图片显示尺寸（毫米），来自 ``wp:extent``（EMU）。"""

    nodes = paragraph._p.xpath(".//wp:extent")
    if not nodes:
        return None, None
    try:
        return (round(int(nodes[0].get("cx")) / EMU_PER_MM, 2),
                round(int(nodes[0].get("cy")) / EMU_PER_MM, 2))
    except (TypeError, ValueError):
        return None, None


def _anchor_offset(paragraph) -> tuple[float | None, float | None, bool]:
    """浮动图片的绝对偏移（毫米）与是否衬于文字下方。"""

    nodes = paragraph._p.xpath(".//wp:anchor")
    if not nodes:
        return None, None, False
    node = nodes[0]
    def offset(axis: str) -> float | None:
        text = node.xpath(f"string(.//wp:position{axis}/wp:posOffset)")
        return round(int(text) / EMU_PER_MM, 2) if text and text.lstrip("-").isdigit() else None
    return offset("H"), offset("V"), node.get("behindDoc") not in (None, "0", "false")


def _role(width: float | None, height: float | None) -> str:
    """按宽高比判断用途：校徽（近正方形小图）/ 学校名称（横幅图）/ 其他。"""

    if not width or not height:
        return "other"
    ratio = width / height
    if 0.7 <= ratio <= 1.4 and max(width, height) <= 45:
        return "logo"
    if ratio >= 2.5 and width >= 35:
        return "wordmark"
    return "other"


def _cover_paragraphs(ctx) -> list[tuple[int, Any]]:
    """封面区域：文档开头到第一个“摘要/目录/第一章”等标志之前。"""

    result: list[tuple[int, Any]] = []
    for index, paragraph in enumerate(ctx.paragraphs, 1):
        text = _normalize(paragraph.text)
        if index > 2 and text and len(text) <= 20 and any(token in text for token in _COVER_END):
            break
        result.append((index, paragraph))
        if index >= 80:      # 兜底：封面不会超过前 80 段
            break
    return result


def _read_media(ctx) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """读取封面区域的图片与艺术字/文本框内容。"""

    images: list[dict[str, Any]] = []
    texts: list[dict[str, Any]] = []
    for index, paragraph in _cover_paragraphs(ctx):
        for blip in paragraph._p.xpath(".//a:blip"):
            embed = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
            part = paragraph.part.related_parts.get(embed) if embed else None
            if part is None:
                continue
            width, height = _extent_mm(paragraph)
            pos_h, pos_v, behind = _anchor_offset(paragraph)
            images.append({
                "sha": hashlib.sha256(part.blob).hexdigest(),
                "partname": str(part.partname),
                "width_mm": width,
                "height_mm": height,
                "index": index,
                "pos_h": pos_h,
                "pos_v": pos_v,
                "behind": behind,
                "role": _role(width, height),
            })
        # 艺术字/文本框：读 w:txbxContent 里的文字
        for box in paragraph._p.iter():
            if not isinstance(box.tag, str) or not box.tag.endswith("}txbxContent"):
                continue
            text = "".join(node.text or "" for node in box.iter(qn("w:t"))).strip()
            if text:
                texts.append({"text": text, "index": index})
    return images, texts


def extract_cover_media_baseline(ctx) -> dict[str, Any]:
    """提取模板封面的图片与艺术字清单（只读，不修改文档）。"""

    images, texts = _read_media(ctx)
    return {"cover_media": {"images": images,
                            "texts": [t for t in texts if not _HINT.search(t["text"])]}}


def _size_close(left: dict[str, Any], right: dict[str, Any]) -> bool:
    """尺寸是否接近：绝对差 ≤ 2 毫米，或相对差 ≤ 15%。"""

    for key in ("width_mm", "height_mm"):
        a, b = left.get(key), right.get(key)
        if a is None or b is None:
            continue
        if abs(a - b) <= SIZE_TOLERANCE_MM:
            continue
        if max(a, b) and abs(a - b) / max(a, b) <= SIZE_TOLERANCE_RATE:
            continue
        return False
    return True


def _position_close(template: dict[str, Any], paper: dict[str, Any]) -> bool:
    """位置是否接近：段落序号相差 ≤ 3 段；浮动图再看绝对偏移。"""

    if abs(template["index"] - paper["index"]) > INDEX_TOLERANCE:
        return False
    for key in ("pos_h", "pos_v"):
        a, b = template.get(key), paper.get(key)
        if a is None or b is None:
            continue
        if abs(a - b) > OFFSET_TOLERANCE_MM:
            return False
    return True


def check_cover_media_errors(baseline: dict, ctx) -> list:
    """逐项比对模板封面的图片/艺术字与论文，缺失或内容不符各出一条结论。"""

    template_items = (baseline.get("cover_media") or {}).get("images") or []
    if not template_items:
        return []          # 模板封面没有图片 → 无所要求
    paper_items, paper_texts = _read_media(ctx)
    anchor = ctx.paragraphs[0] if ctx.paragraphs else None
    findings = []

    for template in template_items:
        role = template["role"]
        same_hash = next((item for item in paper_items if item["sha"] == template["sha"]), None)
        if same_hash is None:
            same_role = [item for item in paper_items if item["role"] == role]
            if role == "logo":
                if same_role:
                    findings.append(make_finding(CODE_DIFF_LOGO, _title(CODE_DIFF_LOGO),
                        "校徽与模板为同一图片（数据一致）",
                        f"封面第 {same_role[0]['index']} 段的校徽图片与模板不同",
                        anchor))
                else:
                    findings.append(make_finding(CODE_MISSING_LOGO, _title(CODE_MISSING_LOGO),
                        f"封面应含模板中的校徽图片（{template['partname']}，"
                        f"{template['width_mm']}×{template['height_mm']} mm）",
                        "封面未识别到校徽图片", anchor))
            elif role == "wordmark":
                if same_role:
                    findings.append(make_finding(CODE_DIFF_NAME, _title(CODE_DIFF_NAME),
                        "学校名称图片/艺术字与模板一致",
                        f"封面第 {same_role[0]['index']} 段的学校名称图与模板不同",
                        anchor))
                else:
                    findings.append(make_finding(CODE_MISSING_NAME, _title(CODE_MISSING_NAME),
                        f"封面应含模板中的学校名称图片/艺术字（{template['partname']}）",
                        "封面未识别到学校名称图片或艺术字", anchor))
            continue
        # 内容一致 → 再看位置与尺寸
        if not _position_close(template, same_hash):
            code = CODE_POS_LOGO if role == "logo" else CODE_POS_OTHER
            findings.append(make_finding(code, _title(code),
                f"应与模板同位置（模板在第 {template['index']} 段"
                + (f"，浮动偏移 {template['pos_h']},{template['pos_v']} mm" if template["pos_h"] is not None else "")
                + "）",
                f"论文中位于第 {same_hash['index']} 段"
                + (f"，浮动偏移 {same_hash['pos_h']},{same_hash['pos_v']} mm" if same_hash["pos_h"] is not None else ""),
                anchor))
        elif not _size_close(template, same_hash):
            code = CODE_POS_LOGO if role == "logo" else CODE_POS_OTHER
            findings.append(make_finding(code, _title(code),
                f"显示尺寸应为 {template['width_mm']}×{template['height_mm']} mm",
                f"实际为 {same_hash['width_mm']}×{same_hash['height_mm']} mm", anchor))

    # 艺术字/文本框：模板封面里有的文字，论文封面也应出现
    for item in (baseline.get("cover_media") or {}).get("texts") or []:
        wanted = _normalize(item["text"])
        if len(wanted) < 2:
            continue
        if any(wanted in _normalize(text["text"]) or _normalize(text["text"]) in wanted
               for text in paper_texts):
            continue
        findings.append(make_finding(CODE_MISSING_NAME, _title(CODE_MISSING_NAME),
            f"封面应含模板中的文字/艺术字「{item['text'][:30]}」",
            "论文封面未识别到对应文字或艺术字", anchor))
    return findings
