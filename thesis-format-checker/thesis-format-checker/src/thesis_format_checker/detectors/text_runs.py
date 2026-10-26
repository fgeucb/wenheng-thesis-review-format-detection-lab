"""只读的有效文字格式读取（段落、图表、公式等大类共用）。

The original font/body checkers are deliberately untouched. Color themes and
local run overrides are resolved here for new object-specific checks only.
"""
from __future__ import annotations

from copy import deepcopy
import re

from docx.oxml import OxmlElement
from thesis_format_checker.core.xmlutil import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from lxml import etree

from thesis_format_checker.core.docx_context import _styles, _default_prop, effective_runs


def _document(paragraph):
    part = paragraph.part
    if hasattr(part, "document"):
        return part.document
    # Notes and headers have no public python-docx Document property. Callers
    # can provide a detached paragraph with the main document body as parent.
    raise ValueError("此故事区域没有可读取的文档样式上下文")


def _first(props, element, attr="val"):
    for prop in props:
        node = prop.find(qn("w:" + element))
        if node is not None and node.get(qn("w:" + attr)) is not None:
            return node.get(qn("w:" + attr))
    return None


def _theme(document):
    relation = next((r for r in document.part.rels.values() if r.reltype.endswith("/theme")), None)
    return etree.fromstring(relation.target_part.blob) if relation is not None else None


def resolve_color(document, attributes):
    """Resolve w:color/shd-style theme attributes, never compare theme indices.

    Pass local XML attribute names (val/themeColor/themeTint/themeShade). None
    means the effective RGB cannot be confirmed; 'auto' is Word's auto mode,
    not a guessed black color or a school-format constant.
    """
    token = attributes.get("themeColor")
    value = attributes.get("val", "auto")
    if token and token != "none":
        root = _theme(document)
        if root is None:
            return None
        aliases = {"text1": "dk1", "text2": "dk2", "background1": "lt1", "background2": "lt2",
                   "dark1": "dk1", "dark2": "dk2", "light1": "lt1", "light2": "lt2", "hyperlink": "hlink", "followedHyperlink": "folHlink"}
        mapping_names = {"text1": "t1", "text2": "t2", "background1": "bg1", "background2": "bg2",
                         "hyperlink": "hyperlink", "followedHyperlink": "followedHyperlink"}
        mapping = document.settings.element.find(qn("w:clrSchemeMapping"))
        if mapping is not None and token in mapping_names:
            token = mapping.get(qn("w:" + mapping_names[token]), token)
        token = aliases.get(token, token)
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        node = root.find(".//a:clrScheme/a:" + token, ns)
        if node is None or len(node) != 1:
            return None
        color = node[0]
        value = color.get("lastClr") if etree.QName(color).localname == "sysClr" else color.get("val")
    if value == "auto":
        return "auto"
    if not value or not re.fullmatch(r"[0-9a-fA-F]{6}", value):
        return None
    # DrawingML/HLS tint/shade effects require a renderer for exact color.
    # Do not use raw theme indexes or stale fallback RGB as an actual color.
    if attributes.get("themeTint") or attributes.get("themeShade"):
        return None
    return value.upper()


def _font(document, props, slot, theme_slot, language):
    for prop in props:
        node = prop.find(qn("w:rFonts"))
        if node is None:
            continue
        theme = node.get(qn("w:" + theme_slot))
        if theme:
            root = _theme(document)
            if root is None or not theme.startswith(("major", "minor")):
                return None
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            family = "majorFont" if theme.startswith("major") else "minorFont"
            target = "ea" if slot == "eastAsia" else "cs" if slot == "cs" else "latin"
            font = root.find(f".//a:{family}/a:{target}", ns)
            if font is not None and font.get("typeface"):
                return font.get("typeface")
            if slot == "eastAsia":
                lang = (language or "").lower()
                script = "Jpan" if lang.startswith("ja") else "Hang" if lang.startswith("ko") else "Hant" if lang in {"zh-tw", "zh-hk", "zh-mo"} else "Hans" if lang in {"zh-cn", "zh-sg"} else None
                if script:
                    font = root.find(f".//a:{family}/a:font[@script='{script}']", ns)
                    return font.get("typeface") if font is not None else None
            return None
        explicit = node.get(qn("w:" + slot))
        if explicit:
            return explicit
    return None


def _table_unresolved(paragraph, explicit_props):
    table = next((a for a in paragraph._p.iterancestors() if a.tag == qn("w:tbl")), None)
    if table is None:
        return set()
    styles = _document(paragraph).styles
    ref = table.find("./" + qn("w:tblPr") + "/" + qn("w:tblStyle"))
    style = next((s for s in styles if ref is not None and s.style_id == ref.get(qn("w:val"))), None)
    if ref is not None and style is None:
        return {"east_asia_font", "ascii_font", "hansi_font", "size_pt", "bold", "color"}
    if style is None:
        style = styles.default(3)
    table_props = [n for s in _styles(style) for n in s.element.iter(qn("w:rPr"))]
    affected = set()
    for key, element, attr in (("east_asia_font", "rFonts", "eastAsia"), ("ascii_font", "rFonts", "ascii"),
                               ("hansi_font", "rFonts", "hAnsi"), ("size_pt", "sz", "val"),
                               ("color", "color", "val"), ("bold", "b", "val")):
        def affects(prop):
            node = prop.find(qn("w:" + element))
            if node is None:
                return False
            if element == "rFonts":
                theme_attr = "cstheme" if attr == "cs" else attr + "Theme"
                return node.get(qn("w:" + attr)) is not None or node.get(qn("w:" + theme_attr)) is not None
            return True
        # Font slots inherit separately; an ASCII override does not resolve an
        # East Asian table font. Table-style toggles are only settled by an
        # absolute direct-run override, not a paragraph style toggle.
        has_table = any(affects(p) for p in table_props)
        has_override = any(affects(p) for p in explicit_props)
        if has_table and not has_override:
            affected.add(key)
    return affected


def effective_text_runs(paragraph):
    """Return (text, properties) per readable run, honoring direct overrides.

    properties include the old effective_runs keys and hansi_font, run,
    character_style, language, vertical_align, field, revision, limitations.
    Unsupported effective properties are None, never a guessed default value.
    """
    document = _document(paragraph)
    default = _default_prop(paragraph, "rPr")
    result = []
    # 复杂域的 begin/separate/end 与普通文字可以共处一段。旧实现只要段内
    # 出现任意域就把所有 run 标成 field，导致含引用的整段正文无法检测字体。
    # 这里按 OOXML 文档顺序只标记实际位于域范围内的 run。
    field_runs = set()
    field_depth = 0
    for node in paragraph._p.iter():
        if node.tag == qn("w:r") and field_depth:
            field_runs.add(node)
        if node.tag == qn("w:fldChar"):
            kind = node.get(qn("w:fldCharType"))
            run_ancestor = next((ancestor for ancestor in node.iterancestors()
                                 if ancestor.tag == qn("w:r")), None)
            if kind == "begin":
                field_depth += 1
                if run_ancestor is not None:
                    field_runs.add(run_ancestor)
            elif kind == "end":
                if run_ancestor is not None:
                    field_runs.add(run_ancestor)
                field_depth = max(0, field_depth - 1)
    for element in paragraph._p.iter(qn("w:r")):
        ancestors = list(element.iterancestors())
        if any(a.tag in {qn("w:txbxContent"), qn("w:del"), qn("w:moveFrom")} for a in ancestors):
            continue
        run = Run(element, paragraph)
        text = run.text
        if not text:
            continue
        props = ([element.rPr] if element.rPr is not None else [])
        props += [s.element.rPr for s in [*_styles(run.style), *_styles(paragraph.style)] if s.element.rPr is not None]
        # Normal/table-style precedence and conditional table formatting are
        # not fully exposed by python-docx. Only direct run values are certain
        # to override them; leave ambiguous inherited values unknown.
        explicit_props = [element.rPr] if element.rPr is not None else []
        if default is not None:
            props.append(default)
        # Reuse prior toggle/style resolution on a detached shell. Nested
        # hyperlink/simple-field runs become readable without rewriting them.
        shell = OxmlElement("w:p")
        if paragraph._p.pPr is not None:
            shell.append(deepcopy(paragraph._p.pPr))
        shell.append(deepcopy(element))
        values = effective_runs(Paragraph(shell, document._body))
        if not values:
            continue
        fmt = dict(values[0][1])
        lang = _first(props, "lang", "eastAsia")
        if not lang:
            theme_lang = document.settings.element.find(qn("w:themeFontLang"))
            lang = theme_lang.get(qn("w:eastAsia")) if theme_lang is not None else None
        fmt["east_asia_font"] = _font(document, props, "eastAsia", "eastAsiaTheme", lang)
        fmt["ascii_font"] = _font(document, props, "ascii", "asciiTheme", lang)
        fmt["hansi_font"] = _font(document, props, "hAnsi", "hAnsiTheme", lang)
        color_node = next((p.find(qn("w:color")) for p in props if p.find(qn("w:color")) is not None), None)
        fmt["color"] = resolve_color(document, {etree.QName(k).localname: v for k, v in color_node.attrib.items()}) if color_node is not None else "auto"
        limits = []
        if fmt["color"] is None:
            limits.append("主题颜色或色调效果无法确定有效值")
        affected = _table_unresolved(paragraph, explicit_props)
        if affected:
            limits.append("表格样式的文字格式叠加无法完整解析")
            for key in affected:
                fmt[key] = None
        if any(p.find(qn("w:rtl")) is not None or p.find(qn("w:cs")) is not None for p in props):
            limits.append("复杂文字方向/字形规则需要Word解析")
            for key in ("east_asia_font", "ascii_font", "hansi_font", "size_pt", "bold"):
                fmt[key] = None
        in_simple_field = any(ancestor.tag == qn("w:fldSimple") for ancestor in ancestors)
        fmt.update(run=run, character_style=run.style.name if run.style is not None else "",
                   language=lang, vertical_align=_first(props, "vertAlign") or "baseline",
                   field=element in field_runs or in_simple_field,
                   revision=bool(paragraph._p.xpath(".//w:ins | .//w:del | .//w:moveFrom | .//w:moveTo | .//w:rPrChange | .//w:pPrChange")) or
                            any(a.tag in {qn("w:ins"), qn("w:moveTo")} for a in ancestors),
                   limitations=limits)
        result.append((text, fmt))
    return result
