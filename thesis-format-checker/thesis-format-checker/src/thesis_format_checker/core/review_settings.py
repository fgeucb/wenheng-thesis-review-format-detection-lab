"""只在输出副本里打开“显示批注”开关，保留原论文的修订与显示设置。"""
from __future__ import annotations

from docx.oxml import OxmlElement
from thesis_format_checker.core.xmlutil import qn
from docx.oxml.parser import parse_xml
from lxml import etree

#: CT_Settings 中排在 w:revisionView 之前的设置项（插入位置用）。
_PREDECESSORS = {
    "writeProtection", "view", "zoom", "removePersonalInformation", "removeDateAndTime",
    "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
    "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
    "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins",
    "alignBordersAndEdges", "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter",
    "gutterAtTop", "hideSpellingErrors", "hideGrammaticalErrors", "activeWritingStyle",
    "proofState", "formsDesign", "attachedTemplate", "linkStyles", "stylePaneFormatFilter",
    "stylePaneSortMethod", "documentType", "mailMerge",
}


def visible_comments_settings(blob: bytes) -> bytes:
    """打开标记/批注显示开关，方便打开输出文件就能看到批注。"""

    root = parse_xml(blob)
    revisions = root.findall(qn("w:revisionView"))
    if len(revisions) > 1:
        raise ValueError("文档包含重复的批注显示设置，无法安全输出")
    view = revisions[0] if revisions else OxmlElement("w:revisionView")
    if not revisions:
        successor = next((node for node in root
                          if etree.QName(node).localname not in _PREDECESSORS), None)
        root.insert(root.index(successor) if successor is not None else len(root), view)
    view.set(qn("w:markup"), "1")
    view.set(qn("w:comments"), "1")
    return etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def validate_visible_comments(blob: bytes) -> bool:
    root = etree.fromstring(blob)
    view = root.find(qn("w:revisionView"))
    return view is not None and all(view.get(qn("w:" + name)) in {"1", "true", "on"}
                                    for name in ("markup", "comments"))
