"""XML 小工具：带缓存的 ``qn()``。

为什么需要
==========

``qn("w:tbl")`` 每次都要做前缀解析 + 字符串拼接。检测器里有大量"逐节点比较标签"
的循环（例如判断某个节点是否在删除内容/文本框里），一次运行能调用到**数千万次**：
实测 67 页论文一次检测里 ``qn`` 被调用约 3300 万次，光是它自己就吃掉二十多秒。

``qn`` 是纯函数（同样的字符串必然得到同样的 Clark 名），因此这里用 ``lru_cache``
缓存一层：调用点写法不变、返回值完全一致，只是第二次之后直接命中缓存。

用法：把热路径模块里的 ``from docx.oxml.ns import qn``
换成 ``from thesis_format_checker.core.xmlutil import qn`` 即可。
"""

from __future__ import annotations

from functools import lru_cache

from docx.oxml.ns import qn as _qn
from lxml import etree  # noqa: F401  （导出给调用方做类型标注时使用）

#: 带缓存的 ``qn``：语义与 ``docx.oxml.ns.qn`` 完全一致（纯函数）。
qn = lru_cache(maxsize=None)(_qn)

#: "不算可见文字"的子树：删除内容、移动内容、文本框内容。
SKIP_SUBTAGS = frozenset(qn(tag) for tag in ("w:del", "w:moveFrom", "w:txbxContent"))

_TEXT_TAG = qn("w:t")
_TAB_TAG = qn("w:tab")
_BREAK_TAGS = frozenset((qn("w:br"), qn("w:cr")))
_TYPE_ATTR = qn("w:type")


def visible_text(element, *, page_breaks_as_formfeed: bool = False) -> str:
    """元素内的可见文字：跳过"删除内容/文本框"子树。

    旧写法是对每个节点回溯一遍祖先链（``any(a.tag in ... for a in node.iterancestors())``），
    一段十几行的文字要跑几百次 Python 迭代；这里改成**一次深度优先遍历**，
    边下探边携带"当前是否在不该看的子树里"，结果完全一致。

    ``page_breaks_as_formfeed``：分页符/分栏符输出 ``\\f``（字符规则需要），
    否则统一输出 ``\\n``。
    """

    chunks: list[str] = []
    # 段落本身可能就在文本框/删除区里：祖先链只查一次（等价于旧逐节点回溯）。
    root_skipped = any(ancestor.tag in SKIP_SUBTAGS for ancestor in element.iterancestors())

    def walk(node, skipped: bool) -> None:
        if not skipped and node.tag in SKIP_SUBTAGS:
            skipped = True
        if not skipped:
            if node.tag == _TEXT_TAG:
                chunks.append(node.text or "")
            elif node.tag == _TAB_TAG:
                chunks.append("\t")
            elif node.tag in _BREAK_TAGS:
                if page_breaks_as_formfeed and node.get(_TYPE_ATTR) in {"page", "column"}:
                    chunks.append("\f")
                else:
                    chunks.append("\n")
        for child in node:
            walk(child, skipped)

    walk(element, root_skipped)
    return "".join(chunks)


__all__ = ["SKIP_SUBTAGS", "qn", "visible_text"]
