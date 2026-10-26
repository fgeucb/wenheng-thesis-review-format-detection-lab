"""基础设施层：与 Word/python-docx 交互、文档模型与批注写入。

* :mod:`~thesis_format_checker.core.models` —— ``Finding`` / ``RunSummary`` 等数据模型；
* :mod:`~thesis_format_checker.core.docx_context` —— 只读的 docx 段落/格式上下文；
* :mod:`~thesis_format_checker.core.effective_format` —— 有效文字格式与脚本切分；
* :mod:`~thesis_format_checker.core.word_session` —— 独立 Word 进程与只读打开；
* :mod:`~thesis_format_checker.core.scope` —— 正文范围识别；
* :mod:`~thesis_format_checker.core.comments` —— 原生批注锚定与写入。

本层不包含任何“某条规则”的判断，只提供读写文档的能力。
"""

from thesis_format_checker.core.models import Finding, RunSummary, Status

__all__ = ["Finding", "RunSummary", "Status"]
