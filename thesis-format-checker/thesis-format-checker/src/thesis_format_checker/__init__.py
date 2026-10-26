"""论文格式对照检测系统。

以模板 A 为唯一标准，对照检查待检测文档 B，并把问题写成 Word 原生批注。

代码分层
========

``core/``        读写 Word/docx 的基础设施（Word 会话、段落上下文、批注写入）
``checks/``      15 个大类检测器（每类一个类、每条规则一个 ``check_<编号>`` 方法）
``detectors/``   按大类拆分的取证实现（模板基准读取 + 结论产出）
``engine/``      调度（pipeline）与报告输出（reporting）

命令入口见 :mod:`thesis_format_checker.cli`。
"""

__version__ = "1.0.0"

__all__ = ["__version__"]
