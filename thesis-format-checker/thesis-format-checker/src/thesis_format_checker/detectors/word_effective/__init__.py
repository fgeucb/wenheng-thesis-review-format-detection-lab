"""旧版 Word 有效格式检测（保持原有算法，按用途拆分的独立子包）。

这一组模块由 Microsoft Word 通过 COM 解析“有效格式”（直接格式 + 样式继承 +
主题字体），用于检查模板与文档之间的页边距、字体、字号、缩进、行距与标点。

* :mod:`runner` —— 对照流程（打开模板与文档、逐条比较、写批注）；
* :mod:`profile` —— 从模板提取标准，并对冲突/样本不足的情况拒绝下结论；
* 其余模块 —— 每条规则一组的比较函数（margins、fonts、font_size、
  first_line_indent、paragraph_spacing、punctuation、headers、headings、
  paper_size）。

调用入口是 :func:`~thesis_format_checker.engine.pipeline.compare_documents`；
本子包只负责“怎么读、怎么比”，不参与调度。
"""
