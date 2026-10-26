"""取证实现层：按大类读取模板基准并给出结论（不含调度）。

每个模块对应一个检测大类，与 ``checks/cXX_*.py`` 的类一一对应：

======================================  ================================
大类检查器（规则入口）                    取证实现（模板/对象读取）
======================================  ================================
``checks/c01_structure.py``              ``detectors/structure.py``
``checks/c02_front.py``                  ``detectors/front.py``
``checks/c03_pages.py``                  ``detectors/pages.py``
``checks/c04_headers.py``                ``detectors/headers.py``
``checks/c05_headings.py``               ``detectors/headings.py``
``checks/c06_toc.py``                    ``detectors/toc.py``
``checks/c07_paragraphs.py``             ``detectors/paragraphs.py``
``checks/c08_figures.py``                ``detectors/figures.py``
``checks/c09_equations.py``              ``detectors/equations.py``
``checks/c10_citations.py``              ``detectors/notes.py``
``checks/c11_bibliography.py``           ``detectors/bibliography.py``
``checks/c12_characters.py``             ``detectors/characters.py``
``checks/c13_consistency.py``            ``detectors/consistency.py``
``checks/c15_appendix.py``               ``detectors/appendices.py`` + ``formats.py``
======================================  ================================

公共工具：``fields.py``（域结构扫描）、``text_runs.py``（有效文字格式）、
``word_effective/``（旧版 Word 有效格式检测，含页边距、字体、行距、标点等）。
"""
