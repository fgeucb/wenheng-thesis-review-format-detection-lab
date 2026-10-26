"""旧版 Word“有效格式”对照流程：打开模板与文档、逐条比较、写入批注。

与 docx 结构检测的区别：这里由 Microsoft Word 通过 COM 解析**有效格式**
（直接格式 + 样式继承 + 主题字体），因此能覆盖 023/025/035/045–047/067–069/
072/074/075/122/123 这些必须依赖 Word 解析的规则。

对外统一入口是 :func:`thesis_format_checker.engine.pipeline.compare_documents`；
本模块只实现“怎么读、怎么比”。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from thesis_format_checker.core.comments import add_comment
from thesis_format_checker.core.effective_format import clean_word_text, read_span_formats
from thesis_format_checker.core.models import RunSummary
from thesis_format_checker.detectors.word_effective.first_line_indent import compare_first_line_indent, read_indent
from thesis_format_checker.detectors.word_effective.font_size import compare_font_sizes
from thesis_format_checker.detectors.word_effective.fonts import compare_fonts
from thesis_format_checker.detectors.word_effective.correspondence import build_body_basis, build_heading_basis
from thesis_format_checker.detectors.word_effective.margins import compare_margin
from thesis_format_checker.detectors.word_effective.paper_size import compare_paper_size
from thesis_format_checker.detectors.word_effective.headers import compare_headers, read_headers
from thesis_format_checker.detectors.word_effective.headings import collect_headings, read_heading, compare_heading
from thesis_format_checker.detectors.word_effective.paragraph_spacing import read_spacing, compare_spacing, blank_paragraph_finding, describe_line_spacing
from thesis_format_checker.detectors.word_effective.punctuation import inspect_punctuation
from thesis_format_checker.detectors.word_effective.pagination import (
    check_chapter_page_breaks, check_heading_pagination, check_toc_layout, verify_page_numbering,
)
from thesis_format_checker.core.word_facts import WordFacts
from thesis_format_checker.core.word_pages import PageFacts
from thesis_format_checker.core.scope import collect_body_paragraphs
from thesis_format_checker.detectors.word_effective.profile import build_template_profile, read_margin
from thesis_format_checker.core.word_session import WordAutomationError, WordSession


def _section_map(mapping: dict[str, Any] | None) -> dict[int, int]:
    raw = (mapping or {}).get("section_map", {})
    if not isinstance(raw, dict):
        raise WordAutomationError("映射配置中的 section_map 必须是对象，例如 {'1': 2}。")
    try:
        return {int(document_index): int(template_index) for document_index, template_index in raw.items()}
    except (TypeError, ValueError) as exc:
        raise WordAutomationError("section_map 的键和值都必须是从 1 开始的节序号。") from exc


def _header_section_map(mapping: dict[str, Any] | None) -> dict[int, int]:
    raw = (mapping or {}).get("header_section_map", {})
    if not isinstance(raw, dict):
        raise WordAutomationError("映射配置中的 header_section_map 必须是对象。")
    try:
        return {int(document_index): int(template_index)
                for document_index, template_index in raw.items()}
    except (TypeError, ValueError) as exc:
        raise WordAutomationError(
            "header_section_map 的键和值都必须是从 1 开始的节序号。") from exc


def _backfill_excerpts(findings, document, facts: WordFacts | None = None) -> None:
    """给只带段号的结论补上段落文字摘录。

    旧版流程记录的是 Word 段落号（含文本框等段落），与 docx 段落号存在偏移；补上
    摘录后调度层可以按文字核对位置，把批注锚定在同一段文字上，而不是换一个位置。

    ``facts``：段落事实表（文本已在正文遍历时读过，直接命中缓存）。
    """

    for finding in findings:
        if getattr(finding, "excerpt", "") or finding.paragraph_index is None:
            continue
        try:
            if facts is not None:
                text = facts.clean_text(finding.paragraph_index).strip()
            else:
                text = clean_word_text(document.Paragraphs(finding.paragraph_index).Range.Text).strip()
        except Exception:
            continue
        if text:
            finding.excerpt = text[:30]


def _load_template_profile(cache, template_path, mapping, heading_styles, document, body):
    """模板剖面（Word 有效格式基准）：同一份模板跨论文复用。

    模板剖面的内容只取决于：模板文件内容、正文范围映射、标题样式映射与代码版本；
    这些都已经在缓存键里。因此"同一所学校模板 + 换一篇论文"时，模板侧的
    有效字体/字号/缩进/行距不再重读 Word（省下每篇论文一次模板格式遍历）。
    """

    if cache is None or not getattr(cache, "enabled", False):
        return build_template_profile(document, body, heading_styles)
    from thesis_format_checker.core.cache import clean_params

    key = cache.key("template-profile", template=template_path,
                    params={"heading_styles": heading_styles or {},
                            "mapping": clean_params(mapping or {})})
    cached = cache.load("template-profile", key, "模板剖面（Word 有效格式）")
    if cached is not None:
        return cached
    profile = build_template_profile(document, body, heading_styles)
    cache.store("template-profile", key, profile, "模板剖面（Word 有效格式）")
    return profile


def compare_word_documents(
    template_path: Path,
    document_path: Path,
    output_path: Path,
    mapping: dict[str, Any] | None = None,
    overwrite: bool = False,
    margin_tolerance_mm: float = 0.2,
    font_size_tolerance_pt: float = 0.1,
    indent_tolerance_pt: float = 0.5,
    write_comments: bool = True,
    document_context=None,
    layout_request: dict | None = None,
    word_cache=None,
) -> RunSummary:
    """旧版 Word“有效格式”对照流程。

    Args:
        write_comments: 是否把结论写进 ``output_path`` 副本。调度层会自己把批注
            写进最终输出，因此正常流程传 ``False``，省掉一批会被丢弃的
            ``Comments.Add`` 与整篇保存。集成测试与命令行单跑仍可用默认 ``True``。
        document_context: 可选的 docx 上下文（调度层传入）。给出时，本流程会在
            **同一次 Word 会话**里顺手读取排版几何，避免为几何再开一次 Word。
        layout_request: 几何读取请求（``{"indices", "want_figures", "want_tables"}``），
            由 :func:`detectors.rendered.content_layout_request` 按文档内容先筛出来。
        word_cache: 可选的结果缓存（调度层传入）。给出时模板剖面按模板内容哈希复用。
    """

    summary = RunSummary(output_path=str(output_path.resolve()))
    output_created = False
    try:
        with WordSession() as session:
            template = session.open_readonly(template_path.resolve())
            template_facts = WordFacts(template, label="template")
            template_body = collect_body_paragraphs(template, mapping, "template",
                                                    facts=template_facts)
            if not template_body:
                raise WordAutomationError(
                    "未能在模板 A 中识别出正文。请检查正文样式，或使用 --mapping 指定 template.body_styles/正文范围。"
                )
            heading_styles = (mapping or {}).get("heading_styles") or {}
            profile = _load_template_profile(word_cache, template_path, mapping, heading_styles,
                                             template, template_body)
            summary.unknowns.extend(profile.unknown_reasons)
            if profile.first_line_indent and profile.first_line_indent.characters is not None and abs(profile.first_line_indent.characters - 2) > 0.05:
                summary.unknowns.append("【072】模板 A 的首行缩进不是用户要求的 2 字符；仍按 2 字符检测，建议核对模板")

            output = session.create_output_copy(document_path.resolve(), output_path.resolve(), overwrite)
            output_created = True
            # 事实表：文本/样式/大纲级别/是否在表格只读一遍，之后四处消费全部命中缓存。
            facts = WordFacts(output, label="document")
            body = collect_body_paragraphs(output, mapping, "document", facts=facts)
            if not body:
                raise WordAutomationError(
                    "未能在文档 B 中识别出正文。请检查正文样式，或使用 --mapping 指定 document.body_styles/正文范围。"
                )
            summary.body_paragraphs_checked = len(body)

            explicit_sections = _section_map(mapping)
            explicit_header_sections = _header_section_map(mapping)
            for section_index in range(1, output.Sections.Count + 1):
                paper_finding = compare_paper_size(output.Sections(section_index), section_index)
                if paper_finding:
                    summary.findings.append(paper_finding)
                expected = profile.uniform_margin
                template_index = explicit_sections.get(section_index)
                header_template_index = explicit_header_sections.get(
                    section_index, template_index)
                if len(set(tuple(h.items()) for h in profile.headers)) == 1:
                    expected_header = profile.headers[0]
                elif (header_template_index is not None
                      and 1 <= header_template_index <= len(profile.headers)):
                    expected_header = profile.headers[header_template_index - 1]
                else:
                    expected_header = None
                    summary.unknowns.append(f"【035】第 {section_index} 节页眉无法对应模板节；请配置 section_map")
                if expected_header is not None:
                    summary.findings.extend(compare_headers(expected_header, read_headers(output.Sections(section_index)), section_index))
                if expected is None:
                    if template_index is None:
                        summary.unknowns.append(
                            f"【025】文档 B 第 {section_index} 节未配置与模板特殊节的对应关系，未检测边距"
                        )
                        continue
                    if not 1 <= template_index <= len(profile.margins):
                        summary.unknowns.append(
                            f"【025】section_map 指向不存在的模板第 {template_index} 节"
                        )
                        continue
                    expected = profile.margins[template_index - 1]
                actual = read_margin(output.Sections(section_index))
                summary.sections_checked += 1
                # 逐方向比较：一个方向一条结论（025.1–025.4）
                summary.findings.extend(
                    compare_margin(expected, actual, section_index, margin_tolerance_mm)
                )

            actual_unknown_fonts = 0
            actual_unknown_sizes = 0
            headings = collect_headings(output, heading_styles, facts=facts)
            # 同一层级 + 同一位置：数量相同时逐个对应，否则取同级最常见样例。
            body_basis = build_body_basis(profile.body_samples, body)
            heading_basis = build_heading_basis(profile.heading_samples, headings)
            for index, level, paragraph in headings:
                sample = heading_basis.get(index)
                if sample is None:
                    level_name = "一二三四五六七八九"[level - 1] if 1 <= level <= 9 else str(level)
                    summary.unknowns.append(
                        f"【{('045' if level == 1 else '046' if level == 2 else '047')}】"
                        f"第 {index} 段{level_name}级标题在模板 A 中没有同级别样例，无法比对"
                    )
                    continue
                heading_text = facts.clean_text(index).strip()
                heading_findings, heading_unknowns = compare_heading(
                    sample.spec, read_heading(output, paragraph), level, index,
                    text=heading_text)
                summary.findings.extend(heading_findings)
                summary.unknowns.extend(heading_unknowns)
            body_indexes = {p.index for p in body}
            template_indexes = {p.index for p in template_body}
            template_blank = any(
                i not in template_indexes and i - 1 in template_indexes and i + 1 in template_indexes
                for i in range(2, template_facts.count)
                if not template_facts.text(i).strip("\r\x07 \t\u3000")
            )
            if template_blank:
                summary.unknowns.append("【075】模板 A 的正文段落间存在空段落，无法将相同空行判为错误")
            else:
                for i in range(2, facts.count):
                    if i - 1 in body_indexes and i + 1 in body_indexes \
                            and not facts.text(i).strip("\r\x07 \t\u3000"):
                        summary.findings.append(blank_paragraph_finding(i))
            no_basis = 0
            for info in body:
                spans = read_span_formats(output, info.paragraph)
                actual_unknown_fonts += sum(1 for span in spans if span.font_name is None)
                actual_unknown_sizes += sum(1 for span in spans if span.size_pt is None)
                sample = body_basis.get(info.index)
                expected_chinese = (sample.chinese_font if sample else None) or profile.chinese_font
                expected_latin = (sample.latin_font if sample else None) or profile.latin_font
                expected_size = (sample.size_pt if sample else None)
                if expected_size is None:
                    expected_size = profile.font_size_pt
                expected_indent = (sample.first_line_indent if sample else None) or profile.first_line_indent
                expected_spacing = (sample.spacing if sample else None) or profile.body_spacing
                if sample is None or not any((expected_chinese, expected_latin, expected_size)):
                    no_basis += 1
                summary.findings.extend(
                    compare_fonts(
                        spans, expected_chinese, expected_latin, info.index, info.text
                    )
                )
                summary.findings.extend(
                    compare_font_sizes(
                        spans, expected_size, info.index, info.text, font_size_tolerance_pt
                    )
                )
                # 缩进只读一次：比较与"无法判定"判断共用同一个实际值。
                actual_indent = read_indent(info.paragraph)
                summary.findings.extend(
                    compare_first_line_indent(
                        info.paragraph,
                        info.text,
                        expected_indent,
                        info.index,
                        indent_tolerance_pt,
                        actual_indent=actual_indent,
                    )
                )
                if expected_spacing:
                    spacing = read_spacing(info.paragraph)
                    if spacing:
                        summary.findings.extend(compare_spacing(expected_spacing, spacing, info.index))
                        if describe_line_spacing(expected_spacing) is None or describe_line_spacing(spacing) is None:
                            summary.unknowns.append(f"【074】第 {info.index} 段行距类型无法识别，请在 Word 中核对")
                    else:
                        summary.unknowns.append(f"【074/075】第 {info.index} 段行距或段间距无法判定")
                summary.findings.extend(inspect_punctuation(info.text, info.index))
                if actual_indent is None or (
                    actual_indent.characters is None and (
                        expected_indent is None
                        or expected_indent.characters is None
                        or abs(expected_indent.characters - 2) > 0.05
                    )
                ):
                    summary.unknowns.append(
                        f"【072】文档 B 第 {info.index} 段的有效首行缩进无法由 Word 确定"
                    )
            if no_basis:
                summary.unknowns.append(
                    f"模板 A 有 {no_basis} 个正文段落无法提供字体或字号基准，这些段落未做字体字号比对"
                )
            if actual_unknown_fonts:
                summary.unknowns.append(
                    f"文档 B 有 {actual_unknown_fonts} 个文字片段的有效字体无法由 Word 确定"
                )
            if actual_unknown_sizes:
                summary.unknowns.append(
                    f"文档 B 有 {actual_unknown_sizes} 个文字片段的有效字号无法由 Word 确定"
                )

            # 需要 Word 排版结果的规则：标题同页（054）、章节另起页（030）、页码（040）。
            # 分页事实只用于判定，不写回文档；读不到时保持原有“无法判定”行为。
            try:
                template_breaks = False
                try:
                    for _, level, paragraph in collect_headings(template, heading_styles,
                                                                facts=template_facts):
                        if level == 1 and bool(paragraph.Format.PageBreakBefore):
                            template_breaks = True
                            break
                except Exception:
                    template_breaks = False
                # 读页码前必须恢复分页（读格式阶段关闭了后台重排，换取大幅提速）。
                session.resume_pagination()
                pagination = PageFacts(output, facts)
                facts.pagination = pagination
                summary.render_facts = pagination.as_dict()
                summary.findings.extend(check_heading_pagination(output, headings, pagination,
                                                                 word_facts=facts))
                summary.findings.extend(check_chapter_page_breaks(
                    output, headings, pagination, template_breaks, word_facts=facts))
                summary.findings.extend(verify_page_numbering(output, pagination))
                summary.findings.extend(check_toc_layout(output, pagination,
                                                          word_facts=facts,
                                                          page_facts=pagination))
            except Exception as exc:
                summary.unknowns.append(f"【054/030/040】分页事实读取失败：{type(exc).__name__}；这些提示保持为无法判定")

            # 只带段号的结论补上文字摘录，供调度层按文字核对锚点位置。
            _backfill_excerpts(summary.findings, output, facts)
            # 调度层会自己把批注写进最终输出：这里的副本批注会被丢弃，
            # 因此默认跳过写入与保存（write_comments=False），省下一批 COM 调用与一次整篇保存。
            if write_comments:
                for finding in summary.findings:
                    add_comment(output, finding)
                output.Save()
            # 排版几何（题注位置/跨页表格/越界）：在**同一次会话**里顺手读出来，
            # 调度层之后不必再打开一次 Word。
            if layout_request and document_context is not None:
                try:
                    from thesis_format_checker.core.word_layout import read_layout_in_session

                    summary.render_layout = read_layout_in_session(
                        output, document_context,
                        paragraph_indices=layout_request.get("indices"),
                        want_figures=bool(layout_request.get("want_figures", True)),
                        want_tables=bool(layout_request.get("want_tables", True)),
                        facts=facts)
                except Exception as exc:
                    summary.unknowns.append(f"排版几何读取失败：{type(exc).__name__}；相关规则保持“无法判定”")
            session.close_document(output, save=False)
        return summary
    except Exception:
        if output_created and output_path.exists():
            try:
                output_path.unlink()
            except OSError:
                pass
        raise
