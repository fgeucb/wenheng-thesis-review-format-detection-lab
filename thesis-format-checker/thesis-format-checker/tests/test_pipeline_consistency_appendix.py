from zipfile import ZipFile

import pytest
from docx import Document
from lxml import etree

from thesis_format_checker.engine.pipeline import run_comparison, semantic_tree
from thesis_format_checker.core.docx_context import DocxContext, UNREADABLE, make_finding
from thesis_format_checker.core.models import RunSummary, Status
from thesis_format_checker.checks.catalog import (
    DEFERRED_RULE_IDS,
    FINE_RULES,
    PENDING_RULE_IDS,
    REMOVED_RULE_IDS,
)
from thesis_format_checker.engine.pipeline import SECTION_RULES, check_sections, extract_baseline
from conftest import comment_heads, family_titles


def family_of(rule_id):
    """细粒度规则编号 → 所属规则族（036.1 → 036）。"""

    return rule_id.split(".")[0]



def test_exact_coverage_and_append_only_baseline():
    """大类十三、十四、十五（133–165）的 33 个规则族全部有归属。"""
    sections = ("consistency", "formats", "appendices", "hygiene", "language_lite")
    covered = {rule for section in sections for rule in SECTION_RULES[section]
               if 133 <= int(rule) <= 165}
    assert len(covered) == 16      # hygiene：154/160/165；language_lite：144/149/151
    removed_here = {rule for rule in REMOVED_RULE_IDS if 133 <= int(rule) <= 165}
    pending_here = {rule for rule in PENDING_RULE_IDS if 133 <= int(rule) <= 165}
    deferred_here = {rule for rule in DEFERRED_RULE_IDS if 133 <= int(rule) <= 165}
    assert pending_here == set()
    assert deferred_here == {"141", "161"}
    # 大类十四：154 与 144/149/151 的可确定性部分已实现，其余保持已删除。
    assert {f"{i:03}" for i in range(144, 155)} - removed_here == {"144", "149", "151", "154"}
    families = covered | removed_here | pending_here | deferred_here
    assert families == {f"{i:03}" for i in range(133, 166)}
    assert len(families) == 33
    ctx = DocxContext(Document())
    old = {"earlier": {"preserved": True}}
    data = extract_baseline(ctx, old)
    assert data["earlier"] is old["earlier"]
    assert set(old) == {"earlier"}
    assert {"consistency", "formats", "appendices"} <= set(data)
    with pytest.raises(ValueError, match="未覆盖"):
        extract_baseline(ctx, data)


def test_failures_are_isolated_and_no_deferred_or_language_business(monkeypatch):
    import thesis_format_checker.detectors.appendices as appendices
    import thesis_format_checker.detectors.consistency as consistency
    import thesis_format_checker.detectors.formats as formats

    def broken(*args, **kwargs):
        raise ValueError("unreadable object")

    monkeypatch.setattr(consistency, "check_consistency_errors", broken)
    monkeypatch.setattr(formats, "check_format_consistency_errors", broken)
    monkeypatch.setattr(appendices, "check_appendix_errors", broken)
    ctx = DocxContext(Document())
    sections = ("consistency", "formats", "appendices")
    baseline = extract_baseline(ctx, sections=sections)
    findings = check_sections(baseline, ctx, sections=sections)
    assert len(findings) == 10
    assert all(f.status == Status.UNKNOWN and UNREADABLE in f.message for f in findings)

    def working_check(*args, **kwargs):
        return [make_finding("134", "题目一致性", "内部一致性", "两处题目不同")]

    monkeypatch.setattr(consistency, "check_consistency_errors", working_check)
    findings = check_sections(baseline, ctx, sections=sections)
    assert any(f.rule_id == "134" and f.status == Status.FAIL for f in findings)
    assert sum(f.status == Status.UNKNOWN for f in findings) == 7
    assert not (set(REMOVED_RULE_IDS) | set(PENDING_RULE_IDS) | set(DEFERRED_RULE_IDS)) & {f.rule_id for f in findings}


def test_pipeline_retains_existing_rules_and_artifacts_without_text_changes(runtime_dir):
    a, b, out = [runtime_dir / name for name in ("template.docx", "paper.docx", "out.docx")]
    for path, title in ((a, "模板示例题目"), (b, "论文实际题目")):
        doc = Document()
        doc.add_paragraph("论文题目：" + title)
        doc.add_paragraph("摘要")
        doc.add_paragraph("论文题目：" + title)
        doc.add_heading("第一章 绪论", 1)
        doc.add_paragraph("正文格式保留 English 123。")
        doc.add_comment(doc.paragraphs[-1].runs, "既有人工批注", author="人工")
        doc.save(path)
    before = [p.read_bytes() for p in (a, b)]
    calls = []
    def legacy(*args, **kwargs):
        calls.append(args)
        finding = make_finding("123", "保留旧检测", "原标准", "既有标点证据")
        finding.paragraph_index = 5   # 指向正文段：旧流程的结论都带段号
        return RunSummary(findings=[finding])
    summary = run_comparison(legacy, a, b, out)
    assert len(calls) == 1 and "123" in summary.counts_by_rule()
    assert "134" not in summary.counts_by_rule()  # Never require the sample title.
    assert not (set(REMOVED_RULE_IDS) | set(PENDING_RULE_IDS) | set(DEFERRED_RULE_IDS)) & {f.rule_id for f in summary.findings}
    # 未实现 / 待确认的**细粒度规则**必须逐条出现在跳过说明里
    pending_codes = [item.code for item in FINE_RULES.values()
                     if item.status.value in {"pending", "deferred"}]
    assert pending_codes
    assert all(any(f"【{code}】" in note for note in summary.skipped) for code in pending_codes)
    assert any("大类十四" in note for note in summary.skipped)
    assert before == [p.read_bytes() for p in (a, b)]
    comments = list(Document(out).comments)
    assert comments[0].text == "既有人工批注"
    assert any("既有标点证据" in c.text for c in comments)   # 123 的“实际”值
    with ZipFile(b) as original, ZipFile(out) as annotated:
        allowed = {"word/document.xml", "word/comments.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
        for name in original.namelist():
            if name not in allowed:
                assert original.read(name) == annotated.read(name), name
        assert semantic_tree(etree.fromstring(original.read("word/document.xml"))) == semantic_tree(etree.fromstring(annotated.read("word/document.xml")))


def test_combined_fixture_reports_all_ten_rules_and_preserves_package(runtime_dir):
    from fixtures_consistency_appendix import fixture, EXPECTED
    a, b, out = [runtime_dir / name for name in ("template.docx", "paper.docx", "out.docx")]
    fixture(a, template=True)
    fixture(b, template=False)
    before = [p.read_bytes() for p in (a, b)]
    summary = run_comparison(lambda *a, **k: RunSummary(), a, b, out)
    failed = {f.rule_id for f in summary.findings if f.status == Status.FAIL}
    # 细粒度拆分后按规则族比较（139.2、155.1 …）
    failed_families = {rule_id.split(".")[0] for rule_id in failed}
    assert EXPECTED <= failed_families, [
        (f.rule_id, f.status.value, f.actual) for f in summary.findings if family_of(f.rule_id) in EXPECTED
    ]
    assert before == [p.read_bytes() for p in (a, b)]
    comments = list(Document(out).comments)
    assert comments[0].text == "保留的人工批注"
    heads = comment_heads(comments)
    for rule in EXPECTED:
        # 批注不再写规则编号：按该族结论自身的标题核对（批注抬头用的就是这个标题）。
        wanted = {f.title for f in summary.findings if family_of(f.rule_id) == rule}
        assert heads & wanted, rule
    for comment in comments[1:]:
        assert "应该：" in comment.text and "实际：" in comment.text
        # 批注已锚定在出问题的文字上，正文里不再重复写“位置：”一行。
        assert "位置：" not in comment.text
    with ZipFile(b) as original, ZipFile(out) as annotated:
        allowed = {"word/document.xml", "word/comments.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
        for name in original.namelist():
            if name not in allowed:
                assert original.read(name) == annotated.read(name), name
        assert semantic_tree(etree.fromstring(original.read("word/document.xml"))) == semantic_tree(etree.fromstring(annotated.read("word/document.xml")))
