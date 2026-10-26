from __future__ import annotations

import os
from pathlib import Path

import pytest

from thesis_format_checker.engine.pipeline import compare_documents
from thesis_format_checker.core.word_session import WordSession


def _make_document(app, path: Path, *, compliant: bool, old_doc: bool = False):
    doc = app.Documents.Add()
    try:
        paragraph = doc.Paragraphs.Add()
        paragraph.Range.Text = "这是用于检测的正文内容，包含 English 123 混排文字。"
        paragraph.Range.Style = -1  # wdStyleNormal；不依赖 Office 界面语言
        paragraph.Range.Font.NameFarEast = "宋体" if compliant else "微软雅黑"
        paragraph.Range.Font.NameAscii = "Times New Roman" if compliant else "Arial"
        paragraph.Range.Font.NameOther = "Times New Roman" if compliant else "Arial"
        paragraph.Range.Font.Size = 12 if compliant else 10
        paragraph.Format.CharacterUnitFirstLineIndent = 2 if compliant else 0
        section = doc.Sections(1)
        # Pure unit conversion avoids a broken optional COM utility method on
        # some Word builds; actual Word document operations remain exercised.
        section.PageSetup.TopMargin = 25 * 72 / 25.4
        section.PageSetup.BottomMargin = 25 * 72 / 25.4
        section.PageSetup.LeftMargin = (30 if compliant else 20) * 72 / 25.4
        section.PageSetup.RightMargin = 30 * 72 / 25.4
        doc.SaveAs2(FileName=str(path), FileFormat=0 if old_doc else 16)
    finally:
        doc.Close(SaveChanges=False)


@pytest.mark.integration
def test_doc_input_and_native_word_comments_round_trip(runtime_dir):
    if os.environ.get("RUN_WORD_INTEGRATION") != "1":
        pytest.skip("设置 RUN_WORD_INTEGRATION=1 后运行真实 Microsoft Word 集成测试")
    template = runtime_dir / "template.docx"
    document = runtime_dir / "document.doc"
    output = runtime_dir / "annotated.docx"
    try:
        with WordSession() as session:
            _make_document(session.app, template, compliant=True)
            _make_document(session.app, document, compliant=False, old_doc=True)
    except Exception as exc:
        pytest.skip(f"当前环境无法启动 Microsoft Word：{exc}")

    summary = compare_documents(template, document, output)
    assert output.is_file()
    # 页边距等按细粒度规则输出（025.1–025.4），这里按规则族统计
    families = {rule_id.split(".")[0] for rule_id in summary.counts_by_rule()}
    assert {"025", "067", "068", "069", "072"}.issubset(families)

    # 重新通过 Word 打开结果，验证真正的原生批注，而不是只检查内存对象。
    with WordSession() as session:
        result = session.open_readonly(output)
        assert result.Comments.Count >= 5
        texts = [result.Comments(index).Range.Text for index in range(1, result.Comments.Count + 1)]
        for rule_id in ("025", "067", "068", "069", "072"):
            failures = [f for f in summary.findings if f.rule_id.split(".")[0] == rule_id
                        and f.status.value == "fail"]
            assert failures
            normalize = lambda text: " ".join(text.split())
            assert any(normalize(f.message) == normalize(text) for f in failures for text in texts), (rule_id, [f.message for f in failures], texts)
        first_scope = result.Comments(1).Scope
        assert first_scope.Start < first_scope.End
