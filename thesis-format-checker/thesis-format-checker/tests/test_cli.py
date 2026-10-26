from pathlib import Path

import pytest

from thesis_format_checker import cli
from thesis_format_checker.core.models import RunSummary


def test_cli_forwards_paths_and_prints_summary(monkeypatch, capsys, runtime_dir):
    called = {}

    def fake_compare(template, document, output, **kwargs):
        called.update(template=template, document=document, output=output)
        return RunSummary(sections_checked=1, body_paragraphs_checked=2, output_path=str(output))

    monkeypatch.setattr(cli, "compare_documents", fake_compare)
    result = cli.main(
        [
            "compare",
            "--template",
            str(runtime_dir / "a.docx"),
            "--document",
            str(runtime_dir / "b.doc"),
            "--output",
            str(runtime_dir / "out.docx"),
        ]
    )
    assert result == 0
    assert called["document"].suffix == ".doc"
    assert "已检查正文段落：2" in capsys.readouterr().out


def test_mapping_rejects_non_object(runtime_dir):
    path = runtime_dir / "mapping.yaml"
    path.write_text("- bad", encoding="utf-8")
    with pytest.raises(Exception, match="顶层必须是对象"):
        cli.load_mapping(path)
