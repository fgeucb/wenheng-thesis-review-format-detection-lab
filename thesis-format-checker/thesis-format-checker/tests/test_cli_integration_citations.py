"""可选的真实 Word 端到端测试：公共命令 + 原生批注。"""
import os
from pathlib import Path
import subprocess
import sys
from zipfile import ZipFile

import pytest
from docx import Document
from lxml import etree

from thesis_format_checker.core.point_comments import content_tree as semantic_tree
from conftest import assert_reported_comments


@pytest.mark.integration
def test_public_command_native_word_comments_and_unchanged_note_parts(runtime_dir):
    if os.environ.get("RUN_WORD_INTEGRATION") != "1" or os.name != "nt":
        pytest.skip("需设置 RUN_WORD_INTEGRATION=1 并使用Windows桌面Word")
    from fixtures_citations import fixture
    root = Path(__file__).resolve().parents[1]
    a, b, output = [runtime_dir / name for name in ("模板.docx", "论文.docx", "批注.docx")]
    fixture(a, template=True)
    fixture(b, template=False)
    before = [p.read_bytes() for p in (a, b)]
    report = runtime_dir / "report.json"
    for phase, args in (
        ("compare", [str(root / "thesis-checker.cmd"), "compare", "--template", str(a), "--document", str(b), "--output", str(output), "--report-json", str(report)]),
        ("reopen", [sys.executable, str(Path(__file__).with_name("test_cli_integration_headers_toc.py")), str(output)]),
    ):
        with (runtime_dir / (phase + "_stdout.log")).open("wb") as stdout, (runtime_dir / (phase + "_stderr.log")).open("wb") as stderr:
            try:
                process = subprocess.run(args, cwd=root, env=os.environ.copy(), stdout=stdout, stderr=stderr, timeout=120)
            except subprocess.TimeoutExpired:
                pytest.fail("本地Word测试超时；已停止等待，不自动重试，不终止用户Word实例")
        assert process.returncode == 0, (runtime_dir / (phase + "_stderr.log")).read_text(encoding="utf-8", errors="replace")[-2500:]
    assert before == [p.read_bytes() for p in (a, b)]
    comments = [c.text for c in Document(output).comments]
    expected = ("105", "106", "107", "108", "110", "111", "113", "114", "115", "117", "118", "121", "122", "123", "124", "125", "128", "129")
    assert_reported_comments(report, Document(output).comments, expected)
    assert "保留的人工批注" in comments
    assert not any("既有检测暂未完成" in c for c in comments)
    with ZipFile(b) as source, ZipFile(output) as annotated:
        allowed = {"word/document.xml", "word/comments.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
        for name in source.namelist():
            if name not in allowed:
                assert source.read(name) == annotated.read(name), name
        assert semantic_tree(etree.fromstring(source.read("word/document.xml"))) == semantic_tree(etree.fromstring(annotated.read("word/document.xml")))
