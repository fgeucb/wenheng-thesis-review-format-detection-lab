from pathlib import Path

import pytest

from thesis_format_checker.core.word_session import WordAutomationError, WordSession


class BrokenDocuments:
    def Open(self, **kwargs):
        raise RuntimeError("corrupt")


class BrokenApp:
    Documents = BrokenDocuments()


def test_missing_input_has_clear_error(runtime_dir):
    with pytest.raises(WordAutomationError, match="输入文件不存在"):
        WordSession._validate_input(runtime_dir / "missing.docx")


def test_unsupported_input_extension_is_rejected(runtime_dir):
    path = runtime_dir / "paper.pdf"
    path.write_bytes(b"x")
    with pytest.raises(WordAutomationError, match="仅支持"):
        WordSession._validate_input(path)


def test_corrupt_document_error_is_wrapped(runtime_dir):
    path = runtime_dir / "broken.docx"
    path.write_bytes(b"not a docx")
    session = WordSession()
    session.app = BrokenApp()
    with pytest.raises(WordAutomationError, match="可能损坏"):
        session.open_readonly(path)


def test_existing_output_requires_overwrite(runtime_dir):
    source = runtime_dir / "source.docx"
    output = runtime_dir / "output.docx"
    source.write_bytes(b"source")
    output.write_bytes(b"existing")
    session = WordSession()
    with pytest.raises(WordAutomationError, match="--overwrite"):
        session.create_output_copy(source, output, overwrite=False)


def test_output_cannot_be_input(runtime_dir):
    source = runtime_dir / "source.docx"
    source.write_bytes(b"source")
    session = WordSession()
    with pytest.raises(WordAutomationError, match="不会覆盖输入原件"):
        session.create_output_copy(source, source, overwrite=True)

