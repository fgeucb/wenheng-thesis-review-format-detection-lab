"""生成可立即试跑的模板 A 和待检测文档 B（需要 Microsoft Word）。"""
from pathlib import Path

from thesis_format_checker.core.word_session import WordSession


def make(app, path: Path, compliant: bool, old_doc: bool = False) -> None:
    doc = app.Documents.Add()
    try:
        paragraph = doc.Paragraphs.Add()
        paragraph.Range.Text = "这是一段示例正文，包含 English 123，用于验证格式批注。"
        paragraph.Range.Style = -1  # wdStyleNormal；不依赖 Office 界面语言
        paragraph.Range.Font.NameFarEast = "宋体" if compliant else "微软雅黑"
        paragraph.Range.Font.NameAscii = "Times New Roman" if compliant else "Arial"
        paragraph.Range.Font.Size = 12 if compliant else 10
        paragraph.Format.CharacterUnitFirstLineIndent = 2 if compliant else 0
        setup = doc.Sections(1).PageSetup
        setup.TopMargin = app.MillimetersToPoints(25)
        setup.BottomMargin = app.MillimetersToPoints(25)
        setup.LeftMargin = app.MillimetersToPoints(30 if compliant else 20)
        setup.RightMargin = app.MillimetersToPoints(30)
        doc.SaveAs2(FileName=str(path), FileFormat=0 if old_doc else 16)
    finally:
        doc.Close(SaveChanges=False)


def main() -> None:
    output = Path(__file__).parent / "generated"
    output.mkdir(exist_ok=True)
    with WordSession() as session:
        make(session.app, output / "template_a.docx", True)
        make(session.app, output / "document_b.doc", False, old_doc=True)
    print(f"示例已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
