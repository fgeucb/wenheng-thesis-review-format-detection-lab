"""Read actual Word page-boundary records for the sample; no source writes."""
from pathlib import Path
import json
from docx import Document
from thesis_format_checker.core.docx_context import DocxContext
from thesis_format_checker.core.word_layout import _paragraph_records, paragraph_matches, _page_start, _page_at, utf16_offsets
from thesis_format_checker.core.point_comments import safe_paragraph
from thesis_format_checker.core.word_session import WordSession

ctx = DocxContext(Document("samples/paper.docx"))
rows = []
with WordSession() as session:
    wd = session.open_readonly(Path("samples/paper.docx").resolve())
    session.resume_pagination()
    records = _paragraph_records(wd)
    matches = paragraph_matches(ctx.texts, [r[2] for r in records])
    end = wd.Content.End
    starts = [_page_start(wd, n, end) for n in range(1, 12)]
    for n in range(3, 10):
        a, b = starts[n - 1:n + 1]
        candidates = [i for i, (s, e, _) in enumerate(records) if s < b and e > a]
        items = []
        for i in dict.fromkeys(candidates[:2] + candidates[-2:]):
            s, e, text = records[i]
            x = matches.get(i)
            p = ctx.paragraphs[x] if x is not None else None
            items.append({"word": i + 1, "xml": x + 1 if x is not None else None, "start": s, "end": e,
                          "text": text[:100], "tail": text[-15:], "utf16_length": utf16_offsets(text)[-1],
                          "safe": safe_paragraph(p) if p is not None else None,
                          "hidden": wd.Range(s, e).Font.Hidden, "first_page": _page_at(wd, s),
                          "last_page": _page_at(wd, e - 1)})
        rows.append({"page": n, "start": a, "stop": b, "items": items})
path = Path("samples/_revision_20261012/boundary_probe.json")
path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
print(path)
