"""Audit old/new conclusions, ZIP parts and native Word anchors; never save inputs.

Run with the project's Python for --word (pywin32); use the bundled artifact
Python for the separate raster comparison tool.
"""
from collections import Counter
from hashlib import sha256
import argparse
import json
from pathlib import Path
from zipfile import ZipFile

from docx import Document
from lxml import etree

from thesis_format_checker.core.docx_context import DocxContext
from thesis_format_checker.core.point_comments import content_tree, marker_offset
from thesis_format_checker.core.xmlutil import qn


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def dump(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def audit(folder):
    source = Path("samples/paper.docx")
    original = Path("samples/_perf/检测结果_tier1_final.docx")
    output = folder / "检测结果_修正版.docx"
    old = json.loads(Path("samples/_perf/report_tier1_final.json").read_text(encoding="utf-8"))
    new = json.loads((folder / "report_修正版.json").read_text(encoding="utf-8"))
    source_doc, old_doc, new_doc = map(Document, (source, original, output))
    audit_inputs = (
        source, Path("samples/template.docx"), Path("samples/template_规范版.docx"),
        original, Path("samples/_perf/report_tier1_final.json"),
    )
    result = {"input_sha256": {str(p): digest(p) for p in audit_inputs if p.exists()}}
    result["comment_counts"] = {
        "old": len(old_doc.comments), "new": len(new_doc.comments),
        "old_unknown": sum("无法判定" in c.text for c in old_doc.comments),
        "new_unknown": sum("无法判定" in c.text for c in new_doc.comments),
        "new_location_prefix": sum(c.text.startswith("第 ") for c in new_doc.comments),
    }
    key = lambda f: (f["rule_id"], f["status"], str(f["expected"]), str(f["actual"]),
                     f["paragraph_index"], f["section_index"])
    old_keys, new_keys = Counter(map(key, old["findings"])), Counter(map(key, new["findings"]))
    result["conclusions"] = {
        "old_total": len(old["findings"]), "new_total": len(new["findings"]),
        "old_status": dict(Counter(f["status"] for f in old["findings"])),
        "new_status": dict(Counter(f["status"] for f in new["findings"])),
        "old_rule_counts": old["counts_by_rule"], "new_rule_counts": new["counts_by_rule"],
        "removed": [{"key": list(k), "count": n} for k, n in (old_keys - new_keys).items()],
        "added": [{"key": list(k), "count": n} for k, n in (new_keys - old_keys).items()],
    }
    # Separately compare decision content without location metadata.
    content_key = lambda f: key(f)[:4]
    before, after = Counter(map(content_key, old["findings"])), Counter(map(content_key, new["findings"]))
    result["conclusions"]["content_removed"] = [{"key": list(k), "count": n} for k, n in (before - after).items()]
    result["conclusions"]["content_added"] = [{"key": list(k), "count": n} for k, n in (after - before).items()]
    result["package"] = {}
    with ZipFile(source) as a, ZipFile(output) as b:
        allowed = {"word/document.xml", "word/comments.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
        changed = [n for n in a.namelist() if n not in b.namelist() or a.read(n) != b.read(n)]
        result["package"]["changed"] = changed
        result["package"]["added"] = sorted(set(b.namelist()) - set(a.namelist()))
        result["package"]["unexpected_changes"] = sorted(set(changed) - allowed)
        result["package"]["body_semantic_equal"] = content_tree(etree.fromstring(a.read("word/document.xml"))) == content_tree(etree.fromstring(b.read("word/document.xml")))
        result["package"]["style_parts_identical"] = all(a.read(n) == b.read(n) for n in a.namelist() if "styles" in n)
        result["package"]["media_count"] = len([n for n in a.namelist() if n.startswith("word/media/")])
        result["package"]["media_identical"] = all(a.read(n) == b.read(n) for n in a.namelist() if n.startswith("word/media/"))
    result["package"]["sections_equal"] = (
        [etree.tostring(s._sectPr, method="c14n") for s in source_doc.sections] ==
        [etree.tostring(s._sectPr, method="c14n") for s in new_doc.sections])
    result["package"]["section_count"] = len(source_doc.sections)
    comments = {str(c.comment_id): c.text for c in new_doc.comments}
    ctx = DocxContext(new_doc)
    markers = {tag: {n.get(qn("w:id")): n for n in new_doc.element.body.iter(qn("w:" + tag))}
               for tag in ("commentRangeStart", "commentRangeEnd", "commentReference")}
    result["anchors_complete"] = all(set(items) == set(comments) for items in markers.values())
    result["page_comments"] = []
    for item in new["page_comments"]:
        identifier = str(item["comment_id"])
        p = ctx.paragraphs[item["paragraph"] - 1]._p
        start, end = markers["commentRangeStart"][identifier], markers["commentRangeEnd"][identifier]
        row = dict(item, start=marker_offset(p, start), end=marker_offset(p, end),
                   text=comments[identifier])
        assert row["start"] == row["end"] == row["offset"]
        assert not p.xpath(".//w:drawing|.//w:pict|.//w:object")
        result["page_comments"].append(row)
    result["page_report_only"] = [
        {k: f.get(k) for k in ("rule_id", "section_index", "page_story", "page_scope",
                              "page_variant", "affected_pages", "expected", "actual", "location")}
        for f in new["findings"] if f["status"] == "fail" and f.get("page_story")
        and not any(f["rule_id"] in c.get("rules", [c["rule"]]) and set(f.get("affected_pages", [])) & set(c["affected_pages"])
                    for c in new["page_comments"])
    ]
    dump(folder / "verification.json", result)
    assert result["comment_counts"]["new_unknown"] == result["comment_counts"]["new_location_prefix"] == 0
    assert not result["package"]["unexpected_changes"]
    assert all(result["package"][k] for k in ("body_semantic_equal", "style_parts_identical", "media_identical", "sections_equal"))
    assert result["anchors_complete"]
    print(json.dumps({k: v for k, v in result.items() if k not in ("conclusions", "page_comments")}, ensure_ascii=False, indent=2))
    return new, result


def word_audit(folder, report):
    from thesis_format_checker.core.word_session import WordSession
    result = {"files": {}, "anchors": [], "acronym_scopes": [], "declaration_scopes": []}
    qa = folder / "word_render"
    qa.mkdir(exist_ok=True)
    with WordSession() as session:
        result["word_version"] = session.app.Version
        for label, path in (("original", Path("samples/paper.docx")),
                            ("revised", folder / "检测结果_修正版.docx")):
            wd = session.open_readonly(path.resolve())
            session.resume_pagination()
            wd.Windows(1).View.Type = 3
            wd.Windows(1).View.ShowRevisionsAndComments = False
            wd.Repaginate()
            result["files"][label] = {
                "pages": wd.ComputeStatistics(2), "sections": wd.Sections.Count,
                "comments": wd.Comments.Count,
                "paragraphs": wd.Paragraphs.Count, "inline_shapes": wd.InlineShapes.Count,
                "shapes": wd.Shapes.Count,
            }
            if label == "revised":
                # Word orders Comments by scope, not XML id. Field-code character positions
                # can change on opening: match the comment text, then verify empty scope,
                # main story and physical page; retain both offsets for the audit.
                for index in range(1, wd.Comments.Count + 1):
                    comment = wd.Comments(index)
                    scope = comment.Scope
                    start, end = int(scope.Start), int(scope.End)
                    text = comment.Range.Text.replace("\r", "\n").rstrip("\n")
                    if "缩略语首次出现" in text:
                        import re
                        token = re.search(r"实际：“([^”]+)”", text).group(1)
                        assert scope.Text == token, (token, scope.Text)
                        result["acronym_scopes"].append({"token": token, "scope": scope.Text,
                                                        "page": int(scope.Information(3))})
                    if "使用授权声明缺少" in text:
                        result["declaration_scopes"].append({"message": text, "scope": scope.Text,
                                                            "page": int(scope.Information(3))})
                    if start == end:
                        candidates = report["page_comments"]
                        for expected in candidates:
                            if text != expected["text"]:
                                continue
                            page = int(scope.Information(3))
                            if page != expected["page"] or any(
                                    a["comment_id"] == expected["comment_id"] for a in result["anchors"]):
                                continue
                            result["anchors"].append({"comment_id": expected["comment_id"],
                                                      "start": start, "end": end, "page": page,
                                                      "snapshot_position": expected["position"],
                                                      "expected_page": expected["page"],
                                                      "story_type": int(scope.StoryType)})
                            break
                assert len(result["anchors"]) == len(report["page_comments"]), result["anchors"]
                assert all(a["page"] == a["expected_page"] and a["story_type"] == 1 for a in result["anchors"])
            wd.ExportAsFixedFormat(OutputFileName=str((qa / (label + ".pdf")).resolve()),
                                   ExportFormat=17, OpenAfterExport=False, Item=0)
            session.close_document(wd, save=False)
    before, after = result["files"]["original"], result["files"]["revised"]
    assert all(before[k] == after[k] for k in ("pages", "sections", "paragraphs", "inline_shapes", "shapes"))
    dump(folder / "word_verification.json", result)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", type=Path, default=Path("samples/_revision_20261012"))
    parser.add_argument("--word", action="store_true")
    args = parser.parse_args()
    _, result = audit(args.folder)
    if args.word:
        word_audit(args.folder, result)
