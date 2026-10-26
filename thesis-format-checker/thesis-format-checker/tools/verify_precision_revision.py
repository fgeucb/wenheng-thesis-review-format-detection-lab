"""Verify the user's declaration, acronym and page examples against v1/v2."""
from collections import Counter
import json
from pathlib import Path
from hashlib import sha256
from docx import Document

folder = Path("samples/_revision_20261012_v2")
previous = Path("samples/_revision_20261012")
old = json.loads((previous / "report_修正版.json").read_text(encoding="utf-8"))
new = json.loads((folder / "report_修正版.json").read_text(encoding="utf-8"))
word = json.loads((folder / "word_verification.json").read_text(encoding="utf-8"))
key = lambda f: tuple(str(f[k]) for k in ("rule_id", "title", "status", "expected", "actual"))
before, after = Counter(map(key, old["findings"])), Counter(map(key, new["findings"]))
content_comments = [c for c in new["page_comments"] if set(c["rules"]) >= {"022.1", "029", "035"}]
assert [c["page"] for c in content_comments] == [6, 8]
assert all(c["story"] == "header" and c["length"] == 0 for c in content_comments)
assert not any("页眉（首页）内容变为" in f["actual"] for f in new["findings"]
               if f["rule_id"] == "029" and f["status"] == "fail")
assert {x["token"]: x["scope"] for x in word["acronym_scopes"] if x["page"] == 5} == {
    "AUC": "AUC", "F1": "F1"}
assert not any(f["rule_id"] == "149.1" and f["excerpt"] in {"GAT", "GNN"}
               for f in new["findings"])
assert len(word["declaration_scopes"]) == 2
assert word["declaration_scopes"][0]["scope"].strip() == "导师签名："
assert word["declaration_scopes"][1]["scope"].startswith("日期：")
assert all(x["page"] == 3 and "使用授权声明" in x["message"] for x in word["declaration_scopes"])
result = {
    "old_findings": len(old["findings"]), "new_findings": len(new["findings"]),
    "old_comments": len(Document(previous / "检测结果_修正版.docx").comments),
    "new_comments": len(Document(folder / "检测结果_修正版.docx").comments),
    "old_status": dict(Counter(f["status"] for f in old["findings"])),
    "new_status": dict(Counter(f["status"] for f in new["findings"])),
    "removed": [{"finding": k, "count": n} for k, n in (before - after).items()],
    "added": [{"finding": k, "count": n} for k, n in (after - before).items()],
    "visible_header_pages": [c["page"] for c in content_comments],
    "page_comments": len(new["page_comments"]),
    "complex_text_report_only": [x for x in new["report_only"] if "无损精确定位" in x],
    "v1_retained_sha256": {p.name: sha256(p.read_bytes()).hexdigest() for p in (
        previous / "检测结果_修正版.docx", previous / "report_修正版.json")},
}
(folder / "precision_verification.json").write_text(
    json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps(result, ensure_ascii=False, indent=2))
