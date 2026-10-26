"""Exhaustive per-finding field diff against the retained tier1 report."""
from collections import Counter, defaultdict, deque
import json
from pathlib import Path

folder = Path("samples/_revision_20261012")
old = json.loads(Path("samples/_perf/report_tier1_final.json").read_text(encoding="utf-8"))
new = json.loads((folder / "report_修正版.json").read_text(encoding="utf-8"))
key = lambda f: tuple(str(f[k]) for k in ("rule_id", "status", "expected", "actual"))
pool = defaultdict(deque)
for index, finding in enumerate(old["findings"]):
    pool[key(finding)].append((index + 1, finding))
changes = []
for index, finding in enumerate(new["findings"]):
    old_index, previous = pool[key(finding)].popleft()
    fields = {name: {"before": previous.get(name), "after": finding.get(name)}
              for name in sorted(previous.keys() | finding.keys())
              if previous.get(name) != finding.get(name)}
    if fields:
        changes.append({"old_index": old_index, "new_index": index + 1,
                        "rule_id": finding["rule_id"], "changes": fields})
assert not any(pool.values())
result = {
    "conclusion_content_identical": True,
    "changed_finding_count": len(changes),
    "changed_field_counts": dict(Counter(k for row in changes for k in row["changes"])),
    "findings": changes,
    "top_level_changed_fields": [k for k in sorted(old.keys() | new.keys()) if old.get(k) != new.get(k)],
}
(folder / "report_diff.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({k: v for k, v in result.items() if k != "findings"}, ensure_ascii=False, indent=2))
