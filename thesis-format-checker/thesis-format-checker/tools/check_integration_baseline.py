"""Re-run the four failing integration fixtures with the untouched source snapshot."""
import json
import os
from pathlib import Path
import subprocess
import sys

root = Path(__file__).resolve().parents[1]
folder = root / "samples/_revision_20261012"
sys.path.insert(0, str(root / "tests"))
from docx import Document
from conftest import confirmed_heads, family_titles, title_families
from fixtures_consistency_appendix import EXPECTED
from test_cli_integration_full import _EXPECTED_RULES

cases = [
    ("citations", "812f63d9ca6f4621812a4e6e93752562", "模板.docx", "论文.docx", "批注.docx",
     ("105", "106", "107", "108", "110", "111", "113", "114", "115", "117", "118", "121", "122", "123", "124", "125", "128", "129")),
    ("consistency", "69d78ffc3afc4e2da76bb61aa436832d", "模板.docx", "论文.docx", "批注.docx", EXPECTED),
    ("figures", "8cae205c273b46fb812da692cd7ccc0e", "模板.docx", "论文.docx", "批注.docx",
     ("070", "071", "073", "076", "077", "078", "079", "080", "081", "083", "084", "090", "091", "092", "093", "096", "097")),
    ("full", "53285c79b0c542e5a50a700321dd3fb3", "标准模板.docx", "待检测文档.docx", "带原生批注.docx", _EXPECTED_RULES),
]
env = dict(os.environ, PYTHONPATH=str(folder / "src"), PYTHONIOENCODING="utf-8")
results = []
outdir = folder / "baseline_integration"
outdir.mkdir(exist_ok=True)
for name, fixture, template, paper, revised, expected in cases:
    inputs = root / "tests/_runtime" / fixture
    output = outdir / f"{name}.docx"
    with (outdir / f"{name}.log").open("wb") as log:
        completed = subprocess.run(
            [sys.executable, "-m", "thesis_format_checker", "compare",
             "--template", str(inputs / template), "--document", str(inputs / paper),
             "--output", str(output), "--report-json", str(outdir / f"{name}.json")],
            env=env, cwd=root, stdout=log, stderr=subprocess.STDOUT, timeout=240)
    assert completed.returncode == 0, name
    old_heads = confirmed_heads(Document(output).comments)
    new_heads = confirmed_heads(Document(inputs / revised).comments)
    missing = lambda heads: sorted(rule for rule in expected if not heads & family_titles(rule))
    if name == "full":
        # Reproduce this test's exact title -> family mapping (different from other tests).
        mapping = title_families()
        missing = lambda heads: sorted(set(expected) - {mapping.get(h) for h in heads})
    results.append({"case": name, "before_missing": missing(old_heads), "after_missing": missing(new_heads),
                    "confirmed_heads_removed": sorted(old_heads - new_heads),
                    "confirmed_heads_added": sorted(new_heads - old_heads)})
    print(results[-1], flush=True)
(folder / "baseline_integration_verification.json").write_text(
    json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
