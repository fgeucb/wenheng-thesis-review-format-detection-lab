"""Task-local verification: reuse freshly measured facts, reread page boundaries.

This is not a production cache bypass. Inputs are pinned to the sample hashes;
the explicit evidence paths and hashes are recorded for the verification report.
"""
import json
import argparse
from hashlib import sha256
from pathlib import Path
from docx import Document
from thesis_format_checker.core.cache import decode, encode, summary_from_payload
from thesis_format_checker.core.docx_context import DocxContext
from thesis_format_checker.core.word_layout import read_pagination
from thesis_format_checker.engine.pipeline import run_comparison
from thesis_format_checker.engine.reporting import write_json_report

parser = argparse.ArgumentParser()
parser.add_argument("--folder", type=Path, default=Path("samples/_revision_20261012"))
parser.add_argument("--reuse-measured-pagination", action="store_true")
args = parser.parse_args()
folder = args.folder
folder.mkdir(parents=True, exist_ok=True)
paper = Path("samples/paper.docx")
template = Path("samples/template_规范版.docx")
assert sha256(paper.read_bytes()).hexdigest() == "7857274e1e1827545b435741fe90eb509ecabe79bd4da32fea41478d0d8864da"
assert sha256(template.read_bytes()).hexdigest() == "7906f463f86b4506e3eb2f3e12688b2dfb8b9895e3702594d85f082993276716"
cache = Path("C:/Users/Lenovo/AppData/Local/thesis-checker/cache")
flow_path = cache / "word-flow/cb0938c2a163f89cc3c713912c92987f.json"
layout_path = cache / "layout/dc1e34f609b0903494ea3ebd8675b99d.json"
flow = decode(json.loads(flow_path.read_text(encoding="utf-8"))["data"])
layout = decode(json.loads(layout_path.read_text(encoding="utf-8"))["data"])
pagination_path = folder / "measured_pagination.json"
if args.reuse_measured_pagination:
    layout["pagination"] = decode(json.loads(pagination_path.read_text(encoding="utf-8")))
else:
    print("Rereading actual Word page boundaries", flush=True)
    layout["pagination"] = read_pagination(paper, DocxContext(Document(paper)))
    pagination_path.write_text(json.dumps(encode(layout["pagination"]), ensure_ascii=False), encoding="utf-8")

def evidence(*args, **kwargs):
    result = summary_from_payload(flow)
    result.render_layout = layout
    return result

summary = run_comparison(evidence, template, paper, folder / "检测结果_修正版.docx",
                         overwrite=True, use_cache=False)
write_json_report(summary, folder / "report_修正版.json")
(folder / "replay_provenance.json").write_text(json.dumps({
    "method": "Reuse Word findings/geometry measured in this task; rerun static detectors and all annotations",
    "pagination": "reuse task measurement" if args.reuse_measured_pagination else "fresh Word read",
    "evidence": {str(p): sha256(p.read_bytes()).hexdigest() for p in (flow_path, layout_path)},
    "findings": len(summary.findings),
}, ensure_ascii=False, indent=2), encoding="utf-8")
print("Finished", len(summary.findings), flush=True)
