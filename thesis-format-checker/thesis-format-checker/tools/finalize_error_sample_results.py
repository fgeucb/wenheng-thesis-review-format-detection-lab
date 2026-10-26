"""对已完成的错误样本结果做无需重新分页的确定性精修。

本工具只做两类可审计修改：

* 把 022.1/022.2 的旧标题改成“与模板不一致”；
* 删除误把“目录”前置标题当成正文标题的 054.1 结论及其原生批注。

批注文档仅改 ``word/comments.xml`` 和对应批注锚点，随后重新执行完整性验证。
"""

from __future__ import annotations

from collections import Counter
from datetime import datetime
import importlib.util
import json
import os
from pathlib import Path
import re
import sys
from zipfile import ZipFile

from docx import Document
from lxml import etree

from thesis_format_checker.core.xmlutil import qn


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_RESULTS = ROOT / "samples" / "错误样本" / "检测结果_20261014_修正版"
TITLE_REPLACEMENTS = {
    "前置部分出现页眉": "前置部分页眉与模板不一致",
    "前置部分出现页脚": "前置部分页脚与模板不一致",
}
TOC_HEADINGS = {"目录", "contents", "图目录", "表目录", "插图目录", "附表目录"}


def _batch_tool():
    name = "run_error_samples_for_finalize"
    spec = importlib.util.spec_from_file_location(name, ROOT / "tools" / "run_error_samples.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def _compact(text: str) -> str:
    return re.sub(r"\s+", "", text or "").casefold()


def _comment_texts(path: Path) -> dict[str, str]:
    return {str(item.comment_id): item.text for item in Document(path).comments}


def _patch_comment_package(path: Path, remove_messages: set[str]) -> tuple[int, int]:
    """删除指定批注并替换标题；返回（删除数、标题替换数）。"""

    text_by_id = _comment_texts(path)
    remove_ids = {identity for identity, text in text_by_id.items() if text in remove_messages}
    replace_ids = {
        identity for identity, text in text_by_id.items()
        if any(old in text for old in TITLE_REPLACEMENTS)
    }
    if not remove_ids and not replace_ids:
        return 0, 0

    temp = path.with_name(path.stem + ".finalizing.docx")
    with ZipFile(path) as source, ZipFile(temp, "w") as target:
        for info in source.infolist():
            data = source.read(info.filename)
            if info.filename == "word/comments.xml":
                root = etree.fromstring(data)
                for comment in list(root.findall(qn("w:comment"))):
                    identity = comment.get(qn("w:id"))
                    if identity in remove_ids:
                        root.remove(comment)
                        continue
                    if identity in replace_ids:
                        for node in comment.iter(qn("w:t")):
                            if node.text:
                                for old, new in TITLE_REPLACEMENTS.items():
                                    node.text = node.text.replace(old, new)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            elif info.filename == "word/document.xml" and remove_ids:
                root = etree.fromstring(data)
                for tag in ("commentRangeStart", "commentRangeEnd"):
                    for node in list(root.iter(qn("w:" + tag))):
                        if node.get(qn("w:id")) in remove_ids:
                            node.getparent().remove(node)
                for node in list(root.iter(qn("w:commentReference"))):
                    if node.get(qn("w:id")) not in remove_ids:
                        continue
                    run = next((ancestor for ancestor in node.iterancestors(qn("w:r"))), None)
                    if run is not None:
                        run.getparent().remove(run)
                    else:
                        node.getparent().remove(node)
                data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
            target.writestr(info, data)
    os.replace(temp, path)
    return len(remove_ids), len(replace_ids)


def _replace_titles(value):
    if isinstance(value, str):
        for old, new in TITLE_REPLACEMENTS.items():
            value = value.replace(old, new)
        return value
    if isinstance(value, list):
        return [_replace_titles(item) for item in value]
    if isinstance(value, dict):
        return {key: _replace_titles(item) for key, item in value.items()}
    return value


def _recount(report: dict) -> None:
    findings = report.get("findings", [])
    report["counts_by_rule"] = dict(Counter(item.get("rule_id") for item in findings))
    report["counts_by_severity"] = dict(Counter(
        item.get("severity_label") for item in findings if item.get("status") == "fail"
    ))


def finalize(results: Path) -> dict:
    batch = _batch_tool()
    summary_path = results / "批量运行汇总.json"
    summary = json.loads(summary_path.read_text(encoding="utf-8"))
    records_by_report = {Path(item["report"]).resolve(): item for item in summary["records"]}
    changes = []
    previous_postprocess = summary.get("postprocess", {})

    for report_path in sorted(results.rglob("检测报告.json")):
        report = json.loads(report_path.read_text(encoding="utf-8"))
        removed = [
            item for item in report.get("findings", [])
            if item.get("status") == "fail" and item.get("rule_id") == "054.1"
            and _compact(item.get("excerpt", "")) in TOC_HEADINGS
        ]
        remove_messages = {item["message"] for item in removed}
        report["findings"] = [item for item in report.get("findings", []) if item not in removed]
        report = _replace_titles(report)

        folder = report_path.parent
        output = folder / "批注文档.docx"
        removed_comments, renamed_comments = _patch_comment_package(output, remove_messages)
        if removed_comments > len(remove_messages):
            raise RuntimeError(
                f"{folder}: 删除批注数 {removed_comments} 超过待删除结论数 {len(remove_messages)}"
            )

        _recount(report)
        comments = list(Document(output).comments)
        report["native_comment_count"] = len(comments)
        report["system_comment_count"] = sum(
            item.author == "论文格式检测" for item in comments
        )
        source = Path(report_path.parent.parent.parent / "错误样本")
        record = records_by_report.get(report_path.resolve())
        if record is None:
            raise RuntimeError(f"汇总中没有报告记录：{report_path}")
        source = Path(record["source"])
        verification = batch.verify_output_integrity(source, output, report=report)
        report["integrity_verification"] = verification
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        verification_path = folder / "完整性验证.json"
        verification_path.write_text(
            json.dumps(verification, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        if not verification["passed"]:
            raise RuntimeError(f"精修后完整性验证失败：{folder}")

        record.update({
            "finding_count": len(report.get("findings", [])),
            "fail_count": sum(item.get("status") == "fail" for item in report.get("findings", [])),
            "comment_count": len(comments),
            "system_comment_count": report["system_comment_count"],
            "integrity_passed": True,
        })
        log_path = Path(record["log"])
        note = (
            f"\n\n结果精修：删除目录标题误报 054.1 共 {len(removed)} 条；"
            f"更新 022 批注标题 {renamed_comments} 条；完整性复验通过。\n"
        )
        text = log_path.read_text(encoding="utf-8")
        if "结果精修：删除目录标题误报" not in text:
            log_path.write_text(text + note, encoding="utf-8")
        changes.append({
            "report": str(report_path),
            "removed_054_toc_findings": len(removed),
            "removed_comments": removed_comments,
            "removed_report_only_findings": len(removed) - removed_comments,
            "renamed_022_comments": renamed_comments,
        })

    summary["generated_at"] = datetime.now().isoformat(timespec="seconds")
    summary["postprocessed_at"] = summary["generated_at"]
    current_postprocess = {
        "description": "删除目录标题 054.1 误报并统一 022.1/022.2 标题",
        "documents": len(changes),
        "removed_054_toc_findings": sum(item["removed_054_toc_findings"] for item in changes),
        "removed_comments": sum(item["removed_comments"] for item in changes),
        "removed_report_only_findings": sum(item["removed_report_only_findings"] for item in changes),
        "renamed_022_comments": sum(item["renamed_022_comments"] for item in changes),
    }
    summary["postprocess"] = {
        key: (max(int(previous_postprocess.get(key, 0)), int(value))
              if key not in {"description"} else value)
        for key, value in current_postprocess.items()
    }
    summary["postprocess"]["latest_run"] = current_postprocess
    summary["completed_count"] = sum(item["status"] == "completed" for item in summary["records"])
    summary["failed_count"] = sum(item["status"] != "completed" for item in summary["records"])
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    audit_path = results / "结果精修汇总.json"
    if any(item["removed_054_toc_findings"] or item["renamed_022_comments"] for item in changes) \
            or not audit_path.exists():
        audit_path.write_text(json.dumps(changes, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary["postprocess"]


def main() -> None:
    results = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else DEFAULT_RESULTS.resolve()
    print(json.dumps(finalize(results), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
