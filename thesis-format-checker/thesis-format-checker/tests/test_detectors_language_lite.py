"""不依赖大模型的文字规范检测：错别字词表、重复字、缩略语全称、数字写法、英文文献大小写。"""

from __future__ import annotations

from docx import Document

from thesis_format_checker.core.docx_context import DocxContext
from thesis_format_checker.core.models import Status
from thesis_format_checker.detectors.language_lite import check_language_errors


def _context(*texts, heading_style=False):
    document = Document()
    for text in texts:
        paragraph = document.add_paragraph(text)
        if heading_style and text == "参考文献":
            paragraph.style = document.styles["Heading 1"]
    return DocxContext(document)


def _find(context, family):
    return [finding for finding in check_language_errors({}, context)
            if finding.rule_id.split(".")[0] == family]


def test_common_typo_is_reported_with_exact_range():
    context = _context("本文提出的方法做为一种新思路，值得推广。")
    findings = _find(context, "144")
    assert [finding.rule_id for finding in findings] == ["144.1"]
    assert "作为" in findings[0].expected and "做为" in findings[0].actual
    assert findings[0].start == 7 and findings[0].end == 9   # “做为”的字符范围


def test_duplicate_function_word_is_reported():
    context = _context("实验结果表明该方法的的精度更高。")
    findings = _find(context, "144")
    assert [finding.rule_id for finding in findings] == ["144.2"]
    assert "的的" in findings[0].actual


def test_legitimate_reduplication_is_not_reported():
    context = _context("我们常常可以看到，数据会渐渐趋于稳定，请看看结果。")
    assert _find(context, "144") == []


def test_acronym_first_use_without_full_name_is_a_reminder():
    # 全文出现过两次及以上的缩略语才提醒（只出现一次更可能是单位/型号/编号）。
    context = _context("本文采用LSTM模型完成预测，LSTM具有记忆能力。")
    findings = _find(context, "149")
    assert len(findings) == 1
    assert findings[0].rule_id == "149.1"


def test_single_occurrence_acronym_is_not_reported():
    """只出现一次的英文大写词（单位、型号、编号）不提醒，避免噪声。"""

    context = _context("设备型号为 XPS13，测试温度 25℃。")
    assert _find(context, "149") == []


def test_acronym_with_full_name_passes():
    context = _context("本文采用长短期记忆网络（LSTM）完成预测。LSTM可用于预测。")
    assert _find(context, "149") == []


def test_explained_gat_does_not_hide_unexplained_auc_and_f1():
    context = _context("引入动态注意力机制与图注意力网络（GAT）实现个股表征的自适应聚合。"
                       "实验结果表明该模型在AUC与F1-score指标上显著优于传统基线方法。",
                       "GAT模型的AUC与F1指标如下。")
    findings = _find(context, "149")
    assert {f.excerpt for f in findings} == {"AUC", "F1"}
    for f in findings:
        assert f.anchor.text[f.start:f.end] == f.excerpt


def test_mixed_year_styles_are_reported_once():
    context = _context("项目始于2018年，相关成果发表于二〇二一年。")
    findings = _find(context, "151")
    assert len(findings) == 1 and findings[0].rule_id == "151.1"
    assert "年份写法应统一" in findings[0].expected


def test_uniform_year_style_passes():
    context = _context("项目始于2018年，相关成果发表于2021年。")
    assert _find(context, "151") == []


def test_english_reference_author_case_mix_is_reported():
    context = _context(
        "参考文献",
        "FAMA E F, FRENCH K R. Common risk factors in the returns on stocks and bonds[J]. "
        "Journal of Financial Economics, 1993, 33(1): 3-56.",
        "Gu S, Kelly B, Xiu D. Empirical asset pricing via machine learning[J]. "
        "The Review of Financial Studies, 2020, 33(5): 2223-2273.",
        "DIEBOLD F X, YILMAZ K. On the network topology of variance decompositions[J]. "
        "Journal of Econometrics, 2014, 182(1): 119-134.",
        heading_style=True,
    )
    findings = _find(context, "119")
    assert len(findings) == 1 and findings[0].rule_id == "119.1"
    assert "大小写" in findings[0].expected
    assert findings[0].status == Status.FAIL
