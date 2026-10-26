from thesis_format_checker.detectors.word_effective.profile import _describe_indent, _describe_spacing, resolve_dominant


def test_resolve_dominant_accepts_clear_standard():
    value, reason = resolve_dominant([("宋体", 90), ("黑体", 10)])
    assert value == "宋体"
    assert reason is None


def test_resolve_dominant_rejects_conflicting_template():
    value, reason = resolve_dominant([("宋体", 60), ("黑体", 40)])
    assert value is None
    assert "冲突" in reason


def test_resolve_dominant_rejects_too_little_evidence():
    value, reason = resolve_dominant([("宋体", 1)])
    assert value is None
    assert reason == "样本不足"


def test_template_conflicts_describe_units_not_raw_tuples_or_rule_codes():
    _, indent_reason = resolve_dominant(
        [((24.0, 2.0), 6), ((0.0, 0.0), 4)], format_value=_describe_indent
    )
    assert "2 字符（约 24 磅）" in indent_reason
    assert "(24.0, 2.0)" not in indent_reason
    _, spacing_reason = resolve_dominant(
        [((1, 18.0, 0.0, 6.0), 6), ((4, 20.0, 0.0, 0.0), 4)],
        format_value=_describe_spacing,
    )
    assert "1.5 倍行距" in spacing_reason
    assert "固定值行距 20 磅" in spacing_reason
    assert "段后 6 磅" in spacing_reason
    assert "(1, 18.0" not in spacing_reason
