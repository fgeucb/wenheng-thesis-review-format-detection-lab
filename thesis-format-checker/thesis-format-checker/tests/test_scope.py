from thesis_format_checker.core.scope import front_end_of, front_matter_end, style_is_excluded


def test_excludes_heading_and_toc_styles():
    assert style_is_excluded("标题 1")
    assert style_is_excluded("TOC 2")
    assert style_is_excluded("Abstract")


def test_keeps_body_styles():
    assert not style_is_excluded("正文")
    assert not style_is_excluded("Normal")


def test_front_matter_end_stops_at_first_abstract_or_contents():
    assert front_matter_end(["封面", "原创性声明", "摘  要", "正文"]) == 2
    assert front_matter_end(["使用授权声明", "ABSTRACT", "目录", "正文"]) == 1


def test_front_matter_end_requires_cover_or_declaration_evidence():
    """正文里恰好有个叫“摘要”的小节时不能把前面全部当前置区裁掉。"""

    assert front_matter_end(["摘要", "正文"]) == 0
    assert front_matter_end(["第一章 绪论", "摘  要", "1.1 小节"]) == 0
    assert front_matter_end(["正文第一段", "参考文献"]) == 0


def test_front_matter_end_recognizes_cover_without_cover_word():
    """封面页通常不出现“封面”二字，而是学校名、毕业设计、题目、姓名、学号等要素。"""

    texts = ["南京农业大学", "本科生毕业论文（设计）", "题    目：基于××的研究",
             "姓    名：张三", "学    号：2022010", "学    院：智慧农业学院",
             "指导教师：李四    职称：教授", "摘  要", "针对××问题……", "第一章 绪论"]
    assert front_matter_end(texts) == 7


def test_front_end_of_prefers_cached_context_value():
    class WithProperty:
        front_matter_end = 3
        paragraphs = []

    assert front_end_of(WithProperty()) == 3

    class Plain:
        paragraphs = [type("P", (), {"text": "论文原创性声明"})(),
                      type("P", (), {"text": "摘  要"})()]

    assert front_end_of(Plain()) == 1
