from thesis_format_checker.core.comments import add_comment
from thesis_format_checker.core.models import Finding, Status


class Range:
    def __init__(self, start, end, text="正文\r"):
        self.Start, self.End, self.Text = start, end, text


class Paragraph:
    def __init__(self):
        self.Range = Range(10, 15)


class Paragraphs:
    Count = 1

    def __call__(self, index):
        return Paragraph()


class Comments:
    def __init__(self):
        self.items = []

    def Add(self, **kwargs):
        self.items.append(kwargs)


class Document:
    def __init__(self):
        self.Paragraphs = Paragraphs()
        self.Comments = Comments()
        self.Content = Range(0, 20)

    def Range(self, start, end):
        return Range(start, end)


def test_comment_anchors_to_finding_span():
    document = Document()
    finding = Finding("069", "字号", Status.FAIL, "测试批注", paragraph_index=1, start=1, end=3)
    add_comment(document, finding)
    item = document.Comments.items[0]
    assert (item["Range"].Start, item["Range"].End) == (11, 13)
    assert item["Text"] == "测试批注"

