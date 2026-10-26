"""对照校验：Word 段落事实的“逐段读取”与“批量读取”是否逐项一致。

背景（性能优化 Tier 0）
======================

旧版 Word 流程会**整篇文档走四遍**：正文段、标题、分页事实、空段扫描，每遍都要
逐段访问 COM（约 15 次调用/段）。优化方案是"一遍建表"：一次 ``Document.Content.Text``
取回全文，段落样式/是否在表格/大纲级别从 OOXML 侧读，页码按页边界 + 字符位置二分
换算。**前提是两种读法得到的值必须完全一致**，因此提供本工具做对照：

* ``--slow``：逐段读（现状做法）——每段的文本、起始位置、页码、节号、样式、是否在表格；
* ``--fast``：批量读（优化做法）——``Content.Text`` 切分 + 累计字符位置 + 页边界二分；
* 逐项 diff，并打印调用次数与耗时；任何不一致都会列出前若干条明细。

用法::

    python tools/verify_word_facts.py samples/paper.docx
    python tools/verify_word_facts.py samples/paper.docx --limit 300   # 只比对前 300 段

结论会写进终端输出；工具只读文档，不修改任何文件。
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from thesis_format_checker.core.word_session import WordSession  # noqa: E402


class Counter:
    """COM 调用计数（把 dict 读取包一层，统计读了多少次）。"""

    def __init__(self) -> None:
        self.calls = 0

    def __call__(self, *_args):
        self.calls += 1


def read_slow(document, limit: int | None, counter: Counter) -> list[dict]:
    """逐段读取（现状做法）。"""

    total = int(document.Paragraphs.Count)
    if limit:
        total = min(total, limit)
    records = []
    for index in range(1, total + 1):
        paragraph = document.Paragraphs(index)
        counter()
        text = str(paragraph.Range.Text)
        counter()
        start = int(paragraph.Range.Start)
        counter()
        end = int(paragraph.Range.End)
        counter()
        page = int(paragraph.Range.Information(3))
        counter()
        section = int(paragraph.Range.Information(2))
        counter()
        try:
            style = str(paragraph.Range.Style.NameLocal or "")
        except Exception:
            style = ""
        counter()
        try:
            in_table = bool(paragraph.Range.Information(12))
        except Exception:
            in_table = False
        counter()
        try:
            outline = int(paragraph.OutlineLevel)
        except Exception:
            outline = None
        counter()
        records.append({"index": index, "text": text, "start": start, "end": end,
                        "page": page, "section": section, "style": style,
                        "in_table": in_table, "outline": outline})
    return records


def read_fast(document, limit: int | None, counter: Counter) -> dict:
    """批量读取（优化做法）：一次取全文 + 页边界二分。"""

    count = int(document.Paragraphs.Count)
    counter(0)
    content_start = int(document.Content.Start)
    counter()
    raw = document.Content.Text
    counter()
    pieces = raw.split("\r")
    counter()
    if pieces and pieces[-1] == "":
        pieces.pop()
    # 段落起始位置 = 内容起点 + 前面各段长度 + 每段的段落标记
    positions = []
    position = content_start
    for piece in pieces:
        positions.append(position)
        position += len(piece) + 1
    # 页边界
    pages = int(document.ComputeStatistics(2))
    counter()
    end = int(document.Content.End)
    counter()
    starts = []
    for number in range(1, pages + 1):
        # wdGoToPage + wdGoToAbsolute：一次调用拿到该页首字符位置
        position_ = int(document.GoTo(1, 1, number).Start)
        counter()
        starts.append(position_)
    starts.append(end)
    import bisect
    computed = []
    for index, position_ in enumerate(positions, start=1):
        if limit and index > limit:
            break
        page = max(1, bisect.bisect_right(starts, position_))
        computed.append({"index": index, "text": pieces[index - 1] + "\r",
                         "start": position_, "page": page})
    return {"count": count, "pieces": len(pieces), "pages": pages, "starts": starts,
            "records": computed}


def compare(slow: list[dict], fast: dict, show: int = 20) -> int:
    """逐项比较，返回不一致条数。"""

    problems = 0
    detail: list[str] = []
    records = {item["index"]: item for item in fast["records"]}
    for item in slow:
        other = records.get(item["index"])
        if other is None:
            continue
        if item["text"] != other["text"]:
            problems += 1
            if len(detail) < show:
                detail.append(f"段 {item['index']} 文本不同：\n    慢={item['text']!r}\n    快={other['text']!r}")
            continue
        if item["start"] != other["start"]:
            problems += 1
            if len(detail) < show:
                detail.append(f"段 {item['index']} 起始位置不同：慢={item['start']} 快={other['start']}")
        if item["page"] != other["page"]:
            problems += 1
            if len(detail) < show:
                detail.append(f"段 {item['index']} 页码不同：慢={item['page']} 快={other['page']}")
    print(f"段落数：Word={fast['count']}，批量文本切出={fast['pieces']}，"
          f"{'一致' if fast['count'] == fast['pieces'] else '不一致（需要回退逐段读取）'}")
    print(f"总页数：{fast['pages']}")
    print(f"不一致：{problems} 处")
    for line in detail:
        print("  -", line)
    return problems


def verify_production(document, limit: int | None, show: int = 20) -> int:
    """校验**正式流程**用的事实表（WordFacts + PageFacts）与逐段读取一致。

    这是"提速不掉精度"的验收项：文本、起始位置、样式、大纲级别、是否在表格、
    页码、节号、节页范围、空白页逐项比对，全部一致才算通过。
    """

    from thesis_format_checker.core.word_facts import WordFacts
    from thesis_format_checker.core.word_pages import PageFacts, read_page_facts

    reference_counter = Counter()
    slow = read_slow(document, limit, reference_counter)
    facts = WordFacts(document, label="probe")
    pagination = PageFacts(document, facts)
    problems = 0
    detail: list[str] = []
    for item in slow:
        index = item["index"]
        checks = (
            ("文本", facts.text(index), item["text"]),
            ("起始位置", facts.start(index), item["start"]),
            ("样式", facts.style_name(index), item["style"]),
            ("大纲级别", facts.outline_level(index), item["outline"]),
            ("是否在表格", facts.in_table(index), item["in_table"]),
            ("页码", pagination.page_of(index), item["page"]),
            ("节号", pagination.section_of(index), item["section"]),
        )
        for name, found, expected in checks:
            if found != expected:
                problems += 1
                if len(detail) < show:
                    detail.append(f"段 {index} {name}：事实表={found!r} 逐段={expected!r}")
    plain = read_page_facts(document)
    if pagination.total_pages != plain["total_pages"]:
        problems += 1
        detail.append(f"总页数：事实表={pagination.total_pages} 逐段={plain['total_pages']}")
    if pagination.sections != plain["sections"]:
        problems += 1
        detail.append(f"节页范围：事实表={pagination.sections} 逐段={plain['sections']}")
    # 空白页口径：**有段落的页绝不能被判成空白页**（安全性约束）；反向允许更严格
    # （例：只有一张图片的页，旧口径不算内容、新口径算内容）。
    with_paragraphs = {page for page in (plain.get("pages") or {}).values() if page}
    mine = set(pagination.content_pages())
    wrong_blank = sorted(with_paragraphs - mine)
    if wrong_blank:
        problems += 1
        detail.append(f"这些页有段落却被判为空白页：{wrong_blank}")
    print(f"事实表读取：{facts.calls} 次 COM（段落属性），分页 {pagination.calls} 次")
    print(f"逐段参考：{reference_counter.calls} 次 COM")
    print(f"空白页（新口径）：{pagination.blank_pages()}；"
          f"无段落的页（旧口径）：{sorted(set(range(1, int(pagination.total_pages or 0) + 1)) - with_paragraphs)}")
    print(f"不一致：{problems} 处")
    for line in detail:
        print("  -", line)
    return problems


def main() -> int:
    parser = argparse.ArgumentParser(description="Word 段落事实读取方式对照校验")
    parser.add_argument("document", type=Path, help="待校验的 .docx")
    parser.add_argument("--limit", type=int, default=0, help="只比对前 N 段（0=全部）")
    parser.add_argument("--production", action="store_true",
                        help="校验正式流程的 WordFacts/PageFacts（而不是 Content.Text 实验）")
    args = parser.parse_args()
    limit = args.limit or None

    slow_counter, fast_counter = Counter(), Counter()
    with WordSession() as session:
        document = session.open_readonly(args.document.resolve())
        # 对照校验必须读"真实页码"：正式流程在读格式阶段会关闭后台重排，
        # 这里恢复分页后再读，否则 Information(3)/(2) 拿到的是无效值。
        session.resume_pagination()
        if args.production:
            started = time.perf_counter()
            problems = verify_production(document, limit)
            print(f"耗时：{time.perf_counter() - started:.1f} 秒")
            return 1 if problems else 0
        started = time.perf_counter()
        slow = read_slow(document, limit, slow_counter)
        slow_seconds = time.perf_counter() - started
        started = time.perf_counter()
        fast = read_fast(document, limit, fast_counter)
        fast_seconds = time.perf_counter() - started
    print(f"逐段读取：{len(slow)} 段，约 {slow_counter.calls} 次 COM 访问，{slow_seconds:.1f} 秒")
    print(f"批量读取：{fast['pieces']} 段，约 {fast_counter.calls} 次 COM 访问，{fast_seconds:.1f} 秒")
    return 1 if compare(slow, fast) else 0


if __name__ == "__main__":
    raise SystemExit(main())
