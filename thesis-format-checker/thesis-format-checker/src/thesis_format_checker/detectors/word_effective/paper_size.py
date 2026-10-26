from __future__ import annotations

from thesis_format_checker.core.models import Finding, Status
from thesis_format_checker.detectors.word_effective.profile import points_to_mm


def compare_paper_size(section, section_index: int, tolerance_mm: float = 0.5) -> Finding | None:
    setup = section.PageSetup
    width = points_to_mm(setup.PageWidth)
    height = points_to_mm(setup.PageHeight)
    if (abs(width - 210) <= tolerance_mm and abs(height - 297) <= tolerance_mm) or (
        abs(width - 297) <= tolerance_mm and abs(height - 210) <= tolerance_mm
    ):
        return None
    return Finding(
        "023", "纸张大小", Status.FAIL,
        f"【023 纸张大小】第 {section_index} 节应为 A4（210 × 297 mm，可横向）；当前为 {width:g} × {height:g} mm。",
        expected="A4 210 × 297 mm", actual=f"{width:g} × {height:g} mm", section_index=section_index,
    )

