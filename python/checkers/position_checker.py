from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from pptx import Presentation

from utils.units import emu_to_inches


@dataclass
class PositionIssue:
    slide_index: int
    shape_a_name: str
    shape_b_name: str | None
    issue_type: str
    description: str
    severity: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "slide_index": self.slide_index,
            "shape_a": self.shape_a_name,
            "shape_b": self.shape_b_name,
            "issue_type": self.issue_type,
            "description": self.description,
            "severity": self.severity,
        }


def _rects_overlap(
    a: tuple[int, int, int, int],
    b: tuple[int, int, int, int],
    tolerance: int,
) -> bool:
    return not (
        a[2] <= b[0] + tolerance or b[2] <= a[0] + tolerance or a[3] <= b[1] + tolerance or b[3] <= a[1] + tolerance
    )


class PositionChecker:
    def check(
        self,
        prs: Presentation,
        slide_indices: list[int],
        check_overlaps: bool = True,
        check_bounds: bool = True,
        check_alignment: bool = False,
        tolerance_px: int = 5,
    ) -> dict[str, Any]:
        tolerance_emu = tolerance_px * 9525
        sw = prs.slide_width
        sh = prs.slide_height
        issues: list[PositionIssue] = []

        for idx in slide_indices:
            slide = prs.slides[idx - 1]
            shapes = list(slide.shapes)

            if check_bounds:
                for shape in shapes:
                    left, top = shape.left, shape.top
                    right, bottom = left + shape.width, top + shape.height
                    if (
                        left < -tolerance_emu
                        or top < -tolerance_emu
                        or right > sw + tolerance_emu
                        or bottom > sh + tolerance_emu
                    ):
                        issues.append(
                            PositionIssue(
                                slide_index=idx,
                                shape_a_name=shape.name,
                                shape_b_name=None,
                                issue_type="out_of_bounds",
                                description=(
                                    f"Shape '{shape.name}' extends outside slide boundaries "
                                    f"(left={emu_to_inches(left):.2f}in, top={emu_to_inches(top):.2f}in, "
                                    f"right={emu_to_inches(right):.2f}in, bottom={emu_to_inches(bottom):.2f}in). "
                                    f"Slide is {emu_to_inches(sw):.2f}in x {emu_to_inches(sh):.2f}in."
                                ),
                                severity="critical",
                            )
                        )

            if check_overlaps:
                for i, sa in enumerate(shapes):
                    rect_a = (sa.left, sa.top, sa.left + sa.width, sa.top + sa.height)
                    for sb in shapes[i + 1 :]:
                        rect_b = (sb.left, sb.top, sb.left + sb.width, sb.top + sb.height)
                        if _rects_overlap(rect_a, rect_b, tolerance_emu):
                            issues.append(
                                PositionIssue(
                                    slide_index=idx,
                                    shape_a_name=sa.name,
                                    shape_b_name=sb.name,
                                    issue_type="overlap",
                                    description=f"Shapes '{sa.name}' and '{sb.name}' overlap.",
                                    severity="warning",
                                )
                            )

        critical = sum(1 for i in issues if i.severity == "critical")
        warnings = sum(1 for i in issues if i.severity == "warning")
        return {
            "issues": [i.to_dict() for i in issues],
            "summary": {
                "total_issues": len(issues),
                "critical": critical,
                "warnings": warnings,
                "slides_checked": len(slide_indices),
            },
        }
