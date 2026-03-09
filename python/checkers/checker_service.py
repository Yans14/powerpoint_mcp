from __future__ import annotations

from typing import Any

from engines.base import BaseEngine
from errors import BridgeError


class CheckerService:
    def __init__(self, engine: BaseEngine) -> None:
        self.engine = engine

    def _prs(self, presentation_id: str):
        session = self.engine.sessions.get(presentation_id)
        if not session:
            raise BridgeError(
                code="not_found",
                message=f"Presentation '{presentation_id}' not found.",
            )
        return session.extra["prs"]

    def dispatch(self, method: str, params: dict[str, Any]) -> Any:
        handlers = {
            "pptx_check_positions": self._check_positions,
            "pptx_check_visual_consistency": self._check_visual,
            "pptx_check_content": self._check_content,
            "pptx_check_template_conformance": self._check_template,
            "pptx_diff_presentations": self._diff,
        }
        handler = handlers.get(method)
        if handler is None:
            raise BridgeError(code="not_found", message=f"Unknown checker method: {method}")
        return handler(params)

    def _check_positions(self, params: dict[str, Any]) -> dict[str, Any]:
        from checkers.position_checker import PositionChecker

        prs = self._prs(str(params["presentation_id"]))
        all_indices = list(range(1, len(prs.slides) + 1))
        indices = [int(i) for i in params.get("slide_indices", [])] or all_indices
        return PositionChecker().check(
            prs,
            slide_indices=indices,
            check_overlaps=bool(params.get("check_overlaps", True)),
            check_bounds=bool(params.get("check_bounds", True)),
            check_alignment=bool(params.get("check_alignment", False)),
            tolerance_px=int(params.get("tolerance_px", 5)),
        )

    def _check_visual(self, params: dict[str, Any]) -> dict[str, Any]:
        from checkers.visual_checker import VisualConsistencyChecker

        prs = self._prs(str(params["presentation_id"]))
        all_indices = list(range(1, len(prs.slides) + 1))
        indices = [int(i) for i in params.get("slide_indices", [])] or all_indices
        return VisualConsistencyChecker().check(
            prs,
            slide_indices=indices,
            check_fonts=bool(params.get("check_fonts", True)),
            check_colors=bool(params.get("check_colors", True)),
            check_sizes=bool(params.get("check_sizes", True)),
        )

    def _check_content(self, params: dict[str, Any]) -> dict[str, Any]:
        from checkers.content_checker import ContentChecker

        prs = self._prs(str(params["presentation_id"]))
        indices = [int(i) for i in params.get("slide_indices", [])] or None
        return ContentChecker().check(
            prs,
            check_empty=bool(params.get("check_empty_placeholders", True)),
            check_default_text=bool(params.get("check_default_text", True)),
            slide_indices=indices,
        )

    def _check_template(self, params: dict[str, Any]) -> dict[str, Any]:
        from checkers.template_checker import TemplateConformanceChecker
        from utils.paths import validate_existing_file

        prs = self._prs(str(params["presentation_id"]))
        template_path = validate_existing_file(str(params["template_path"]), expected_suffixes=(".pptx", ".potx"))
        return TemplateConformanceChecker().check(
            prs,
            template_path=str(template_path),
            check_theme=bool(params.get("check_theme", True)),
            check_fonts=bool(params.get("check_fonts", True)),
            check_layouts=bool(params.get("check_layouts", True)),
        )

    def _diff(self, params: dict[str, Any]) -> dict[str, Any]:
        from checkers.diff import PresentationDiffer

        prs_a = self._prs(str(params["presentation_id_a"]))
        prs_b = self._prs(str(params["presentation_id_b"]))
        return PresentationDiffer().diff(prs_a, prs_b, deep_diff=bool(params.get("deep_diff", True)))
