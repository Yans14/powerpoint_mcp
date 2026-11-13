"""Read-only discovery of layouts, masters and theme."""

from __future__ import annotations

from pathlib import Path
from typing import Any


class OOXMLDiscoveryMixin:
    """Read-only discovery of layouts, masters and theme."""

    def get_layouts(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        layouts: list[dict[str, Any]] = []
        for idx, layout in enumerate(prs.slide_layouts, start=1):
            placeholder_types = []
            for placeholder in layout.placeholders:
                placeholder_types.append(str(placeholder.placeholder_format.type))

            layouts.append(
                {
                    "index": idx,
                    "name": layout.name,
                    "placeholder_types": placeholder_types,
                }
            )

        return {
            "presentation_id": session.id,
            "layouts": layouts,
        }

    def get_layout_detail(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        layout_name = str(params["layout_name"])
        layout = self._find_layout(prs, layout_name)

        placeholders = [self._placeholder_payload(placeholder) for placeholder in layout.placeholders]

        return {
            "presentation_id": session.id,
            "layout_name": layout.name,
            "placeholder_count": len(placeholders),
            "placeholders": placeholders,
        }

    def get_masters(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        masters = []
        for idx, master in enumerate(prs.slide_masters, start=1):
            masters.append(
                {
                    "index": idx,
                    "name": getattr(master, "name", f"Master {idx}"),
                    "layout_count": len(master.slide_layouts),
                }
            )

        return {
            "presentation_id": session.id,
            "masters": masters,
        }

    def get_theme(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        self.flush_session(session)
        theme_data = self._extract_theme(Path(session.working_path))

        return {
            "presentation_id": session.id,
            "theme": theme_data,
        }
