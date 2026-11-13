"""Read-only discovery of layouts, masters and theme."""

from __future__ import annotations

from typing import Any


class COMDiscoveryMixin:
    """Read-only discovery of layouts, masters and theme."""

    def get_layouts(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        self._get_session(presentation_id)
        layouts = self.worker.call(self._get_layouts_threadsafe, presentation_id)
        return {"presentation_id": presentation_id, "layouts": layouts}

    def _get_layouts_threadsafe(self, session_id: str) -> list[dict[str, Any]]:
        prs = self._require_presentation(session_id)
        layouts: list[dict[str, Any]] = []
        master = prs.SlideMasters(1)
        for i in range(1, master.CustomLayouts.Count + 1):
            layout = master.CustomLayouts(i)
            placeholder_types = []
            try:
                phs = layout.Shapes.Placeholders
                for p in range(1, phs.Count + 1):
                    placeholder_types.append(str(phs(p).PlaceholderFormat.Type))
            except Exception:
                pass
            layouts.append({"index": i, "name": layout.Name, "placeholder_types": placeholder_types})
        return layouts

    def get_layout_detail(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        layout_name = str(params["layout_name"])
        detail = self.worker.call(self._get_layout_detail_threadsafe, presentation_id, layout_name)
        return {"presentation_id": presentation_id, **detail}

    def _get_layout_detail_threadsafe(self, session_id: str, layout_name: str) -> dict[str, Any]:
        layout = self._find_layout_by_name(session_id, layout_name)
        placeholders = []
        try:
            phs = layout.Shapes.Placeholders
            for i in range(1, phs.Count + 1):
                ph = phs(i)
                placeholders.append(self._placeholder_payload(ph))
        except Exception:
            pass
        return {
            "layout_name": layout_name,
            "placeholder_count": len(placeholders),
            "placeholders": placeholders,
        }

    def get_masters(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        masters = self.worker.call(self._get_masters_threadsafe, presentation_id)
        return {"presentation_id": presentation_id, "masters": masters}

    def _get_masters_threadsafe(self, session_id: str) -> list[dict[str, Any]]:
        prs = self._require_presentation(session_id)
        masters = []
        for i in range(1, prs.SlideMasters.Count + 1):
            master = prs.SlideMasters(i)
            masters.append(
                {
                    "index": i,
                    "name": getattr(master, "Name", f"Master {i}"),
                    "layout_count": master.CustomLayouts.Count,
                }
            )
        return masters

    def get_theme(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        theme = self.worker.call(self._get_theme_threadsafe, presentation_id)
        return {"presentation_id": presentation_id, "theme": theme}

    def _get_theme_threadsafe(self, session_id: str) -> dict[str, Any]:
        prs = self._require_presentation(session_id)
        colors: dict[str, str] = {}
        fonts: dict[str, str] = {"major": "", "minor": ""}
        try:
            scheme = prs.SlideMaster.Theme.ThemeColorScheme
            for i in range(1, scheme.Count + 1):
                colors[f"theme_{i}"] = str(scheme.Colors(i).RGB)
        except Exception:
            pass
        try:
            fonts["major"] = str(prs.SlideMaster.Theme.ThemeFontScheme.MajorFont(1).Name)
        except Exception:
            pass
        try:
            fonts["minor"] = str(prs.SlideMaster.Theme.ThemeFontScheme.MinorFont(1).Name)
        except Exception:
            pass
        return {"colors": colors, "fonts": fonts}
