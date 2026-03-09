from __future__ import annotations

import base64
import os
import re
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path
from typing import Any

from com_worker import COMWorker
from engines.base import BaseEngine, EngineSession
from errors import BridgeError, ensure
from utils.colors import normalize_color
from utils.paths import validate_existing_file, validate_output_file
from utils.units import to_emu

try:  # pragma: no cover - Windows only
    import win32com.client  # type: ignore[import]
except Exception:  # pragma: no cover
    win32com = None  # type: ignore[assignment]


def _hex_to_bgr_int(color_value: str) -> int:
    color = normalize_color(color_value)
    rr = int(color[0:2], 16)
    gg = int(color[2:4], 16)
    bb = int(color[4:6], 16)
    return (bb << 16) | (gg << 8) | rr


def _to_points(value) -> float:
    """Convert any unit string ('2in', '24pt', EMU int) to PowerPoint points."""
    return to_emu(value) / 12700


# MSO AutoShapeType constants
_MSO_SHAPE_MAP: dict[str, int] = {
    "rectangle": 1,
    "rounded_rectangle": 5,
    "oval": 9,
    "diamond": 4,
    "triangle": 7,
    "right_arrow": 33,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "pentagon": 56,
    "hexagon": 10,
    "chevron": 52,
    "star_5_point": 12,
    "line_inverse": 183,
    "cross": 11,
    "frame": 158,
    "rectangular_callout": 105,
    "rounded_rectangular_callout": 106,
    "cloud_callout": 108,
    "cloud": 179,
}

# MSO Connector type constants
_MSO_CONNECTOR_MAP: dict[str, int] = {
    "straight": 1,
    "elbow": 2,
    "curve": 3,
}

# Excel chart type constants (for COM AddChart2 / AddChart)
_XL_CHART_TYPE_MAP: dict[str, int] = {
    "column_clustered": 51,
    "column_stacked": 52,
    "column_stacked_100": 53,
    "bar_clustered": 57,
    "bar_stacked": 58,
    "bar_stacked_100": 59,
    "line": 4,
    "line_markers": 65,
    "line_stacked": 63,
    "pie": 5,
    "pie_exploded": 69,
    "doughnut": -4120,
    "area": 1,
    "area_stacked": 76,
    "area_stacked_100": 77,
    "xy_scatter": -4169,
    "xy_scatter_lines": 74,
    "xy_scatter_smooth": 72,
    "bubble": 15,
    "radar": -4151,
    "stock_hlc": 88,
    "stock_ohlc": 89,
    "three_d_column": 54,
    "three_d_bar_clustered": 60,
    "three_d_pie": -4102,
    "three_d_line": -4101,
}

_XL_CHART_TYPE_REVERSE: dict[int, str] = {v: k for k, v in _XL_CHART_TYPE_MAP.items()}

# Alignment map (ppAlign constants)
_ALIGN_MAP: dict[str, int] = {
    "left": 1,
    "center": 2,
    "right": 3,
    "justify": 4,
}

# Z-order commands (MsoZOrderCmd)
_ZORDER_MAP: dict[str, int] = {
    "front": 0,  # msoBringToFront
    "back": 1,  # msoSendToBack
    "forward": 2,  # msoBringForward
    "backward": 3,  # msoSendBackward
}

# COM legend position constants
_LEGEND_POS_MAP: dict[str, int] = {
    "bottom": -4107,
    "corner": 2,
    "left": -4131,
    "right": -4152,
    "top": -4160,
}


class COMEngine(BaseEngine):
    name = "COM"

    def __init__(self, metadata: dict[str, Any]) -> None:  # pragma: no cover - Windows only
        if win32com is None:
            raise BridgeError(
                code="dependency_missing",
                message="pywin32 is not installed. COM mode requires pywin32 on Windows.",
            )

        self.metadata = metadata
        self.worker = COMWorker()
        self.sessions: dict[str, EngineSession] = {}
        self._presentations: dict[str, Any] = {}
        self._app = self.worker.call(self._init_app)

    def _init_app(self):  # pragma: no cover - Windows only
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = True
        return app

    def get_engine_info(
        self, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:  # pragma: no cover - Windows only
        return {
            "engine": self.name,
            **self.metadata,
        }

    def create_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        session_id = str(uuid.uuid4())
        width = params.get("width")
        height = params.get("height")
        template_path = params.get("template_path")

        if template_path:
            source = validate_existing_file(str(template_path), expected_suffixes=(".pptx", ".potx"))
            fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
            os.close(fd)
            Path(tmp_path).unlink(missing_ok=True)
            shutil.copy2(source, tmp_path)
            self.worker.call(self._open_presentation_for_session, session_id, tmp_path)
            original_path = str(source)
        else:
            fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
            os.close(fd)
            Path(tmp_path).unlink(missing_ok=True)
            self.worker.call(self._create_blank_presentation_for_session, session_id, tmp_path, width, height)
            original_path = ""

        self.sessions[session_id] = EngineSession(
            id=session_id,
            original_path=original_path,
            working_path=tmp_path,
            engine=self.name,
            dirty=True,
        )

        return {
            "success": True,
            "presentation_id": session_id,
            "engine": self.name,
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def _create_blank_presentation_for_session(self, session_id: str, tmp_path: str, width, height) -> None:
        prs = self._app.Presentations.Add()

        if width is not None:
            prs.PageSetup.SlideWidth = to_emu(width) / 12700
        if height is not None:
            prs.PageSetup.SlideHeight = to_emu(height) / 12700

        prs.SaveAs(tmp_path, 24)
        self._presentations[session_id] = prs

    def _open_presentation_for_session(self, session_id: str, path: str) -> None:
        prs = self._app.Presentations.Open(path, WithWindow=False)
        self._presentations[session_id] = prs

    def open_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        source = validate_existing_file(str(params["file_path"]), expected_suffixes=(".pptx", ".potx"))

        fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
        os.close(fd)
        Path(tmp_path).unlink(missing_ok=True)
        shutil.copy2(source, tmp_path)

        session_id = str(uuid.uuid4())
        self.worker.call(self._open_presentation_for_session, session_id, tmp_path)

        self.sessions[session_id] = EngineSession(
            id=session_id,
            original_path=str(source),
            working_path=tmp_path,
            engine=self.name,
            dirty=False,
        )

        layout_names = self.worker.call(self._list_layout_names, session_id)

        return {
            "success": True,
            "presentation_id": session_id,
            "slide_count": self.worker.call(self._slide_count, session_id),
            "layout_names": layout_names,
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def save_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        session = self._get_session(str(params["presentation_id"]))
        output_path = validate_output_file(str(params["output_path"]), expected_suffixes=(".pptx",))
        self.worker.call(self._save_copy_as, session.id, str(output_path))
        session.dirty = False

        return {
            "success": True,
            "presentation_id": session.id,
            "saved_path": str(output_path),
            "file_size_bytes": output_path.stat().st_size,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def _save_copy_as(self, session_id: str, output_path: str) -> None:
        prs = self._require_presentation(session_id)
        prs.Save()
        prs.SaveCopyAs(output_path)

    def close_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        session = self._get_session(presentation_id)
        self.worker.call(self._close_presentation_threadsafe, presentation_id)
        self.sessions.pop(presentation_id, None)
        try:
            Path(session.working_path).unlink(missing_ok=True)
        except Exception:
            pass

        return {
            "success": True,
            "presentation_id": presentation_id,
            "closed": True,
        }

    def _close_presentation_threadsafe(self, session_id: str) -> None:
        prs = self._presentations.pop(session_id, None)
        if prs is not None:
            prs.Close()

    def list_open_presentations(self, params: dict[str, Any] | None = None) -> dict[str, Any]:  # pragma: no cover
        return {
            "presentations": [
                {
                    "presentation_id": session.id,
                    "original_path": session.original_path,
                    "working_path": session.working_path,
                    "engine": session.engine,
                    "dirty": session.dirty,
                    "opened_at": session.opened_at.isoformat(),
                }
                for session in self.sessions.values()
            ]
        }

    def get_presentation_state(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        self._get_session(presentation_id)
        slides = self.worker.call(self._presentation_state_slides, presentation_id)
        return {
            "presentation_id": presentation_id,
            "engine": self.name,
            "slide_count": len(slides),
            "slides": slides,
        }

    def _presentation_state_slides(self, session_id: str) -> list[dict[str, Any]]:
        prs = self._require_presentation(session_id)
        items: list[dict[str, Any]] = []
        for i in range(1, prs.Slides.Count + 1):
            slide = prs.Slides(i)
            items.append(
                {
                    "index": i,
                    "title": self._slide_title(slide),
                    "layout": slide.CustomLayout.Name if slide.CustomLayout else "",
                    "shape_count": slide.Shapes.Count,
                }
            )
        return items

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

    def _list_layout_names(self, session_id: str) -> list[str]:
        return [item["name"] for item in self._get_layouts_threadsafe(session_id)]

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

    def get_slide(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        data = self.worker.call(self._get_slide_threadsafe, presentation_id, slide_index)
        return {"presentation_id": presentation_id, **data}

    def _get_slide_threadsafe(self, session_id: str, slide_index: int) -> dict[str, Any]:
        slide = self._slide(session_id, slide_index)
        placeholders = []
        try:
            phs = slide.Shapes.Placeholders
            for i in range(1, phs.Count + 1):
                placeholders.append(self._placeholder_payload(phs(i)))
        except Exception:
            pass

        shapes = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            shapes.append(self._shape_payload(shape))

        return {
            "slide_index": slide_index,
            "title": self._slide_title(slide),
            "layout": slide.CustomLayout.Name if slide.CustomLayout else "",
            "shapes": shapes,
            "placeholders": placeholders,
            "notes": self._notes_text(slide),
        }

    def add_slide(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        layout_name = str(params["layout_name"])
        position = params.get("position")

        result = self.worker.call(self._add_slide_threadsafe, presentation_id, layout_name, position)
        return {
            "success": True,
            "presentation_id": presentation_id,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }

    def _add_slide_threadsafe(self, session_id: str, layout_name: str, position: int | None) -> dict[str, Any]:
        prs = self._require_presentation(session_id)
        layout = self._find_layout_by_name(session_id, layout_name)

        if position is None:
            position = prs.Slides.Count + 1
        ensure(1 <= int(position) <= prs.Slides.Count + 1, "validation_error", "position is out of bounds")

        slide = prs.Slides.AddSlide(int(position), layout)
        placeholders = []
        try:
            phs = slide.Shapes.Placeholders
            for i in range(1, phs.Count + 1):
                placeholders.append(self._placeholder_payload(phs(i)))
        except Exception:
            pass

        return {
            "added_slide_index": int(position),
            "layout_name": layout_name,
            "placeholders": placeholders,
        }

    def duplicate_slide(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        source_index = int(params["source_index"])
        target_position = params.get("target_position")

        duplicated_slide_index = self.worker.call(
            self._duplicate_slide_threadsafe,
            presentation_id,
            source_index,
            target_position,
        )

        return {
            "success": True,
            "presentation_id": presentation_id,
            "source_index": source_index,
            "duplicated_slide_index": duplicated_slide_index,
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }

    def _duplicate_slide_threadsafe(self, session_id: str, source_index: int, target_position: int | None) -> int:
        prs = self._require_presentation(session_id)
        ensure(1 <= source_index <= prs.Slides.Count, "validation_error", "source_index is out of bounds")

        duplicate_range = prs.Slides(source_index).Duplicate()
        duplicate_slide = duplicate_range(1)

        if target_position is not None:
            ensure(
                1 <= int(target_position) <= prs.Slides.Count, "validation_error", "target_position is out of bounds"
            )
            duplicate_slide.MoveTo(int(target_position))
            return int(target_position)

        return int(source_index + 1)

    def delete_slide(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        self.worker.call(self._delete_slide_threadsafe, presentation_id, slide_index)

        return {
            "success": True,
            "presentation_id": presentation_id,
            "deleted_slide_index": slide_index,
            "warning": f"Slide indices above {slide_index} shifted down by 1.",
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }

    def _delete_slide_threadsafe(self, session_id: str, slide_index: int) -> None:
        prs = self._require_presentation(session_id)
        ensure(1 <= slide_index <= prs.Slides.Count, "validation_error", "slide_index out of bounds")
        prs.Slides(slide_index).Delete()

    def reorder_slides(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        new_order = [int(value) for value in params["new_order"]]
        self.worker.call(self._reorder_slides_threadsafe, presentation_id, new_order)
        return {
            "success": True,
            "presentation_id": presentation_id,
            "new_order": new_order,
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }

    def _reorder_slides_threadsafe(self, session_id: str, new_order: list[int]) -> None:
        prs = self._require_presentation(session_id)
        count = prs.Slides.Count
        expected = list(range(1, count + 1))
        ensure(len(new_order) == count, "validation_error", "new_order length must equal slide_count")
        ensure(sorted(new_order) == expected, "validation_error", "new_order must include each index once")

        original_ids = [prs.Slides(i).SlideID for i in range(1, count + 1)]
        desired_ids = [original_ids[index - 1] for index in new_order]

        for target_position, slide_id in enumerate(desired_ids, start=1):
            current_index = None
            for idx in range(1, prs.Slides.Count + 1):
                if prs.Slides(idx).SlideID == slide_id:
                    current_index = idx
                    break
            ensure(current_index is not None, "engine_error", "Failed to resolve slide during reordering")
            prs.Slides(current_index).MoveTo(target_position)

    def move_slide(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        from_index = int(params["from_index"])
        to_index = int(params["to_index"])
        self.worker.call(self._move_slide_threadsafe, presentation_id, from_index, to_index)

        return {
            "success": True,
            "presentation_id": presentation_id,
            "from_index": from_index,
            "to_index": to_index,
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }

    def _move_slide_threadsafe(self, session_id: str, from_index: int, to_index: int) -> None:
        prs = self._require_presentation(session_id)
        ensure(1 <= from_index <= prs.Slides.Count, "validation_error", "from_index out of bounds")
        ensure(1 <= to_index <= prs.Slides.Count, "validation_error", "to_index out of bounds")
        prs.Slides(from_index).MoveTo(to_index)

    def set_slide_background(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])

        color_hex = params.get("color_hex")
        image_path = params.get("image_path")
        grad_start = params.get("gradient_start_color_hex")
        grad_end = params.get("gradient_end_color_hex")

        if image_path:
            validate_existing_file(str(image_path), expected_suffixes=(".png", ".jpg", ".jpeg", ".bmp", ".gif"))

        warning = self.worker.call(
            self._set_slide_background_threadsafe,
            presentation_id,
            slide_index,
            color_hex,
            image_path,
            grad_start,
            grad_end,
        )

        payload: dict[str, Any] = {
            "success": True,
            "presentation_id": presentation_id,
            "slide_index": slide_index,
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }
        if warning:
            payload["warning"] = warning
        return payload

    def _set_slide_background_threadsafe(
        self,
        session_id: str,
        slide_index: int,
        color_hex: str | None,
        image_path: str | None,
        grad_start: str | None,
        grad_end: str | None,
    ) -> str:
        slide = self._slide(session_id, slide_index)
        slide.FollowMasterBackground = False

        warning = ""
        fill = slide.Background.Fill

        if color_hex:
            fill.Solid()
            fill.ForeColor.RGB = _hex_to_bgr_int(str(color_hex))

        if image_path:
            fill.UserPicture(str(image_path))

        if grad_start and grad_end:
            try:
                fill.TwoColorGradient(1, 1)
                fill.ForeColor.RGB = _hex_to_bgr_int(str(grad_start))
                fill.BackColor.RGB = _hex_to_bgr_int(str(grad_end))
            except Exception:
                warning = "Two-color gradient fallback may not render identically in all templates."

        return warning

    def get_slide_snapshot(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        width_px = int(params.get("width_px") or 1280)

        snapshot = self.worker.call(self._export_slide_snapshot_threadsafe, presentation_id, slide_index, width_px)

        return {
            "presentation_id": presentation_id,
            "slide_index": slide_index,
            "mime_type": "image/jpeg",
            "snapshot_base64": snapshot,
            "width_px": width_px,
        }

    def _export_slide_snapshot_threadsafe(self, session_id: str, slide_index: int, width_px: int) -> str:
        prs = self._require_presentation(session_id)
        ensure(1 <= slide_index <= prs.Slides.Count, "validation_error", "slide_index out of bounds")
        slide = prs.Slides(slide_index)

        width_points = prs.PageSetup.SlideWidth
        height_points = prs.PageSetup.SlideHeight
        ratio = float(height_points) / float(width_points) if width_points else 0.5625
        height_px = max(1, int(width_px * ratio))

        fd, tmp_path = tempfile.mkstemp(suffix=".jpg", prefix="pptx-snapshot-")
        os.close(fd)
        try:
            slide.Export(tmp_path, "JPG", width_px, height_px)
            encoded = base64.b64encode(Path(tmp_path).read_bytes()).decode("ascii")
            return encoded
        finally:
            Path(tmp_path).unlink(missing_ok=True)

    def get_placeholders(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        placeholders = self.worker.call(self._get_placeholders_threadsafe, presentation_id, slide_index)
        return {
            "presentation_id": presentation_id,
            "slide_index": slide_index,
            "placeholders": placeholders,
        }

    def _get_placeholders_threadsafe(self, session_id: str, slide_index: int) -> list[dict[str, Any]]:
        slide = self._slide(session_id, slide_index)
        results = []
        try:
            phs = slide.Shapes.Placeholders
            for i in range(1, phs.Count + 1):
                results.append(self._placeholder_payload(phs(i)))
        except Exception:
            pass
        return results

    def set_placeholder_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        placeholder_name = str(params["placeholder_name"])

        self.worker.call(self._set_placeholder_text_threadsafe, presentation_id, slide_index, params)

        return {
            "success": True,
            "presentation_id": presentation_id,
            "slide_index": slide_index,
            "placeholder_name": placeholder_name,
            "text_content": str(params.get("text_content", "")),
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }

    def _set_placeholder_text_threadsafe(self, session_id: str, slide_index: int, params: dict[str, Any]) -> None:
        slide = self._slide(session_id, slide_index)
        placeholder = self._find_placeholder_by_name(slide, str(params["placeholder_name"]))

        ensure(bool(placeholder.HasTextFrame), "conflict", "Target placeholder does not support text")

        text_range = placeholder.TextFrame.TextRange
        text_range.Text = str(params.get("text_content", ""))

        if params.get("font_name"):
            text_range.Font.Name = str(params["font_name"])
        if params.get("font_size_pt"):
            text_range.Font.Size = float(params["font_size_pt"])
        if params.get("bold") is not None:
            text_range.Font.Bold = -1 if bool(params["bold"]) else 0
        if params.get("italic") is not None:
            text_range.Font.Italic = -1 if bool(params["italic"]) else 0
        if params.get("underline") is not None:
            text_range.Font.Underline = -1 if bool(params["underline"]) else 0
        if params.get("color_hex"):
            text_range.Font.Color.RGB = _hex_to_bgr_int(str(params["color_hex"]))

        alignment = params.get("alignment")
        if alignment:
            align_map = {
                "left": 1,
                "center": 2,
                "right": 3,
                "justify": 4,
            }
            text_range.ParagraphFormat.Alignment = align_map[str(alignment)]

    def set_placeholder_image(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        image_path = validate_existing_file(
            str(params["image_path"]),
            expected_suffixes=(".png", ".jpg", ".jpeg", ".bmp", ".gif"),
        )

        self.worker.call(
            self._set_placeholder_image_threadsafe,
            presentation_id,
            slide_index,
            str(params["placeholder_name"]),
            str(image_path),
        )

        return {
            "success": True,
            "presentation_id": presentation_id,
            "slide_index": slide_index,
            "placeholder_name": str(params["placeholder_name"]),
            "image_path": str(image_path),
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }

    def _set_placeholder_image_threadsafe(
        self, session_id: str, slide_index: int, placeholder_name: str, image_path: str
    ) -> None:
        slide = self._slide(session_id, slide_index)
        placeholder = self._find_placeholder_by_name(slide, placeholder_name)
        try:
            placeholder.Fill.UserPicture(image_path)
        except Exception:
            slide.Shapes.AddPicture(
                image_path, False, True, placeholder.Left, placeholder.Top, placeholder.Width, placeholder.Height
            )

    def clear_placeholder(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        placeholder_name = str(params["placeholder_name"])

        warning = self.worker.call(
            self._clear_placeholder_threadsafe,
            presentation_id,
            slide_index,
            placeholder_name,
        )

        payload: dict[str, Any] = {
            "success": True,
            "presentation_id": presentation_id,
            "slide_index": slide_index,
            "placeholder_name": placeholder_name,
            "presentation_state": self.get_presentation_state({"presentation_id": presentation_id}),
        }
        if warning:
            payload["warning"] = warning
        return payload

    def _clear_placeholder_threadsafe(self, session_id: str, slide_index: int, placeholder_name: str) -> str:
        slide = self._slide(session_id, slide_index)
        placeholder = self._find_placeholder_by_name(slide, placeholder_name)
        if bool(placeholder.HasTextFrame):
            placeholder.TextFrame.TextRange.Text = ""
            return ""
        return "Non-text placeholder clear is best-effort in COM mode."

    def get_placeholder_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_index = int(params["slide_index"])
        placeholder_name = str(params["placeholder_name"])

        payload = self.worker.call(
            self._get_placeholder_text_threadsafe,
            presentation_id,
            slide_index,
            placeholder_name,
        )

        return {
            "presentation_id": presentation_id,
            "slide_index": slide_index,
            "placeholder_name": placeholder_name,
            **payload,
        }

    def _get_placeholder_text_threadsafe(
        self, session_id: str, slide_index: int, placeholder_name: str
    ) -> dict[str, Any]:
        slide = self._slide(session_id, slide_index)
        placeholder = self._find_placeholder_by_name(slide, placeholder_name)
        ensure(bool(placeholder.HasTextFrame), "conflict", "Target placeholder does not contain text")

        text_range = placeholder.TextFrame.TextRange
        return {
            "raw_text": str(text_range.Text),
            "paragraphs": [
                {
                    "text": str(text_range.Text),
                    "runs": [
                        {
                            "text": str(text_range.Text),
                            "bold": bool(text_range.Font.Bold == -1),
                            "italic": bool(text_range.Font.Italic == -1),
                            "underline": bool(text_range.Font.Underline == -1),
                            "font_name": str(text_range.Font.Name),
                            "font_size_pt": float(text_range.Font.Size),
                        }
                    ],
                }
            ],
        }

    # ------------------------------------------------------------------ #
    # Phase 3–7: Full COM implementations
    # ------------------------------------------------------------------ #

    # --- Shape text / text-box / notes ---

    def set_placeholder_rich_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        ph_name = str(params["placeholder_name"])
        paragraphs = params["paragraphs"]
        self.worker.call(self._set_placeholder_rich_text_ts, pid, si, ph_name, paragraphs)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "placeholder_name": ph_name,
            "paragraph_count": len(paragraphs),
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_placeholder_rich_text_ts(self, sid: str, si: int, ph_name: str, paragraphs: list[dict[str, Any]]) -> None:
        slide = self._slide(sid, si)
        ph = self._find_placeholder_by_name(slide, ph_name)
        ensure(bool(ph.HasTextFrame), "conflict", "Placeholder does not support text")
        self._write_rich_text_com(ph.TextFrame, paragraphs)

    def add_text_box(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_text_box_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_text_box_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        left = _to_points(params["left"])
        top = _to_points(params["top"])
        width = _to_points(params["width"])
        height = _to_points(params["height"])
        shape = slide.Shapes.AddTextbox(1, left, top, width, height)  # 1 = msoTextOrientationHorizontal

        if params.get("paragraphs"):
            self._write_rich_text_com(shape.TextFrame, params["paragraphs"])
        else:
            text_content = str(params.get("text_content", ""))
            shape.TextFrame.TextRange.Text = text_content
            tr = shape.TextFrame.TextRange
            self._apply_com_font(tr.Font, params)
            alignment = params.get("alignment")
            if alignment and alignment in _ALIGN_MAP:
                tr.ParagraphFormat.Alignment = _ALIGN_MAP[alignment]

        return {"shape_name": str(shape.Name), "shape_id": int(shape.Id)}

    def set_slide_notes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        notes_text = str(params.get("notes_text", ""))
        self.worker.call(self._set_slide_notes_ts, pid, si, notes_text)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "notes_length": len(notes_text),
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_slide_notes_ts(self, sid: str, si: int, notes_text: str) -> None:
        slide = self._slide(sid, si)
        notes_page = slide.NotesPage
        phs = notes_page.Shapes.Placeholders
        ensure(phs.Count >= 2, "engine_error", "Notes placeholder not found")
        phs(2).TextFrame.TextRange.Text = notes_text

    def set_shape_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        self.worker.call(self._set_shape_text_ts, pid, si, params)
        self._mark_dirty(pid)
        shape_name = str(params.get("shape_name", params.get("shape_id", "")))
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": shape_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_shape_text_ts(self, sid: str, si: int, params: dict[str, Any]) -> None:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTextFrame), "conflict", "Shape does not support text")

        if params.get("paragraphs"):
            self._write_rich_text_com(shape.TextFrame, params["paragraphs"])
        else:
            shape.TextFrame.TextRange.Text = str(params.get("text_content", ""))

    # --- Read operations ---

    def get_slide_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        text_items = self.worker.call(self._get_slide_text_ts, pid, si)
        total_len = sum(len(item.get("raw_text", "")) for item in text_items)
        return {
            "presentation_id": pid,
            "slide_index": si,
            "total_text_length": total_len,
            "item_count": len(text_items),
            "text_items": text_items,
        }

    def _get_slide_text_ts(self, sid: str, si: int) -> list[dict[str, Any]]:
        slide = self._slide(sid, si)
        return self._extract_text_items_com(slide)

    def get_shape_details(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        details = self.worker.call(self._get_shape_details_ts, pid, si, params)
        return {"presentation_id": pid, "slide_index": si, **details}

    def _get_shape_details_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        return self._detailed_shape_payload_com(shape)

    def get_slide_xml(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        session = self._get_session(pid)
        self.worker.call(self._save_in_place_ts, pid)
        xml_content = self._read_slide_xml(session.working_path, si)
        return {"presentation_id": pid, "slide_index": si, "xml_content": xml_content}

    def _save_in_place_ts(self, sid: str) -> None:
        prs = self._require_presentation(sid)
        prs.Save()

    def _read_slide_xml(self, pptx_path: str, slide_index: int) -> str:
        import xml.dom.minidom

        slide_part = f"ppt/slides/slide{slide_index}.xml"
        with zipfile.ZipFile(pptx_path, "r") as zf:
            ensure(slide_part in zf.namelist(), "not_found", f"Slide XML part not found: {slide_part}")
            raw = zf.read(slide_part)
        dom = xml.dom.minidom.parseString(raw)
        return dom.toprettyxml(indent="  ")

    def get_slide_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        shapes = self.worker.call(self._get_slide_shapes_ts, pid, si)
        return {
            "presentation_id": pid,
            "slide_index": si,
            "shape_count": len(shapes),
            "shapes": shapes,
        }

    def _get_slide_shapes_ts(self, sid: str, si: int) -> list[dict[str, Any]]:
        slide = self._slide(sid, si)
        shapes: list[dict[str, Any]] = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            info: dict[str, Any] = {
                "shape_id": int(shape.Id),
                "name": str(shape.Name),
                "shape_type": str(shape.Type),
                "left": float(shape.Left) / 72.0,
                "top": float(shape.Top) / 72.0,
                "width": float(shape.Width) / 72.0,
                "height": float(shape.Height) / 72.0,
                "is_placeholder": bool(getattr(shape, "PlaceholderFormat", None) is not None),
                "has_text_frame": bool(shape.HasTextFrame),
                "has_table": bool(shape.HasTable),
                "has_chart": bool(getattr(shape, "HasChart", False)),
            }
            if info["is_placeholder"]:
                try:
                    info["placeholder_name"] = int(shape.PlaceholderFormat.Idx)
                except Exception:
                    pass
            if int(shape.Type) == 6:  # msoGroup
                try:
                    info["child_count"] = int(shape.GroupItems.Count)
                except Exception:
                    pass
            shapes.append(info)
        return shapes

    # --- Shape creation ---

    def add_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_shape_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_shape_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape_type_str = str(params["shape_type"]).lower()
        ensure(shape_type_str in _MSO_SHAPE_MAP, "validation_error", f"Unknown shape type: {shape_type_str}")
        mso_type = _MSO_SHAPE_MAP[shape_type_str]

        left = _to_points(params["left"])
        top = _to_points(params["top"])
        width = _to_points(params["width"])
        height = _to_points(params["height"])

        shape = slide.Shapes.AddShape(mso_type, left, top, width, height)

        if params.get("fill_hex"):
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = _hex_to_bgr_int(str(params["fill_hex"]))
        if params.get("line_hex"):
            shape.Line.ForeColor.RGB = _hex_to_bgr_int(str(params["line_hex"]))
        if params.get("text"):
            shape.TextFrame.TextRange.Text = str(params["text"])

        return {"shape_name": str(shape.Name), "shape_id": int(shape.Id)}

    def add_image(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        image_path = validate_existing_file(
            str(params["image_path"]),
            expected_suffixes=(".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".svg"),
        )
        result = self.worker.call(self._add_image_ts, pid, si, str(image_path), params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_image_ts(self, sid: str, si: int, image_path: str, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        left = _to_points(params.get("left", "0in"))
        top = _to_points(params.get("top", "0in"))

        has_w = params.get("width") is not None
        has_h = params.get("height") is not None

        if has_w and has_h:
            width = _to_points(params["width"])
            height = _to_points(params["height"])
            shape = slide.Shapes.AddPicture(image_path, False, True, left, top, width, height)
        else:
            shape = slide.Shapes.AddPicture(image_path, False, True, left, top, -1, -1)
            shape.ScaleWidth(1, True)
            shape.ScaleHeight(1, True)
            if has_w:
                shape.Width = _to_points(params["width"])
            if has_h:
                shape.Height = _to_points(params["height"])

        return {
            "shape_name": str(shape.Name),
            "shape_id": int(shape.Id),
            "left_inches": float(shape.Left) / 72.0,
            "top_inches": float(shape.Top) / 72.0,
            "width_inches": float(shape.Width) / 72.0,
            "height_inches": float(shape.Height) / 72.0,
        }

    def add_line(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_line_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_line_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        x1 = _to_points(params["begin_x"])
        y1 = _to_points(params["begin_y"])
        x2 = _to_points(params["end_x"])
        y2 = _to_points(params["end_y"])

        shape = slide.Shapes.AddLine(x1, y1, x2, y2)

        color_hex = params.get("color_hex", "000000")
        shape.Line.ForeColor.RGB = _hex_to_bgr_int(str(color_hex))
        width_pt = float(params.get("width_pt", 1.0))
        shape.Line.Weight = width_pt

        dash_style = params.get("dash_style")
        if dash_style:
            dash_map = {"solid": 1, "dash": 4, "dot": 3, "dash_dot": 5, "long_dash": 7}
            if dash_style.lower() in dash_map:
                shape.Line.DashStyle = dash_map[dash_style.lower()]

        line_name = str(params.get("line_name", "Connector"))
        try:
            shape.Name = line_name
        except Exception:
            pass

        return {
            "line_name": str(shape.Name),
            "begin": {"x_inches": float(x1) / 72.0, "y_inches": float(y1) / 72.0},
            "end": {"x_inches": float(x2) / 72.0, "y_inches": float(y2) / 72.0},
        }

    def delete_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        deleted_name = self.worker.call(self._delete_shape_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "deleted_shape": deleted_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _delete_shape_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        name = str(shape.Name)
        shape.Delete()
        return name

    # --- Shape properties & cloning ---

    def set_shape_properties(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._set_shape_properties_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_shape_properties_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)

        if params.get("left") is not None:
            shape.Left = _to_points(params["left"])
        if params.get("top") is not None:
            shape.Top = _to_points(params["top"])
        if params.get("width") is not None:
            shape.Width = _to_points(params["width"])
        if params.get("height") is not None:
            shape.Height = _to_points(params["height"])
        if params.get("rotation") is not None:
            shape.Rotation = float(params["rotation"])

        fill_hex = params.get("fill_hex")
        if fill_hex is not None:
            if str(fill_hex).lower() == "none":
                shape.Fill.Visible = 0
            else:
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = _hex_to_bgr_int(str(fill_hex))

        line_hex = params.get("line_hex")
        if line_hex is not None:
            if str(line_hex).lower() == "none":
                shape.Line.Visible = 0
            else:
                shape.Line.Visible = -1
                shape.Line.ForeColor.RGB = _hex_to_bgr_int(str(line_hex))

        if params.get("line_width_pt") is not None:
            shape.Line.Weight = float(params["line_width_pt"])
        if params.get("name"):
            shape.Name = str(params["name"])

        return {
            "shape_name": str(shape.Name),
            "left_inches": float(shape.Left) / 72.0,
            "top_inches": float(shape.Top) / 72.0,
            "width_inches": float(shape.Width) / 72.0,
            "height_inches": float(shape.Height) / 72.0,
            "rotation": float(shape.Rotation),
        }

    def clone_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        target_si = int(params.get("target_slide_index", si))
        result = self.worker.call(self._clone_shape_ts, pid, si, target_si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "source_slide_index": si,
            "target_slide_index": target_si,
            "cloned_shape_name": result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _clone_shape_ts(self, sid: str, si: int, target_si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)

        offset_l = _to_points(params["offset_left"]) if params.get("offset_left") else 18  # 0.25in default
        offset_t = _to_points(params["offset_top"]) if params.get("offset_top") else 18

        if si == target_si:
            dup = shape.Duplicate()(1)
            dup.Left = shape.Left + offset_l
            dup.Top = shape.Top + offset_t
            return str(dup.Name)
        else:
            target_slide = self._slide(sid, target_si)
            shape.Copy()
            pasted = target_slide.Shapes.Paste()(1)
            pasted.Left = shape.Left + offset_l
            pasted.Top = shape.Top + offset_t
            return str(pasted.Name)

    # --- Table operations ---

    def add_table(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_table_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_table_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        rows = int(params["rows"])
        cols = int(params["cols"])
        ensure(rows >= 1 and cols >= 1, "validation_error", "rows and cols must be >= 1")

        left = _to_points(params["left"])
        top = _to_points(params["top"])
        width = _to_points(params["width"])
        height = _to_points(params["height"])

        shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
        table = shape.Table

        data = params.get("data")
        if data:
            for r_idx, row_data in enumerate(data):
                if r_idx >= rows:
                    break
                for c_idx, cell_val in enumerate(row_data):
                    if c_idx >= cols:
                        break
                    table.Cell(r_idx + 1, c_idx + 1).Shape.TextFrame.TextRange.Text = str(cell_val)

        return {"shape_name": str(shape.Name), "shape_id": int(shape.Id), "rows": rows, "cols": cols}

    def get_table(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._get_table_ts, pid, si, params)
        return {"presentation_id": pid, "slide_index": si, **result}

    def _get_table_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTable), "conflict", "Shape is not a table")
        table = shape.Table
        rows = int(table.Rows.Count)
        cols = int(table.Columns.Count)
        cells: list[list[dict[str, Any]]] = []
        for r in range(1, rows + 1):
            row_cells: list[dict[str, Any]] = []
            for c in range(1, cols + 1):
                cell_tf = table.Cell(r, c).Shape.TextFrame.TextRange
                row_cells.append({"text": str(cell_tf.Text), "paragraphs": []})
            cells.append(row_cells)
        return {"shape_name": str(shape.Name), "rows": rows, "cols": cols, "cells": cells}

    def set_table_cell(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        self.worker.call(self._set_table_cell_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": str(params.get("shape_name", params.get("shape_id", ""))),
            "row": int(params["row"]),
            "col": int(params["col"]),
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_table_cell_ts(self, sid: str, si: int, params: dict[str, Any]) -> None:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTable), "conflict", "Shape is not a table")
        table = shape.Table
        row = int(params["row"])
        col = int(params["col"])
        ensure(
            0 <= row < table.Rows.Count and 0 <= col < table.Columns.Count,
            "validation_error",
            "row/col out of bounds",
        )
        cell = table.Cell(row + 1, col + 1)
        cell_tf = cell.Shape.TextFrame.TextRange
        cell_tf.Text = str(params.get("text", ""))
        self._apply_com_font(cell_tf.Font, params)
        if params.get("fill_hex"):
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = _hex_to_bgr_int(str(params["fill_hex"]))

    def set_table_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        rows_written = self.worker.call(self._set_table_data_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": str(params.get("shape_name", params.get("shape_id", ""))),
            "rows_written": rows_written,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_table_data_ts(self, sid: str, si: int, params: dict[str, Any]) -> int:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTable), "conflict", "Shape is not a table")
        table = shape.Table
        data = params["data"]
        rows_written = 0
        for r_idx, row_data in enumerate(data):
            if r_idx >= table.Rows.Count:
                break
            for c_idx, cell_val in enumerate(row_data):
                if c_idx >= table.Columns.Count:
                    break
                cell = table.Cell(r_idx + 1, c_idx + 1)
                if isinstance(cell_val, dict):
                    cell.Shape.TextFrame.TextRange.Text = str(cell_val.get("text", ""))
                    self._apply_com_font(cell.Shape.TextFrame.TextRange.Font, cell_val)
                    if cell_val.get("fill_hex"):
                        cell.Shape.Fill.Solid()
                        cell.Shape.Fill.ForeColor.RGB = _hex_to_bgr_int(str(cell_val["fill_hex"]))
                else:
                    cell.Shape.TextFrame.TextRange.Text = str(cell_val)
            rows_written += 1
        return rows_written

    def set_table_cell_merge(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        self.worker.call(self._set_table_cell_merge_ts, pid, si, params)
        self._mark_dirty(pid)
        sr = int(params["start_row"])
        sc = int(params["start_col"])
        er = int(params["end_row"])
        ec = int(params["end_col"])
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": str(params.get("shape_name", params.get("shape_id", ""))),
            "merged_range": f"({sr},{sc})->({er},{ec})",
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_table_cell_merge_ts(self, sid: str, si: int, params: dict[str, Any]) -> None:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTable), "conflict", "Shape is not a table")
        table = shape.Table
        sr = int(params["start_row"])
        sc = int(params["start_col"])
        er = int(params["end_row"])
        ec = int(params["end_col"])
        ensure(
            0 <= sr <= er < table.Rows.Count and 0 <= sc <= ec < table.Columns.Count,
            "validation_error",
            "Merge range out of bounds",
        )
        table.Cell(sr + 1, sc + 1).Merge(table.Cell(er + 1, ec + 1))

    # --- Z-order, grouping, ungrouping ---

    def set_shape_z_order(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        action = str(params.get("action", "front")).lower()
        self.worker.call(self._set_shape_z_order_ts, pid, si, action, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": str(params.get("shape_name", params.get("shape_id", ""))),
            "action": action,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_shape_z_order_ts(self, sid: str, si: int, action: str, params: dict[str, Any]) -> None:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(action in _ZORDER_MAP, "validation_error", f"Unknown z-order action: {action}")
        shape.ZOrder(_ZORDER_MAP[action])

    def group_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._group_shapes_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _group_shapes_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape_names = params.get("shape_names")
        shape_ids = params.get("shape_ids")
        ensure(shape_names or shape_ids, "validation_error", "Provide shape_names or shape_ids")

        names: list[str] = []
        if shape_names:
            names = [str(n) for n in shape_names]
        elif shape_ids:
            for sid_val in shape_ids:
                shape = self._find_shape_com(slide, shape_id=int(sid_val))
                names.append(str(shape.Name))

        ensure(len(names) >= 2, "validation_error", "At least 2 shapes required for grouping")

        shape_range = slide.Shapes.Range(names)
        group = shape_range.Group()
        group_name = str(params.get("group_name", "Group"))
        try:
            group.Name = group_name
        except Exception:
            group_name = str(group.Name)
        return {"group_name": group_name, "shapes_grouped": len(names)}

    def ungroup_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        ungrouped = self.worker.call(self._ungroup_shapes_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "ungrouped_shapes": ungrouped,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _ungroup_shapes_ts(self, sid: str, si: int, params: dict[str, Any]) -> list[str]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(int(shape.Type) == 6, "conflict", "Shape is not a group")
        shape_range = shape.Ungroup()
        return [str(shape_range(i).Name) for i in range(1, shape_range.Count + 1)]

    def copy_shape_between_decks(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        src_pid = str(params["source_presentation_id"])
        tgt_pid = str(params["target_presentation_id"])
        src_si = int(params["source_slide_index"])
        tgt_si = int(params["target_slide_index"])
        result = self.worker.call(self._copy_shape_between_decks_ts, src_pid, tgt_pid, src_si, tgt_si, params)
        self._mark_dirty(tgt_pid)
        return {
            "success": True,
            "source_presentation_id": src_pid,
            "target_presentation_id": tgt_pid,
            "source_slide_index": src_si,
            "target_slide_index": tgt_si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": tgt_pid}),
        }

    def _copy_shape_between_decks_ts(
        self, src_sid: str, tgt_sid: str, src_si: int, tgt_si: int, params: dict[str, Any]
    ) -> dict[str, Any]:
        src_slide = self._slide(src_sid, src_si)
        tgt_slide = self._slide(tgt_sid, tgt_si)
        shape = self._require_shape_com(src_slide, params)
        src_name = str(shape.Name)

        shape.Copy()
        pasted = tgt_slide.Shapes.Paste()(1)

        if params.get("offset_left"):
            pasted.Left = pasted.Left + _to_points(params["offset_left"])
        if params.get("offset_top"):
            pasted.Top = pasted.Top + _to_points(params["offset_top"])

        return {"source_shape_name": src_name, "new_shape_name": str(pasted.Name)}

    # --- Find/replace ---

    def find_replace_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        find_text = str(params["find_text"])
        replace_text = str(params["replace_text"])
        case_sensitive = bool(params.get("case_sensitive", True))
        slide_indices = params.get("slide_indices")
        result = self.worker.call(
            self._find_replace_text_ts, pid, find_text, replace_text, case_sensitive, slide_indices
        )
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "find_text": find_text,
            "replace_text": replace_text,
            "total_replacements": result["total"],
            "slides_searched": result["slides_searched"],
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _find_replace_text_ts(
        self,
        sid: str,
        find_text: str,
        replace_text: str,
        case_sensitive: bool,
        slide_indices: list[int] | None,
    ) -> dict[str, Any]:
        prs = self._require_presentation(sid)
        total = 0
        slides_searched = 0
        for i in range(1, prs.Slides.Count + 1):
            if slide_indices and i not in slide_indices:
                continue
            slide = prs.Slides(i)
            slides_searched += 1
            total += self._replace_in_shapes_com(slide, find_text, replace_text, case_sensitive)
        return {"total": total, "slides_searched": slides_searched}

    # --- Paragraph spacing & text box properties ---

    def set_paragraph_spacing(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        self.worker.call(self._set_paragraph_spacing_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": str(params.get("shape_name", params.get("shape_id", ""))),
            "paragraph_index": int(params["paragraph_index"]),
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_paragraph_spacing_ts(self, sid: str, si: int, params: dict[str, Any]) -> None:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTextFrame), "conflict", "Shape does not support text")
        para_idx = int(params["paragraph_index"])
        pf = shape.TextFrame.TextRange.Paragraphs(para_idx + 1).ParagraphFormat

        if "line_spacing" in params:
            val = params["line_spacing"]
            if val is not None:
                pf.SpaceWithin = float(val)
        if "space_before" in params:
            val = params["space_before"]
            if val is not None:
                pf.SpaceBefore = float(val)
        if "space_after" in params:
            val = params["space_after"]
            if val is not None:
                pf.SpaceAfter = float(val)

    def set_text_box_properties(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        shape_name = self.worker.call(self._set_text_box_properties_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "shape_name": shape_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_text_box_properties_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTextFrame), "conflict", "Shape does not have a text frame")
        tf = shape.TextFrame

        if params.get("margin_left"):
            tf.MarginLeft = _to_points(params["margin_left"])
        if params.get("margin_right"):
            tf.MarginRight = _to_points(params["margin_right"])
        if params.get("margin_top"):
            tf.MarginTop = _to_points(params["margin_top"])
        if params.get("margin_bottom"):
            tf.MarginBottom = _to_points(params["margin_bottom"])

        if params.get("word_wrap") is not None:
            tf.WordWrap = -1 if bool(params["word_wrap"]) else 0

        auto_size = params.get("auto_size")
        if auto_size is not None:
            auto_map = {"none": 0, "shape_to_fit_text": 1, "text_to_fit_shape": 2}
            if auto_size in auto_map:
                tf.AutoSize = auto_map[auto_size]

        vert = params.get("vertical_alignment")
        if vert is not None:
            vert_map = {"top": 1, "middle": 3, "bottom": 4}
            if vert in vert_map:
                shape.TextFrame2.VerticalAnchor = vert_map[vert]

        return str(shape.Name)

    # --- Table style ---

    def set_table_style(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        shape_name = self.worker.call(self._set_table_style_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "shape_name": shape_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_table_style_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(shape.HasTable), "conflict", "Shape is not a table")
        table = shape.Table

        if params.get("first_row") is not None:
            table.FirstRow = -1 if bool(params["first_row"]) else 0
        if params.get("last_row") is not None:
            table.LastRow = -1 if bool(params["last_row"]) else 0
        if params.get("first_col") is not None:
            table.FirstCol = -1 if bool(params["first_col"]) else 0
        if params.get("last_col") is not None:
            table.LastCol = -1 if bool(params["last_col"]) else 0
        if params.get("banded_rows") is not None:
            table.HorizBanding = -1 if bool(params["banded_rows"]) else 0
        if params.get("banded_cols") is not None:
            table.VertBanding = -1 if bool(params["banded_cols"]) else 0

        return str(shape.Name)

    # --- Gradient fill ---

    def set_shape_fill_gradient(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        shape_name = self.worker.call(self._set_shape_fill_gradient_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "shape_name": shape_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_shape_fill_gradient_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        fill = shape.Fill
        stops = params.get("stops", [])
        angle = params.get("angle")

        if len(stops) == 2:
            fill.TwoColorGradient(1, 1)  # msoGradientHorizontal, variant 1
            fill.ForeColor.RGB = _hex_to_bgr_int(str(stops[0]["color_hex"]))
            fill.BackColor.RGB = _hex_to_bgr_int(str(stops[1]["color_hex"]))
        elif len(stops) > 2:
            fill.TwoColorGradient(1, 1)
            fill.ForeColor.RGB = _hex_to_bgr_int(str(stops[0]["color_hex"]))
            fill.BackColor.RGB = _hex_to_bgr_int(str(stops[-1]["color_hex"]))
            try:
                gs = fill.GradientStops
                while gs.Count > 2:
                    gs.Delete(gs.Count)
                for stop in stops[1:-1]:
                    gs.Insert(
                        _hex_to_bgr_int(str(stop["color_hex"])),
                        float(stop["position"]),
                    )
            except Exception:
                pass  # Older Office may not support GradientStops

        if angle is not None:
            try:
                fill.GradientAngle = float(angle)
            except Exception:
                pass  # Office 2013+ only

        return str(shape.Name)

    # --- Connector ---

    def add_connector(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_connector_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_connector_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        conn_type_str = str(params.get("connector_type", "straight")).lower()
        conn_type = _MSO_CONNECTOR_MAP.get(conn_type_str, 1)

        # Create connector with temporary coordinates
        connector = slide.Shapes.AddConnector(conn_type, 0, 0, 100, 100)

        # Connect begin shape
        begin_name = params.get("begin_shape_name")
        begin_id = params.get("begin_shape_id")
        if begin_name or begin_id:
            begin_shape = self._find_shape_com(
                slide, shape_name=begin_name, shape_id=int(begin_id) if begin_id else None
            )
            site = int(params.get("begin_connection_site", 0)) + 1  # COM is 1-based
            connector.ConnectorFormat.BeginConnect(begin_shape, site)

        # Connect end shape
        end_name = params.get("end_shape_name")
        end_id = params.get("end_shape_id")
        if end_name or end_id:
            end_shape = self._find_shape_com(slide, shape_name=end_name, shape_id=int(end_id) if end_id else None)
            site = int(params.get("end_connection_site", 0)) + 1
            connector.ConnectorFormat.EndConnect(end_shape, site)

        if params.get("color_hex"):
            connector.Line.ForeColor.RGB = _hex_to_bgr_int(str(params["color_hex"]))
        if params.get("width_pt"):
            connector.Line.Weight = float(params["width_pt"])

        connector.RerouteConnections()

        return {"shape_name": str(connector.Name), "shape_id": int(connector.Id)}

    # --- Charts ---

    def add_chart(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_chart_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_chart_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        chart_type_str = str(params["chart_type"]).lower()
        ensure(chart_type_str in _XL_CHART_TYPE_MAP, "validation_error", f"Unknown chart type: {chart_type_str}")
        xl_type = _XL_CHART_TYPE_MAP[chart_type_str]

        left = _to_points(params.get("left", "1in"))
        top = _to_points(params.get("top", "1in"))
        width = _to_points(params.get("width", "8in"))
        height = _to_points(params.get("height", "5in"))

        # Try AddChart2 (Office 2013+), fallback to AddChart
        try:
            chart_shape = slide.Shapes.AddChart2(-1, xl_type, left, top, width, height)
        except Exception:
            chart_shape = slide.Shapes.AddChart(xl_type, left, top, width, height)

        chart = chart_shape.Chart
        series_data = params.get("series", [])
        categories = params.get("categories", [])
        is_xy = chart_type_str.startswith("xy_")
        is_bubble = chart_type_str.startswith("bubble")

        if series_data:
            chart.ChartData.Activate()
            wb = chart.ChartData.Workbook
            try:
                wb.Application.Visible = False
            except Exception:
                pass
            try:
                ws = wb.Worksheets(1)
                ws.Cells.Clear()

                if is_xy or is_bubble:
                    for s_idx, s in enumerate(series_data):
                        col_offset = s_idx * (3 if is_bubble else 2)
                        s_name = str(s.get("name", f"Series {s_idx + 1}"))
                        ws.Cells(1, col_offset + 1).Value = f"{s_name} X"
                        ws.Cells(1, col_offset + 2).Value = f"{s_name} Y"
                        if is_bubble:
                            ws.Cells(1, col_offset + 3).Value = f"{s_name} Size"
                        for dp_idx, dp in enumerate(s.get("data_points", [])):
                            ws.Cells(dp_idx + 2, col_offset + 1).Value = dp.get("x", 0)
                            ws.Cells(dp_idx + 2, col_offset + 2).Value = dp.get("y", 0)
                            if is_bubble:
                                ws.Cells(dp_idx + 2, col_offset + 3).Value = dp.get("size", 10)
                    max_pts = max((len(s.get("data_points", [])) for s in series_data), default=0)
                    total_cols = len(series_data) * (3 if is_bubble else 2)
                    data_range = ws.Range(ws.Cells(1, 1), ws.Cells(max_pts + 1, total_cols))
                    chart.SetSourceData(data_range)
                else:
                    # Category charts: col A = categories, cols B+ = series values
                    ws.Cells(1, 1).Value = "Category"
                    for s_idx, s in enumerate(series_data):
                        ws.Cells(1, s_idx + 2).Value = str(s.get("name", f"Series {s_idx + 1}"))
                    for c_idx, cat in enumerate(categories):
                        ws.Cells(c_idx + 2, 1).Value = str(cat)
                    for s_idx, s in enumerate(series_data):
                        for v_idx, val in enumerate(s.get("values", [])):
                            ws.Cells(v_idx + 2, s_idx + 2).Value = val
                    num_rows = max(len(categories), max((len(s.get("values", [])) for s in series_data), default=0))
                    data_range = ws.Range(ws.Cells(1, 1), ws.Cells(num_rows + 1, len(series_data) + 1))
                    chart.SetSourceData(data_range)
            finally:
                try:
                    wb.Close(False)
                except Exception:
                    pass

        # Apply optional chart formatting
        if params.get("title"):
            chart.HasTitle = True
            chart.ChartTitle.Text = str(params["title"])
        if params.get("has_legend") is not None:
            chart.HasLegend = bool(params["has_legend"])
        if params.get("legend_position") and chart.HasLegend:
            pos = str(params["legend_position"]).lower()
            if pos in _LEGEND_POS_MAP:
                chart.Legend.Position = _LEGEND_POS_MAP[pos]
        if params.get("has_data_labels") is not None:
            for s_idx in range(1, chart.SeriesCollection().Count + 1):
                chart.SeriesCollection(s_idx).HasDataLabels = bool(params["has_data_labels"])
        if params.get("chart_style"):
            try:
                chart.ChartStyle = int(params["chart_style"])
            except Exception:
                pass

        return {
            "shape_name": str(chart_shape.Name),
            "shape_id": int(chart_shape.Id),
            "chart_type": chart_type_str,
        }

    def get_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._get_chart_data_ts, pid, si, params)
        return {"presentation_id": pid, "slide_index": si, **result}

    def _get_chart_data_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(getattr(shape, "HasChart", False)), "conflict", "Shape is not a chart")
        chart = shape.Chart

        chart_type_int = int(chart.ChartType)
        chart_type_str = _XL_CHART_TYPE_REVERSE.get(chart_type_int, f"unknown_{chart_type_int}")

        result: dict[str, Any] = {
            "shape_name": str(shape.Name),
            "chart_type": chart_type_str,
            "categories": [],
            "series": [],
            "has_legend": bool(chart.HasLegend),
            "has_title": bool(chart.HasTitle),
            "has_data_labels": False,
        }

        if chart.HasTitle:
            result["title"] = str(chart.ChartTitle.Text)

        try:
            chart.ChartData.Activate()
            wb = chart.ChartData.Workbook
            try:
                wb.Application.Visible = False
            except Exception:
                pass
            try:
                ws = wb.Worksheets(1)
                used = ws.UsedRange
                rows = int(used.Rows.Count)
                cols = int(used.Columns.Count)

                # Read categories from column A (rows 2+)
                cats = []
                for r in range(2, rows + 1):
                    val = ws.Cells(r, 1).Value
                    cats.append(str(val) if val is not None else "")
                result["categories"] = cats

                # Read series from columns B+
                series_list = []
                for c in range(2, cols + 1):
                    s_name = ws.Cells(1, c).Value
                    vals = []
                    for r in range(2, rows + 1):
                        v = ws.Cells(r, c).Value
                        vals.append(float(v) if v is not None else None)
                    series_list.append({"name": str(s_name) if s_name else f"Series {c - 1}", "values": vals})
                result["series"] = series_list
            finally:
                try:
                    wb.Close(False)
                except Exception:
                    pass
        except Exception:
            pass

        # Check data labels on first series
        try:
            if chart.SeriesCollection().Count > 0:
                result["has_data_labels"] = bool(chart.SeriesCollection(1).HasDataLabels)
        except Exception:
            pass

        return result

    def update_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._update_chart_data_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _update_chart_data_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(getattr(shape, "HasChart", False)), "conflict", "Shape is not a chart")
        chart = shape.Chart
        chart_type_int = int(chart.ChartType)
        chart_type_str = _XL_CHART_TYPE_REVERSE.get(chart_type_int, f"unknown_{chart_type_int}")

        series_data = params.get("series", [])
        categories = params.get("categories")

        if series_data or categories:
            chart.ChartData.Activate()
            wb = chart.ChartData.Workbook
            try:
                wb.Application.Visible = False
            except Exception:
                pass
            try:
                ws = wb.Worksheets(1)

                if categories:
                    for c_idx, cat in enumerate(categories):
                        ws.Cells(c_idx + 2, 1).Value = str(cat)

                for s_idx, s in enumerate(series_data):
                    col = s_idx + 2
                    if s.get("name"):
                        ws.Cells(1, col).Value = str(s["name"])
                    for v_idx, val in enumerate(s.get("values", [])):
                        ws.Cells(v_idx + 2, col).Value = val

                num_cats = len(categories) if categories else 0
                num_vals = max((len(s.get("values", [])) for s in series_data), default=0) if series_data else 0
                num_rows = max(num_cats, num_vals)
                if num_rows > 0:
                    num_cols = len(series_data) + 1 if series_data else int(ws.UsedRange.Columns.Count)
                    data_range = ws.Range(ws.Cells(1, 1), ws.Cells(num_rows + 1, num_cols))
                    chart.SetSourceData(data_range)
            finally:
                try:
                    wb.Close(False)
                except Exception:
                    pass

        return {"shape_name": str(shape.Name), "chart_type": chart_type_str}

    def set_chart_style(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        shape_name = self.worker.call(self._set_chart_style_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": shape_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_chart_style_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(getattr(shape, "HasChart", False)), "conflict", "Shape is not a chart")
        chart = shape.Chart

        if params.get("has_legend") is not None:
            chart.HasLegend = bool(params["has_legend"])
        if params.get("legend_position") and chart.HasLegend:
            pos = str(params["legend_position"]).lower()
            if pos in _LEGEND_POS_MAP:
                chart.Legend.Position = _LEGEND_POS_MAP[pos]
        if params.get("has_data_labels") is not None:
            for s_idx in range(1, chart.SeriesCollection().Count + 1):
                chart.SeriesCollection(s_idx).HasDataLabels = bool(params["has_data_labels"])
        if params.get("chart_style"):
            try:
                chart.ChartStyle = int(params["chart_style"])
            except Exception:
                pass
        if "title" in params:
            title_val = params["title"]
            if title_val is None or title_val == "":
                chart.HasTitle = False
            else:
                chart.HasTitle = True
                chart.ChartTitle.Text = str(title_val)

        return str(shape.Name)

    def shutdown(self) -> None:  # pragma: no cover - Windows only
        for session_id in list(self.sessions.keys()):
            try:
                self.worker.call(self._close_presentation_threadsafe, session_id)
            except Exception:
                pass
            try:
                Path(self.sessions[session_id].working_path).unlink(missing_ok=True)
            except Exception:
                pass
        self.sessions.clear()

        try:
            self.worker.call(self._shutdown_app)
        except Exception:
            pass
        self.worker.shutdown()

    def _shutdown_app(self) -> None:
        if self._app is not None:
            self._app.Quit()

    def _get_session(self, presentation_id: str) -> EngineSession:
        session = self.sessions.get(presentation_id)
        if session is None:
            raise BridgeError("not_found", "Presentation session not found.", {"presentation_id": presentation_id})
        return session

    def _require_presentation(self, session_id: str):
        prs = self._presentations.get(session_id)
        if prs is None:
            raise BridgeError("not_found", "Presentation COM object not found.", {"presentation_id": session_id})
        return prs

    def _slide_count(self, session_id: str) -> int:
        prs = self._require_presentation(session_id)
        return int(prs.Slides.Count)

    def _slide(self, session_id: str, slide_index: int):
        prs = self._require_presentation(session_id)
        ensure(1 <= slide_index <= prs.Slides.Count, "validation_error", "slide_index out of bounds")
        return prs.Slides(slide_index)

    def _find_layout_by_name(self, session_id: str, layout_name: str):
        prs = self._require_presentation(session_id)
        for m in range(1, prs.SlideMasters.Count + 1):
            master = prs.SlideMasters(m)
            for i in range(1, master.CustomLayouts.Count + 1):
                layout = master.CustomLayouts(i)
                if layout.Name == layout_name:
                    return layout

        available = []
        for m in range(1, prs.SlideMasters.Count + 1):
            master = prs.SlideMasters(m)
            for i in range(1, master.CustomLayouts.Count + 1):
                available.append(master.CustomLayouts(i).Name)

        raise BridgeError(
            "validation_error",
            f"Layout '{layout_name}' not found.",
            {"layout_name": layout_name, "available_layouts": available},
        )

    def _find_placeholder_by_name(self, slide, placeholder_name: str):
        phs = slide.Shapes.Placeholders
        for i in range(1, phs.Count + 1):
            ph = phs(i)
            if ph.Name == placeholder_name:
                return ph

        available = [phs(i).Name for i in range(1, phs.Count + 1)]
        raise BridgeError(
            "not_found",
            f"Placeholder '{placeholder_name}' not found.",
            {"placeholder_name": placeholder_name, "available_placeholders": available},
        )

    def _placeholder_payload(self, placeholder) -> dict[str, Any]:
        return {
            "name": str(placeholder.Name),
            "idx": int(placeholder.PlaceholderFormat.Idx),
            "type": str(placeholder.PlaceholderFormat.Type),
            "left_inches": float(placeholder.Left) / 72.0,
            "top_inches": float(placeholder.Top) / 72.0,
            "width_inches": float(placeholder.Width) / 72.0,
            "height_inches": float(placeholder.Height) / 72.0,
        }

    def _shape_payload(self, shape) -> dict[str, Any]:
        payload = {
            "shape_id": int(shape.Id),
            "name": str(shape.Name),
            "type": str(shape.Type),
            "left_inches": float(shape.Left) / 72.0,
            "top_inches": float(shape.Top) / 72.0,
            "width_inches": float(shape.Width) / 72.0,
            "height_inches": float(shape.Height) / 72.0,
            "is_placeholder": bool(getattr(shape, "PlaceholderFormat", None) is not None),
        }
        try:
            if shape.HasTextFrame:
                payload["text"] = str(shape.TextFrame.TextRange.Text)
        except Exception:
            pass
        return payload

    def _slide_title(self, slide) -> str:
        try:
            if slide.Shapes.Title is not None:
                text = str(slide.Shapes.Title.TextFrame.TextRange.Text)
                if text.strip():
                    return text.strip()
        except Exception:
            pass

        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                if shape.HasTextFrame:
                    text = str(shape.TextFrame.TextRange.Text)
                    if text.strip():
                        return text.strip().splitlines()[0]
            except Exception:
                continue
        return ""

    def _notes_text(self, slide) -> str:
        try:
            notes_page = slide.NotesPage
            phs = notes_page.Shapes.Placeholders
            if phs.Count >= 2:
                return str(phs(2).TextFrame.TextRange.Text)
        except Exception:
            pass
        return ""

    def _mark_dirty(self, presentation_id: str) -> None:
        session = self.sessions.get(presentation_id)
        if session:
            session.dirty = True

    # --- Shape lookup helpers ---

    def _find_shape_com(self, slide, *, shape_name: str | None = None, shape_id: int | None = None):
        """Find shape by name or id, searching top-level and inside groups."""
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            if shape_name and str(shape.Name) == shape_name:
                return shape
            if shape_id is not None and int(shape.Id) == shape_id:
                return shape
            # Search inside groups
            if int(shape.Type) == 6:  # msoGroup
                try:
                    for j in range(1, shape.GroupItems.Count + 1):
                        child = shape.GroupItems(j)
                        if shape_name and str(child.Name) == shape_name:
                            return child
                        if shape_id is not None and int(child.Id) == shape_id:
                            return child
                except Exception:
                    pass

        available = []
        for i in range(1, slide.Shapes.Count + 1):
            s = slide.Shapes(i)
            available.append(f"{s.Name} (id={s.Id})")
        raise BridgeError(
            "not_found",
            f"Shape not found (name={shape_name}, id={shape_id}).",
            {"available_shapes": available},
        )

    def _require_shape_com(self, slide, params: dict[str, Any]):
        """Extract shape_name/shape_id from params and call _find_shape_com."""
        name = params.get("shape_name")
        sid = params.get("shape_id")
        ensure(name or sid is not None, "validation_error", "Provide shape_name or shape_id")
        return self._find_shape_com(
            slide, shape_name=str(name) if name else None, shape_id=int(sid) if sid is not None else None
        )

    def _apply_com_font(self, font, data: dict[str, Any]) -> None:
        """Apply font formatting from a dict to a COM Font object."""
        if data.get("font_name"):
            font.Name = str(data["font_name"])
        if data.get("font_size_pt"):
            font.Size = float(data["font_size_pt"])
        if data.get("bold") is not None:
            font.Bold = -1 if bool(data["bold"]) else 0
        if data.get("italic") is not None:
            font.Italic = -1 if bool(data["italic"]) else 0
        if data.get("underline") is not None:
            font.Underline = -1 if bool(data["underline"]) else 0
        if data.get("color_hex"):
            font.Color.RGB = _hex_to_bgr_int(str(data["color_hex"]))

    def _write_rich_text_com(self, text_frame, paragraphs_data: list[dict[str, Any]]) -> None:
        """Write multi-paragraph rich text to a COM TextFrame."""
        # Build combined text with \r as paragraph separator
        para_texts: list[str] = []
        for p in paragraphs_data:
            runs = p.get("runs", [])
            if runs:
                para_texts.append("".join(str(r.get("text", "")) for r in runs))
            else:
                para_texts.append(str(p.get("text", "")))

        combined = "\r".join(para_texts)
        text_frame.TextRange.Text = combined

        # Apply formatting per paragraph
        for p_idx, p_data in enumerate(paragraphs_data):
            para = text_frame.TextRange.Paragraphs(p_idx + 1)

            alignment = p_data.get("alignment")
            if alignment and alignment in _ALIGN_MAP:
                para.ParagraphFormat.Alignment = _ALIGN_MAP[alignment]

            level = p_data.get("level")
            if level is not None:
                para.IndentLevel = int(level) + 1  # COM indent is 1-based

            runs = p_data.get("runs", [])
            if runs:
                # Apply per-run formatting using Characters(start, length)
                char_offset = 1  # COM Characters is 1-based
                for run_data in runs:
                    run_text = str(run_data.get("text", ""))
                    length = len(run_text)
                    if length > 0:
                        char_range = para.Characters(char_offset, length)
                        self._apply_com_font(char_range.Font, run_data)
                    char_offset += length
            else:
                # Apply paragraph-level font formatting if present
                self._apply_com_font(para.Font, p_data)

    def _extract_text_items_com(self, slide) -> list[dict[str, Any]]:
        """Extract text content from all shapes on a slide."""
        items: list[dict[str, Any]] = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            self._extract_shape_text_com(shape, items)
        return items

    def _extract_shape_text_com(self, shape, items: list[dict[str, Any]], parent_group: str | None = None) -> None:
        """Extract text from a single shape, recursing into groups and tables."""
        base: dict[str, Any] = {
            "shape_name": str(shape.Name),
            "shape_id": int(shape.Id),
            "shape_type": str(shape.Type),
            "is_placeholder": bool(getattr(shape, "PlaceholderFormat", None) is not None),
            "left_inches": float(shape.Left) / 72.0,
            "top_inches": float(shape.Top) / 72.0,
            "width_inches": float(shape.Width) / 72.0,
            "height_inches": float(shape.Height) / 72.0,
        }
        if parent_group:
            base["parent_group"] = parent_group

        if bool(shape.HasTable):
            table = shape.Table
            rows = int(table.Rows.Count)
            cols = int(table.Columns.Count)
            raw_parts: list[str] = []
            table_cells: list[dict[str, Any]] = []
            for r in range(1, rows + 1):
                for c in range(1, cols + 1):
                    cell_text = str(table.Cell(r, c).Shape.TextFrame.TextRange.Text)
                    raw_parts.append(cell_text)
                    table_cells.append({"row": r - 1, "col": c - 1, "text": cell_text, "paragraphs": []})
            base["content_type"] = "table"
            base["raw_text"] = "\n".join(raw_parts)
            base["rows"] = rows
            base["cols"] = cols
            base["table_cells"] = table_cells
            base["paragraphs"] = []
            items.append(base)
        elif int(shape.Type) == 6:  # msoGroup
            base["content_type"] = "group"
            base["raw_text"] = ""
            base["paragraphs"] = []
            items.append(base)
            for j in range(1, shape.GroupItems.Count + 1):
                self._extract_shape_text_com(shape.GroupItems(j), items, parent_group=str(shape.Name))
        elif bool(shape.HasTextFrame):
            raw_text = str(shape.TextFrame.TextRange.Text)
            base["content_type"] = "text"
            base["raw_text"] = raw_text
            base["paragraphs"] = []
            items.append(base)
        elif int(shape.Type) == 13:  # msoPicture
            base["content_type"] = "picture"
            base["raw_text"] = ""
            base["paragraphs"] = []
            items.append(base)
        else:
            base["content_type"] = "other"
            base["raw_text"] = ""
            base["paragraphs"] = []
            items.append(base)

    def _detailed_shape_payload_com(self, shape) -> dict[str, Any]:
        """Build a detailed dict describing a single COM shape."""
        payload: dict[str, Any] = {
            "shape_id": int(shape.Id),
            "name": str(shape.Name),
            "shape_type": str(shape.Type),
            "left_inches": float(shape.Left) / 72.0,
            "top_inches": float(shape.Top) / 72.0,
            "width_inches": float(shape.Width) / 72.0,
            "height_inches": float(shape.Height) / 72.0,
            "rotation": float(shape.Rotation),
            "is_placeholder": bool(getattr(shape, "PlaceholderFormat", None) is not None),
        }

        has_text = bool(shape.HasTextFrame)
        payload["has_text"] = has_text
        if has_text:
            payload["text"] = str(shape.TextFrame.TextRange.Text)
            payload["paragraphs"] = []

        has_table = bool(shape.HasTable)
        payload["has_table"] = has_table
        if has_table:
            table = shape.Table
            rows = int(table.Rows.Count)
            cols = int(table.Columns.Count)
            payload["table_rows"] = rows
            payload["table_cols"] = cols
            cells: list[list[dict[str, Any]]] = []
            for r in range(1, rows + 1):
                row_cells: list[dict[str, Any]] = []
                for c in range(1, cols + 1):
                    cell_text = str(table.Cell(r, c).Shape.TextFrame.TextRange.Text)
                    row_cells.append({"text": cell_text, "paragraphs": []})
                cells.append(row_cells)
            payload["table_cells"] = cells

        is_picture = int(shape.Type) == 13
        payload["is_picture"] = is_picture
        if is_picture:
            payload["image_content_type"] = "image/unknown"

        is_group = int(shape.Type) == 6
        payload["is_group"] = is_group
        if is_group:
            count = int(shape.GroupItems.Count)
            payload["child_count"] = count
            children = []
            for j in range(1, count + 1):
                children.append(self._shape_payload(shape.GroupItems(j)))
            payload["children"] = children

        return payload

    def _replace_in_shapes_com(self, slide, find_text: str, replace_text: str, case_sensitive: bool) -> int:
        """Walk all shapes on a slide doing find/replace. Returns total replacement count."""
        flags = 0 if case_sensitive else re.IGNORECASE
        total = 0
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            total += self._replace_in_shape_com(shape, find_text, replace_text, flags)
        return total

    def _replace_in_shape_com(self, shape, find_text: str, replace_text: str, flags: int) -> int:
        """Replace text within a single shape (text, table, or group)."""
        total = 0
        if bool(shape.HasTextFrame):
            tr = shape.TextFrame.TextRange
            original = str(tr.Text)
            new_text, count = re.subn(re.escape(find_text), replace_text, original, flags=flags)
            if count > 0:
                tr.Text = new_text
                total += count

        if bool(shape.HasTable):
            table = shape.Table
            for r in range(1, table.Rows.Count + 1):
                for c in range(1, table.Columns.Count + 1):
                    cell_tr = table.Cell(r, c).Shape.TextFrame.TextRange
                    original = str(cell_tr.Text)
                    new_text, count = re.subn(re.escape(find_text), replace_text, original, flags=flags)
                    if count > 0:
                        cell_tr.Text = new_text
                        total += count

        if int(shape.Type) == 6:  # msoGroup
            for j in range(1, shape.GroupItems.Count + 1):
                total += self._replace_in_shape_com(shape.GroupItems(j), find_text, replace_text, flags)

        return total
