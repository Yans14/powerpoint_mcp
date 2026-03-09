from __future__ import annotations

import base64
import os
import shutil
import tempfile
import uuid
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
    # Phase 3 stubs – COM implementation pending
    # ------------------------------------------------------------------ #

    def _com_stub(self, method_name: str, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        raise BridgeError(
            code="engine_error",
            message=f"COM engine does not yet implement '{method_name}'. Use OOXML mode.",
            details={"method": method_name},
        )

    def set_placeholder_rich_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_placeholder_rich_text", params)

    def add_text_box(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("add_text_box", params)

    def get_slide_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("get_slide_text", params)

    def get_shape_details(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("get_shape_details", params)

    def add_table(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("add_table", params)

    def get_table(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("get_table", params)

    def set_table_cell(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_table_cell", params)

    def set_table_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_table_data", params)

    def add_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("add_shape", params)

    def delete_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("delete_shape", params)

    def set_slide_notes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_slide_notes", params)

    def set_shape_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_shape_text", params)

    def get_slide_xml(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("get_slide_xml", params)

    # --- Phase 4 stubs ---

    def set_shape_properties(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_shape_properties", params)

    def clone_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("clone_shape", params)

    def group_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("group_shapes", params)

    def ungroup_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("ungroup_shapes", params)

    def set_shape_z_order(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_shape_z_order", params)

    def add_image(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("add_image", params)

    def add_line(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("add_line", params)

    def find_replace_text(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("find_replace_text", params)

    # --- Phase 5 stubs ---

    def add_chart(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("add_chart", params)

    def get_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("get_chart_data", params)

    def update_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("update_chart_data", params)

    def set_chart_style(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_chart_style", params)

    # --- Phase 6 stubs ---

    def copy_shape_between_decks(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("copy_shape_between_decks", params)

    def get_slide_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("get_slide_shapes", params)

    def set_table_cell_merge(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_table_cell_merge", params)

    # --- Phase 7 stubs ---

    def set_paragraph_spacing(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_paragraph_spacing", params)

    def set_text_box_properties(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_text_box_properties", params)

    def set_table_style(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_table_style", params)

    def set_shape_fill_gradient(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("set_shape_fill_gradient", params)

    def add_connector(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        return self._com_stub("add_connector", params)

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
