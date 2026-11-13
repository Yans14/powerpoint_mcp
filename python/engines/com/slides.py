"""Slide-level operations: add, duplicate, delete, reorder, background, notes, text search."""

from __future__ import annotations

import zipfile
from typing import Any

from engines.com.common import (
    _hex_to_bgr_int,
)
from errors import ensure
from utils.paths import validate_existing_file


class COMSlidesMixin:
    """Slide-level operations: add, duplicate, delete, reorder, background, notes, text search."""

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
