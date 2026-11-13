"""Placeholders and text formatting."""

from __future__ import annotations

from typing import Any

from engines.com.common import (
    _ALIGN_MAP,
    _hex_to_bgr_int,
    _to_points,
)
from errors import ensure
from utils.paths import validate_existing_file


class COMTextMixin:
    """Placeholders and text formatting."""

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
