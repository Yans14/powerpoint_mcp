"""Placeholders and text formatting."""

from __future__ import annotations

from typing import Any

from pptx.dml.color import RGBColor
from pptx.util import Pt

from errors import ensure
from utils.colors import normalize_color
from utils.paths import validate_existing_file
from utils.units import to_emu


class OOXMLTextMixin:
    """Placeholders and text formatting."""

    def get_placeholders(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide = self._get_slide_obj(session, int(params["slide_index"]))

        placeholders = [self._placeholder_payload(placeholder) for placeholder in slide.placeholders]
        return {
            "presentation_id": session.id,
            "slide_index": int(params["slide_index"]),
            "placeholders": placeholders,
        }

    def set_placeholder_text(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)
        placeholder = self._find_placeholder(slide, str(params["placeholder_name"]))

        ensure(placeholder.has_text_frame, "conflict", "Target placeholder does not support text.")

        text_frame = placeholder.text_frame
        text_frame.clear()
        paragraph = text_frame.paragraphs[0]
        paragraph.text = str(params.get("text_content", ""))

        alignment = params.get("alignment")
        if alignment and alignment in self._ALIGN_MAP:
            paragraph.alignment = self._ALIGN_MAP[alignment]

        if paragraph.runs:
            run = paragraph.runs[0]
            if params.get("font_name"):
                run.font.name = str(params["font_name"])
            if params.get("font_size_pt"):
                run.font.size = Pt(float(params["font_size_pt"]))
            if params.get("bold") is not None:
                run.font.bold = bool(params["bold"])
            if params.get("italic") is not None:
                run.font.italic = bool(params["italic"])
            if params.get("underline") is not None:
                run.font.underline = bool(params["underline"])
            if params.get("color_hex"):
                run.font.color.rgb = RGBColor.from_string(normalize_color(str(params["color_hex"])))

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "placeholder_name": str(params["placeholder_name"]),
            "text_content": str(params.get("text_content", "")),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_placeholder_image(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)
        placeholder = self._find_placeholder(slide, str(params["placeholder_name"]))

        image_path = validate_existing_file(
            str(params["image_path"]),
            expected_suffixes=(".png", ".jpg", ".jpeg", ".bmp", ".gif"),
        )

        if hasattr(placeholder, "insert_picture"):
            placeholder.insert_picture(str(image_path))
        else:
            slide.shapes.add_picture(
                str(image_path), placeholder.left, placeholder.top, placeholder.width, placeholder.height
            )

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "placeholder_name": str(params["placeholder_name"]),
            "image_path": str(image_path),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def clear_placeholder(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)
        placeholder_name = str(params["placeholder_name"])
        placeholder = self._find_placeholder(slide, placeholder_name)

        warning = ""
        if placeholder.has_text_frame:
            placeholder.text_frame.clear()
        else:
            warning = "Non-text placeholder reset is limited in OOXML mode."

        self._mark_modified(session)

        payload: dict[str, Any] = {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "placeholder_name": placeholder_name,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
        if warning:
            payload["warning"] = warning
        return payload

    def get_placeholder_text(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide = self._get_slide_obj(session, int(params["slide_index"]))
        placeholder = self._find_placeholder(slide, str(params["placeholder_name"]))

        ensure(placeholder.has_text_frame, "conflict", "Target placeholder does not contain text.")

        paragraphs = []
        for paragraph in placeholder.text_frame.paragraphs:
            runs = []
            for run in paragraph.runs:
                color_hex = None
                try:
                    if run.font.color and run.font.color.rgb:
                        color_hex = str(run.font.color.rgb)
                except Exception:
                    color_hex = None

                runs.append(
                    {
                        "text": run.text,
                        "bold": bool(run.font.bold) if run.font.bold is not None else None,
                        "italic": bool(run.font.italic) if run.font.italic is not None else None,
                        "underline": bool(run.font.underline) if run.font.underline is not None else None,
                        "font_name": run.font.name,
                        "font_size_pt": float(run.font.size.pt) if run.font.size else None,
                        "color_hex": color_hex,
                    }
                )
            paragraphs.append({"text": paragraph.text, "runs": runs})

        return {
            "presentation_id": session.id,
            "slide_index": int(params["slide_index"]),
            "placeholder_name": str(params["placeholder_name"]),
            "paragraphs": paragraphs,
            "raw_text": placeholder.text,
        }

    def set_placeholder_rich_text(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)
        placeholder = self._find_placeholder(slide, str(params["placeholder_name"]))

        ensure(placeholder.has_text_frame, "conflict", "Target placeholder does not support text.")

        paragraphs_data = params.get("paragraphs", [])
        ensure(len(paragraphs_data) > 0, "validation_error", "At least one paragraph is required.")

        self._write_paragraphs(placeholder.text_frame, paragraphs_data)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "placeholder_name": str(params["placeholder_name"]),
            "paragraph_count": len(paragraphs_data),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def add_text_box(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        left = to_emu(params["left"])
        top = to_emu(params["top"])
        width = to_emu(params["width"])
        height = to_emu(params["height"])

        from pptx.util import Emu as EmuVal

        text_box = slide.shapes.add_textbox(EmuVal(left), EmuVal(top), EmuVal(width), EmuVal(height))
        text_frame = text_box.text_frame

        # Support rich paragraphs or simple text_content
        paragraphs_data = params.get("paragraphs")
        if paragraphs_data:
            self._write_paragraphs(text_frame, paragraphs_data)
        else:
            text_content = str(params.get("text_content", ""))
            paragraph = text_frame.paragraphs[0]
            paragraph.text = text_content

            alignment = params.get("alignment")
            if alignment and alignment in self._ALIGN_MAP:
                paragraph.alignment = self._ALIGN_MAP[alignment]

            if paragraph.runs:
                run = paragraph.runs[0]
                self._apply_run_formatting(run, params)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": text_box.name,
            "shape_id": int(text_box.shape_id),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_shape_text(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)
        ensure(
            hasattr(shape, "has_text_frame") and shape.has_text_frame,
            "conflict",
            "Shape does not have a text frame.",
        )

        text_frame = shape.text_frame
        paragraphs_data = params.get("paragraphs")

        if paragraphs_data:
            self._write_paragraphs(text_frame, paragraphs_data)
        else:
            text_content = str(params.get("text_content", ""))
            text_frame.paragraphs[0].text = text_content

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_paragraph_spacing(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)
        ensure(shape.has_text_frame, "conflict", "Shape does not contain a text frame.")

        paragraph_index = int(params["paragraph_index"])
        ensure(
            0 <= paragraph_index < len(shape.text_frame.paragraphs),
            "validation_error",
            f"Paragraph index out of bounds. Max: {len(shape.text_frame.paragraphs) - 1}",
        )
        paragraph = shape.text_frame.paragraphs[paragraph_index]

        if "line_spacing" in params:
            val = params["line_spacing"]
            paragraph.line_spacing = Pt(float(val)) if val is not None else None

        if "space_before" in params:
            val = params["space_before"]
            paragraph.space_before = Pt(float(val)) if val is not None else None

        if "space_after" in params:
            val = params["space_after"]
            paragraph.space_after = Pt(float(val)) if val is not None else None

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "paragraph_index": paragraph_index,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_text_box_properties(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)
        ensure(shape.has_text_frame, "conflict", "Shape does not contain a text frame.")

        tf = shape.text_frame

        if "margin_left" in params:
            tf.margin_left = to_emu(params["margin_left"]) if params["margin_left"] else 0
        if "margin_right" in params:
            tf.margin_right = to_emu(params["margin_right"]) if params["margin_right"] else 0
        if "margin_top" in params:
            tf.margin_top = to_emu(params["margin_top"]) if params["margin_top"] else 0
        if "margin_bottom" in params:
            tf.margin_bottom = to_emu(params["margin_bottom"]) if params["margin_bottom"] else 0

        if "word_wrap" in params:
            tf.word_wrap = bool(params["word_wrap"]) if params["word_wrap"] is not None else None

        if "auto_size" in params:
            val = params["auto_size"]
            if val is None:
                tf.auto_size = None
            else:
                from pptx.enum.text import MSO_AUTO_SIZE

                auto_map = {
                    "none": MSO_AUTO_SIZE.NONE,
                    "shape_to_fit_text": MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT,
                    "text_to_fit_shape": MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE,
                }
                tf.auto_size = auto_map.get(str(val).lower())

        if "vertical_alignment" in params:
            val = params["vertical_alignment"]
            if val is None:
                tf.vertical_anchor = None
            else:
                from pptx.enum.text import MSO_VERTICAL_ANCHOR

                val_map = {
                    "top": MSO_VERTICAL_ANCHOR.TOP,
                    "middle": MSO_VERTICAL_ANCHOR.MIDDLE,
                    "bottom": MSO_VERTICAL_ANCHOR.BOTTOM,
                }
                tf.vertical_anchor = val_map.get(str(val).lower())

        self._mark_modified(session)
        return {
            "success": True,
            "shape_name": shape.name,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
