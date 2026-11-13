"""Slide-level operations: add, duplicate, delete, reorder, background, notes, text search."""

from __future__ import annotations

from typing import Any

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE

from engines.ooxml_rels import copy_slide_content
from errors import ensure
from utils.colors import normalize_color
from utils.paths import validate_existing_file
from utils.units import emu_to_inches


class OOXMLSlidesMixin:
    """Slide-level operations: add, duplicate, delete, reorder, background, notes, text search."""

    def get_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        notes_text = ""
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = slide.notes_slide.notes_text_frame.text

        return {
            "presentation_id": session.id,
            "slide_index": slide_index,
            "title": self._slide_title(slide),
            "layout": slide.slide_layout.name,
            "shapes": [self._shape_payload(shape) for shape in slide.shapes],
            "placeholders": [self._placeholder_payload(placeholder) for placeholder in slide.placeholders],
            "notes": notes_text,
        }

    def add_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)
        layout_name = str(params["layout_name"])
        layout = self._find_layout(prs, layout_name)

        new_slide = prs.slides.add_slide(layout)
        new_index = len(prs.slides)

        requested_position = params.get("position")
        if requested_position is not None:
            position = int(requested_position)
            ensure(position >= 1 and position <= len(prs.slides), "validation_error", "position is out of bounds")
            if position != new_index:
                order = list(range(1, len(prs.slides) + 1))
                moved = order.pop(new_index - 1)
                order.insert(position - 1, moved)
                self._reorder_slide_ids(prs, order)
                new_index = position

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "added_slide_index": new_index,
            "layout_name": layout_name,
            "placeholders": [self._placeholder_payload(placeholder) for placeholder in new_slide.placeholders],
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def duplicate_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)
        source_index = int(params["source_index"])
        source_slide = self._get_slide_obj(session, source_index)

        duplicate = prs.slides.add_slide(source_slide.slide_layout)

        # Copy all shapes, background and slide-level data, carrying image/chart/link relationships along.
        copy_warnings = copy_slide_content(source_slide, duplicate)

        if source_slide.has_notes_slide and source_slide.notes_slide.notes_text_frame:
            duplicate.notes_slide.notes_text_frame.text = source_slide.notes_slide.notes_text_frame.text

        duplicate_index = len(prs.slides)
        requested_position = params.get("target_position")
        if requested_position is not None:
            position = int(requested_position)
            ensure(
                position >= 1 and position <= len(prs.slides), "validation_error", "target_position is out of bounds"
            )
            if position != duplicate_index:
                order = list(range(1, len(prs.slides) + 1))
                moved = order.pop(duplicate_index - 1)
                order.insert(position - 1, moved)
                self._reorder_slide_ids(prs, order)
                duplicate_index = position

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "source_index": source_index,
            "duplicated_slide_index": duplicate_index,
            "warnings": copy_warnings,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def delete_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)
        slide_index = int(params["slide_index"])
        self._assert_slide_index(prs, slide_index)

        sld_id_lst = prs.slides._sldIdLst
        slide_ids = list(sld_id_lst)
        slide_id = slide_ids[slide_index - 1]
        rel_id = slide_id.rId
        prs.part.drop_rel(rel_id)
        sld_id_lst.remove(slide_id)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "deleted_slide_index": slide_index,
            "warning": f"Slide indices above {slide_index} shifted down by 1.",
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def reorder_slides(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        new_order = [int(value) for value in params["new_order"]]
        self._validate_order(prs, new_order)
        self._reorder_slide_ids(prs, new_order)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "new_order": new_order,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def move_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        from_index = int(params["from_index"])
        to_index = int(params["to_index"])
        self._assert_slide_index(prs, from_index)
        self._assert_slide_index(prs, to_index)

        order = list(range(1, len(prs.slides) + 1))
        moved = order.pop(from_index - 1)
        order.insert(to_index - 1, moved)
        self._reorder_slide_ids(prs, order)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "from_index": from_index,
            "to_index": to_index,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_slide_background(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        warnings: list[str] = []

        color_hex = params.get("color_hex")
        image_path = params.get("image_path")
        grad_start = params.get("gradient_start_color_hex")
        grad_end = params.get("gradient_end_color_hex")

        ensure(
            any([color_hex, image_path, grad_start and grad_end]), "validation_error", "No background input provided."
        )

        if color_hex:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor.from_string(normalize_color(str(color_hex)))

        if image_path:
            image = validate_existing_file(str(image_path), expected_suffixes=(".png", ".jpg", ".jpeg", ".bmp", ".gif"))
            picture = slide.shapes.add_picture(str(image), 0, 0, prs.slide_width, prs.slide_height)
            self._send_shape_to_back(slide, picture)
            warnings.append(
                "OOXML fallback inserts a full-slide picture shape at back. Use COM mode for native background picture fidelity."
            )

        if grad_start and grad_end:
            fill = slide.background.fill
            try:
                fill.gradient()
                grad_stops = fill.gradient_stops
                grad_stops[0].color.rgb = RGBColor.from_string(normalize_color(str(grad_start)))
                grad_stops[-1].color.rgb = RGBColor.from_string(normalize_color(str(grad_end)))
            except Exception:
                warnings.append("Gradient background is partially supported in OOXML mode.")

        self._mark_modified(session)

        payload: dict[str, Any] = {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
        if warnings:
            payload["warning"] = " ".join(warnings)
        return payload

    def get_slide_text(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        text_items = self._extract_text_items(slide.shapes)
        total_text = "".join(item.get("raw_text", "") for item in text_items)

        return {
            "presentation_id": session.id,
            "slide_index": slide_index,
            "total_text_length": len(total_text),
            "item_count": len(text_items),
            "text_items": text_items,
        }

    def set_slide_notes(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        notes_text = str(params.get("notes_text", ""))
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = notes_text

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "notes_length": len(notes_text),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def get_slide_xml(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        xml_bytes = etree.tostring(slide._element, pretty_print=True, encoding="unicode")

        return {
            "presentation_id": session.id,
            "slide_index": slide_index,
            "xml_content": xml_bytes,
        }

    def find_replace_text(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        find_text = str(params["find_text"])
        replace_text = str(params["replace_text"])
        case_sensitive = bool(params.get("case_sensitive", True))

        # Optional: limit to specific slides
        slide_indices = params.get("slide_indices")
        if slide_indices:
            slides = [prs.slides[i - 1] for i in slide_indices if 1 <= i <= len(prs.slides)]
        else:
            slides = list(prs.slides)

        total_replacements = 0

        for slide in slides:
            total_replacements += self._replace_in_shapes(slide.shapes, find_text, replace_text, case_sensitive)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "find_text": find_text,
            "replace_text": replace_text,
            "total_replacements": total_replacements,
            "slides_searched": len(slides),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def get_slide_shapes(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shapes_list: list[dict[str, Any]] = []
        for shape in slide.shapes:
            shape_info: dict[str, Any] = {
                "shape_id": int(shape.shape_id),
                "name": shape.name,
                "shape_type": str(shape.shape_type).split("(")[0].strip() if shape.shape_type else "unknown",
                "left": emu_to_inches(shape.left) if shape.left is not None else None,
                "top": emu_to_inches(shape.top) if shape.top is not None else None,
                "width": emu_to_inches(shape.width) if shape.width is not None else None,
                "height": emu_to_inches(shape.height) if shape.height is not None else None,
            }

            # Add key flags
            shape_info["is_placeholder"] = shape.is_placeholder
            shape_info["has_text_frame"] = shape.has_text_frame
            shape_info["has_table"] = shape.has_table
            shape_info["has_chart"] = shape.has_chart

            if shape.is_placeholder:
                try:
                    shape_info["placeholder_name"] = shape.placeholder_format.idx
                except Exception:
                    pass

            if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                shape_info["child_count"] = len(shape.shapes)

            shapes_list.append(shape_info)

        return {
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_count": len(shapes_list),
            "shapes": shapes_list,
        }
