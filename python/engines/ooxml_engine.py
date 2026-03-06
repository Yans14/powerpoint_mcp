from __future__ import annotations

import base64
import copy
import os
import shutil
import subprocess
import tempfile
import uuid
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from engines.base import BaseEngine, EngineSession
from errors import BridgeError, ensure
from utils.paths import validate_existing_file, validate_output_file
from utils.units import emu_to_inches, to_emu

_COLOR_NAMESPACES = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _normalize_color(color_value: str) -> str:
    value = color_value.strip()
    if value.startswith("#"):
        value = value[1:]
    if len(value) != 6:
        raise BridgeError(
            code="validation_error",
            message="Color must be 6 hex characters.",
            details={"color": color_value},
        )
    try:
        int(value, 16)
    except ValueError as exc:
        raise BridgeError(
            code="validation_error",
            message="Color must be valid hex.",
            details={"color": color_value},
        ) from exc
    return value.upper()


class OOXMLEngine(BaseEngine):
    name = "OOXML"

    def __init__(self, metadata: dict[str, Any]) -> None:
        self.metadata = metadata
        self.sessions: dict[str, EngineSession] = {}

    def get_engine_info(self, params: dict[str, Any] | None = None) -> dict[str, Any]:
        return {
            "engine": self.name,
            **self.metadata,
        }

    def create_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        width = params.get("width")
        height = params.get("height")
        template_path = params.get("template_path")

        if template_path:
            template = validate_existing_file(str(template_path), expected_suffixes=(".pptx", ".potx"))
            fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
            os.close(fd)
            Path(tmp_path).unlink(missing_ok=True)
            shutil.copy2(template, tmp_path)
            prs = Presentation(str(tmp_path))
            original_path = str(template)
        else:
            prs = Presentation()
            if width is not None:
                prs.slide_width = to_emu(width)
            if height is not None:
                prs.slide_height = to_emu(height)
            fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
            os.close(fd)
            Path(tmp_path).unlink(missing_ok=True)
            prs.save(tmp_path)
            prs = Presentation(tmp_path)
            original_path = ""

        session_id = str(uuid.uuid4())
        session = EngineSession(
            id=session_id,
            original_path=original_path,
            working_path=tmp_path,
            engine=self.name,
            dirty=True,
            extra={"prs": prs},
        )
        self.sessions[session_id] = session

        return {
            "success": True,
            "presentation_id": session_id,
            "engine": self.name,
            "slide_width_emu": prs.slide_width,
            "slide_height_emu": prs.slide_height,
            "slide_count": len(prs.slides),
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def open_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        source = validate_existing_file(str(params["file_path"]), expected_suffixes=(".pptx", ".potx"))

        fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
        os.close(fd)
        Path(tmp_path).unlink(missing_ok=True)
        shutil.copy2(source, tmp_path)

        session_id = str(uuid.uuid4())
        prs = Presentation(tmp_path)
        session = EngineSession(
            id=session_id,
            original_path=str(source),
            working_path=tmp_path,
            engine=self.name,
            dirty=False,
            extra={"prs": prs},
        )
        self.sessions[session_id] = session

        return {
            "success": True,
            "presentation_id": session_id,
            "slide_count": len(prs.slides),
            "layout_names": [layout.name for layout in prs.slide_layouts],
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def save_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        output_path = validate_output_file(str(params["output_path"]), expected_suffixes=(".pptx",))
        prs = self._prs(session)
        prs.save(session.working_path)
        shutil.copy2(session.working_path, output_path)
        session.dirty = False

        return {
            "success": True,
            "presentation_id": session.id,
            "saved_path": str(output_path),
            "file_size_bytes": output_path.stat().st_size,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def close_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        self.sessions.pop(session.id, None)
        try:
            Path(session.working_path).unlink(missing_ok=True)
        except Exception:
            pass

        return {
            "success": True,
            "presentation_id": session.id,
            "closed": True,
        }

    def list_open_presentations(self, params: dict[str, Any] | None = None) -> dict[str, Any]:
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

    def get_presentation_state(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        slides = []
        for index, slide in enumerate(prs.slides, start=1):
            slides.append(
                {
                    "index": index,
                    "title": self._slide_title(slide),
                    "layout": slide.slide_layout.name,
                    "shape_count": len(slide.shapes),
                }
            )

        return {
            "presentation_id": session.id,
            "engine": self.name,
            "slide_count": len(prs.slides),
            "slides": slides,
        }

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
        theme_data = self._extract_theme(Path(session.working_path))

        return {
            "presentation_id": session.id,
            "theme": theme_data,
        }

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

        self._persist(session)

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

        # Copy non-placeholder shapes directly.
        for shape in source_slide.shapes:
            if getattr(shape, "is_placeholder", False):
                continue
            new_el = copy.deepcopy(shape.element)
            duplicate.shapes._spTree.insert_element_before(new_el, "p:extLst")

        # Copy placeholder text where names match.
        for src_placeholder in source_slide.placeholders:
            try:
                dst_placeholder = self._find_placeholder(duplicate, src_placeholder.name)
            except BridgeError:
                continue
            if (
                hasattr(src_placeholder, "has_text_frame")
                and src_placeholder.has_text_frame
                and dst_placeholder.has_text_frame
            ):
                dst_placeholder.text = src_placeholder.text

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

        self._persist(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "source_index": source_index,
            "duplicated_slide_index": duplicate_index,
            "warning": "OOXML duplicate copies placeholders and non-placeholder shapes; complex animations may require COM mode.",
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

        self._persist(session)

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

        self._persist(session)

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

        self._persist(session)

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
            fill.fore_color.rgb = RGBColor.from_string(_normalize_color(str(color_hex)))

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
                grad_stops[0].color.rgb = RGBColor.from_string(_normalize_color(str(grad_start)))
                grad_stops[-1].color.rgb = RGBColor.from_string(_normalize_color(str(grad_end)))
            except Exception:
                warnings.append("Gradient background is partially supported in OOXML mode.")

        self._persist(session)

        payload: dict[str, Any] = {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
        if warnings:
            payload["warning"] = " ".join(warnings)
        return payload

    def get_slide_snapshot(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        width_px = int(params.get("width_px") or 1280)
        self._assert_slide_index(self._prs(session), slide_index)

        soffice = shutil.which("soffice")
        pdftoppm = shutil.which("pdftoppm")

        if not soffice or not pdftoppm:
            raise BridgeError(
                code="dependency_missing",
                message="pptx_get_slide_snapshot requires LibreOffice (soffice) and poppler (pdftoppm) in OOXML mode.",
                details={
                    "missing": [
                        dependency
                        for dependency, ok in {
                            "soffice": bool(soffice),
                            "pdftoppm": bool(pdftoppm),
                        }.items()
                        if not ok
                    ],
                    "install_hint": "Install LibreOffice and poppler-utils to enable slide snapshots.",
                },
            )

        self._persist(session)

        with tempfile.TemporaryDirectory(prefix="pptx-snapshot-") as tmp_dir:
            source_path = Path(session.working_path)
            pdf_out_dir = Path(tmp_dir)

            convert_result = subprocess.run(
                [
                    str(soffice),
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(pdf_out_dir),
                    str(source_path),
                ],
                capture_output=True,
                text=True,
                check=False,
            )

            if convert_result.returncode != 0:
                raise BridgeError(
                    code="engine_error",
                    message="LibreOffice PDF export failed.",
                    details={
                        "stdout": convert_result.stdout,
                        "stderr": convert_result.stderr,
                    },
                )

            pdf_path = pdf_out_dir / f"{source_path.stem}.pdf"
            ensure(pdf_path.exists(), "engine_error", "Expected PDF output was not generated.")

            image_prefix = pdf_out_dir / "slide"
            render_result = subprocess.run(
                [
                    str(pdftoppm),
                    "-f",
                    str(slide_index),
                    "-l",
                    str(slide_index),
                    "-jpeg",
                    "-singlefile",
                    "-scale-to-x",
                    str(width_px),
                    str(pdf_path),
                    str(image_prefix),
                ],
                capture_output=True,
                text=True,
                check=False,
            )
            if render_result.returncode != 0:
                raise BridgeError(
                    code="engine_error",
                    message="pdftoppm rendering failed.",
                    details={
                        "stdout": render_result.stdout,
                        "stderr": render_result.stderr,
                    },
                )

            image_path = Path(f"{image_prefix}.jpg")
            ensure(image_path.exists(), "engine_error", "Slide snapshot image was not generated.")
            encoded = base64.b64encode(image_path.read_bytes()).decode("ascii")

        return {
            "presentation_id": session.id,
            "slide_index": slide_index,
            "mime_type": "image/jpeg",
            "snapshot_base64": encoded,
            "width_px": width_px,
        }

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
        if alignment:
            paragraph.alignment = {
                "left": PP_ALIGN.LEFT,
                "center": PP_ALIGN.CENTER,
                "right": PP_ALIGN.RIGHT,
                "justify": PP_ALIGN.JUSTIFY,
            }[str(alignment)]

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
                run.font.color.rgb = RGBColor.from_string(_normalize_color(str(params["color_hex"])))

        self._persist(session)

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

        self._persist(session)

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

        self._persist(session)

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

    def shutdown(self) -> None:
        for session in list(self.sessions.values()):
            try:
                Path(session.working_path).unlink(missing_ok=True)
            except Exception:
                pass
        self.sessions.clear()

    def _get_session(self, presentation_id: str) -> EngineSession:
        session = self.sessions.get(presentation_id)
        if not session:
            raise BridgeError(
                code="not_found",
                message="Presentation session not found.",
                details={"presentation_id": presentation_id},
            )
        return session

    def _prs(self, session: EngineSession) -> Presentation:
        prs = session.extra.get("prs")
        if prs is None:
            raise BridgeError(code="engine_error", message="Session presentation object is missing.")
        return prs

    def _persist(self, session: EngineSession) -> None:
        prs = self._prs(session)
        prs.save(session.working_path)
        session.dirty = True

    def _assert_slide_index(self, prs: Presentation, slide_index: int) -> None:
        if slide_index < 1 or slide_index > len(prs.slides):
            raise BridgeError(
                code="validation_error",
                message="slide_index is out of bounds.",
                details={"slide_index": slide_index, "slide_count": len(prs.slides)},
            )

    def _get_slide_obj(self, session: EngineSession, slide_index: int):
        prs = self._prs(session)
        self._assert_slide_index(prs, slide_index)
        return prs.slides[slide_index - 1]

    def _slide_title(self, slide) -> str:
        title_shape = slide.shapes.title
        if title_shape is not None and hasattr(title_shape, "text") and title_shape.text:
            return title_shape.text.strip()

        for shape in slide.shapes:
            if hasattr(shape, "has_text_frame") and shape.has_text_frame and shape.text:
                return shape.text.strip().splitlines()[0]

        return ""

    def _find_layout(self, prs: Presentation, layout_name: str):
        for layout in prs.slide_layouts:
            if layout.name == layout_name:
                return layout

        available = [layout.name for layout in prs.slide_layouts]
        raise BridgeError(
            code="validation_error",
            message=f"Layout '{layout_name}' not found.",
            details={"layout_name": layout_name, "available_layouts": available},
        )

    def _placeholder_payload(self, placeholder) -> dict[str, Any]:
        return {
            "name": placeholder.name,
            "idx": int(placeholder.placeholder_format.idx),
            "type": str(placeholder.placeholder_format.type),
            "left_inches": emu_to_inches(placeholder.left),
            "top_inches": emu_to_inches(placeholder.top),
            "width_inches": emu_to_inches(placeholder.width),
            "height_inches": emu_to_inches(placeholder.height),
        }

    def _shape_payload(self, shape) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "shape_id": int(shape.shape_id),
            "name": shape.name,
            "type": str(shape.shape_type),
            "left_inches": emu_to_inches(shape.left),
            "top_inches": emu_to_inches(shape.top),
            "width_inches": emu_to_inches(shape.width),
            "height_inches": emu_to_inches(shape.height),
            "is_placeholder": bool(getattr(shape, "is_placeholder", False)),
        }

        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            payload["text"] = shape.text

        return payload

    def _find_placeholder(self, slide, placeholder_name: str):
        for placeholder in slide.placeholders:
            if placeholder.name == placeholder_name:
                return placeholder

        available = [placeholder.name for placeholder in slide.placeholders]
        raise BridgeError(
            code="not_found",
            message=f"Placeholder '{placeholder_name}' not found.",
            details={"placeholder_name": placeholder_name, "available_placeholders": available},
        )

    def _validate_order(self, prs: Presentation, new_order: list[int]) -> None:
        count = len(prs.slides)
        if len(new_order) != count:
            raise BridgeError(
                code="validation_error",
                message="new_order length must equal slide_count.",
                details={"new_order_length": len(new_order), "slide_count": count},
            )

        expected = set(range(1, count + 1))
        received = set(new_order)
        if expected != received:
            raise BridgeError(
                code="validation_error",
                message="new_order must contain each current slide index exactly once.",
                details={"expected": sorted(expected), "received": new_order},
            )

    def _reorder_slide_ids(self, prs: Presentation, new_order: list[int]) -> None:
        sld_id_lst = prs.slides._sldIdLst
        slide_ids = list(sld_id_lst)
        reordered = [slide_ids[index - 1] for index in new_order]

        for node in list(sld_id_lst):
            sld_id_lst.remove(node)
        for node in reordered:
            sld_id_lst.append(node)

    def _send_shape_to_back(self, slide, shape) -> None:
        sp_tree = slide.shapes._spTree
        element = shape.element
        sp_tree.remove(element)
        sp_tree.insert(2, element)

    def _extract_theme(self, path: Path) -> dict[str, Any]:
        if not path.exists():
            return {"colors": {}, "fonts": {"major": "", "minor": ""}}

        colors: dict[str, str] = {}
        fonts = {"major": "", "minor": ""}

        with zipfile.ZipFile(path, "r") as archive:
            try:
                xml_data = archive.read("ppt/theme/theme1.xml")
            except KeyError:
                return {"colors": colors, "fonts": fonts}

        root = etree.fromstring(xml_data)

        clr_scheme = root.find(".//a:clrScheme", namespaces=_COLOR_NAMESPACES)
        if clr_scheme is not None:
            for child in list(clr_scheme):
                name = etree.QName(child).localname
                srgb = child.find(".//a:srgbClr", namespaces=_COLOR_NAMESPACES)
                sys_clr = child.find(".//a:sysClr", namespaces=_COLOR_NAMESPACES)
                if srgb is not None:
                    colors[name] = srgb.attrib.get("val", "")
                elif sys_clr is not None:
                    colors[name] = sys_clr.attrib.get("lastClr", "")

        major_font = root.find(".//a:fontScheme/a:majorFont/a:latin", namespaces=_COLOR_NAMESPACES)
        minor_font = root.find(".//a:fontScheme/a:minorFont/a:latin", namespaces=_COLOR_NAMESPACES)
        if major_font is not None:
            fonts["major"] = major_font.attrib.get("typeface", "")
        if minor_font is not None:
            fonts["minor"] = minor_font.attrib.get("typeface", "")

        return {"colors": colors, "fonts": fonts}
