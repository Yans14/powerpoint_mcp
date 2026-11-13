"""Shared state and helpers of the OOXMLEngine; the domain mixins build on these."""

from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

from engines.base import BaseEngine, EngineSession
from engines.ooxml.common import _COLOR_NAMESPACES, logger
from errors import SESSION_NOT_FOUND_HINT, BridgeError, ensure
from utils.colors import normalize_color
from utils.units import emu_to_inches
from workspace import Workspace


class OOXMLCore(BaseEngine):
    """Shared state and helpers of the OOXMLEngine; the domain mixins build on these."""

    name = "OOXML"

    _ALIGN_MAP = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }

    _REVERSE_ALIGN_MAP = {v: k for k, v in _ALIGN_MAP.items()}

    _SHAPE_TYPE_MAP = {
        "rectangle": MSO_SHAPE.RECTANGLE,
        "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
        "oval": MSO_SHAPE.OVAL,
        "diamond": MSO_SHAPE.DIAMOND,
        "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
        "right_arrow": MSO_SHAPE.RIGHT_ARROW,
        "left_arrow": MSO_SHAPE.LEFT_ARROW,
        "up_arrow": MSO_SHAPE.UP_ARROW,
        "down_arrow": MSO_SHAPE.DOWN_ARROW,
        "pentagon": MSO_SHAPE.REGULAR_PENTAGON,
        "hexagon": MSO_SHAPE.HEXAGON,
        "chevron": MSO_SHAPE.CHEVRON,
        "star_5_point": MSO_SHAPE.STAR_5_POINT,
        "line_inverse": MSO_SHAPE.LINE_INVERSE,
        "cross": MSO_SHAPE.CROSS,
        "frame": MSO_SHAPE.FRAME,
        "rectangular_callout": MSO_SHAPE.RECTANGULAR_CALLOUT,
        "rounded_rectangular_callout": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
        "cloud_callout": MSO_SHAPE.CLOUD_CALLOUT,
        "cloud": MSO_SHAPE.CLOUD,
    }

    _CHART_TYPE_MAP = {
        "column_clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
        "column_stacked_100": XL_CHART_TYPE.COLUMN_STACKED_100,
        "bar_clustered": XL_CHART_TYPE.BAR_CLUSTERED,
        "bar_stacked": XL_CHART_TYPE.BAR_STACKED,
        "bar_stacked_100": XL_CHART_TYPE.BAR_STACKED_100,
        "line": XL_CHART_TYPE.LINE,
        "line_markers": XL_CHART_TYPE.LINE_MARKERS,
        "line_stacked": XL_CHART_TYPE.LINE_STACKED,
        "pie": XL_CHART_TYPE.PIE,
        "pie_exploded": XL_CHART_TYPE.PIE_EXPLODED,
        "doughnut": XL_CHART_TYPE.DOUGHNUT,
        "area": XL_CHART_TYPE.AREA,
        "area_stacked": XL_CHART_TYPE.AREA_STACKED,
        "area_stacked_100": XL_CHART_TYPE.AREA_STACKED_100,
        "xy_scatter": XL_CHART_TYPE.XY_SCATTER,
        "xy_scatter_lines": XL_CHART_TYPE.XY_SCATTER_LINES,
        "xy_scatter_smooth": XL_CHART_TYPE.XY_SCATTER_SMOOTH,
        "bubble": XL_CHART_TYPE.BUBBLE,
        "radar": XL_CHART_TYPE.RADAR,
        "stock_hlc": XL_CHART_TYPE.STOCK_HLC,
        "stock_ohlc": XL_CHART_TYPE.STOCK_OHLC,
        "three_d_column": XL_CHART_TYPE.THREE_D_COLUMN,
        "three_d_bar_clustered": XL_CHART_TYPE.THREE_D_BAR_CLUSTERED,
        "three_d_pie": XL_CHART_TYPE.THREE_D_PIE,
        "three_d_line": XL_CHART_TYPE.THREE_D_LINE,
    }

    _LEGEND_POSITION_MAP = {
        "bottom": XL_LEGEND_POSITION.BOTTOM,
        "corner": XL_LEGEND_POSITION.CORNER,
        "left": XL_LEGEND_POSITION.LEFT,
        "right": XL_LEGEND_POSITION.RIGHT,
        "top": XL_LEGEND_POSITION.TOP,
    }

    supports_rollback = True

    def __init__(self, metadata: dict[str, Any]) -> None:
        self.metadata = metadata
        self.sessions: dict[str, EngineSession] = {}
        self.workspace = Workspace()

    def shutdown(self) -> None:
        self.sessions.clear()
        self.workspace.cleanup()

    def flush_session(self, session: EngineSession) -> None:
        if session.flushed_revision != session.revision:
            self._prs(session).save(session.working_path)
            session.flushed_revision = session.revision

    def reload_session(self, session: EngineSession) -> None:
        session.extra["prs"] = Presentation(session.working_path)
        session.revision += 1
        session.flushed_revision = session.revision

    def _new_session_files(self) -> tuple[str, str]:
        """Create a private directory for one session. Returns (directory, working file path)."""
        session_dir = self.workspace.new_dir("session-")
        return str(session_dir), str(session_dir / "presentation.pptx")

    def _remove_session_files(self, session: EngineSession) -> None:
        # Holds the working copy plus anything derived from it (rollback snapshot, rendered PDF).
        session_dir = session.extra.get("session_dir")
        if session_dir:
            shutil.rmtree(session_dir, ignore_errors=True)

    def _get_session(self, presentation_id: str) -> EngineSession:
        session = self.sessions.get(presentation_id)
        if not session:
            raise BridgeError(
                code="not_found",
                message="Presentation session not found.",
                details={"presentation_id": presentation_id},
                hint=SESSION_NOT_FOUND_HINT,
            )
        return session

    def _prs(self, session: EngineSession) -> Presentation:
        prs = session.extra.get("prs")
        if prs is None:
            raise BridgeError(code="engine_error", message="Session presentation object is missing.")
        return prs

    def _mark_modified(self, session: EngineSession) -> None:
        """Record an in-memory edit.

        The working file is not rewritten here: serialising the whole deck after every mutation
        dominates the cost of small edits on large files. Readers of the file call flush_session.
        """
        session.revision += 1
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
            code="not_found",
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

    def _apply_run_formatting(self, run, data: dict[str, Any]) -> None:
        """Apply font formatting from a dict of optional formatting keys."""
        if data.get("font_name"):
            run.font.name = str(data["font_name"])
        if data.get("font_size_pt"):
            run.font.size = Pt(float(data["font_size_pt"]))
        if data.get("bold") is not None:
            run.font.bold = bool(data["bold"])
        if data.get("italic") is not None:
            run.font.italic = bool(data["italic"])
        if data.get("underline") is not None:
            run.font.underline = bool(data["underline"])
        if data.get("color_hex"):
            run.font.color.rgb = RGBColor.from_string(normalize_color(str(data["color_hex"])))

    def _write_paragraphs(self, text_frame, paragraphs_data: list[dict[str, Any]]) -> None:
        """Write rich paragraph data into a text frame (clears existing content)."""
        text_frame.clear()
        for p_idx, p_data in enumerate(paragraphs_data):
            paragraph = text_frame.paragraphs[0] if p_idx == 0 else text_frame.add_paragraph()

            alignment = p_data.get("alignment")
            if alignment and alignment in self._ALIGN_MAP:
                paragraph.alignment = self._ALIGN_MAP[alignment]

            level = p_data.get("level")
            if level is not None:
                paragraph.level = int(level)

            runs_data = p_data.get("runs", [])
            if not runs_data:
                run = paragraph.add_run()
                run.text = str(p_data.get("text", ""))
            else:
                for r_data in runs_data:
                    run = paragraph.add_run()
                    run.text = str(r_data.get("text", ""))
                    self._apply_run_formatting(run, r_data)

    def _require_shape(self, slide, params: dict[str, Any]):
        """Resolve a shape from params by shape_name or shape_id, raising on missing."""
        identifier = params.get("shape_name") or params.get("shape_id")
        ensure(identifier is not None, "validation_error", "shape_name or shape_id is required.")
        return self._find_shape(slide, identifier)

    def _extract_text_items(self, shapes, depth: int = 0) -> list[dict[str, Any]]:
        """Recursively extract text from all shapes including tables, groups, etc."""
        items: list[dict[str, Any]] = []

        for shape in shapes:
            shape_info: dict[str, Any] = {
                "shape_name": shape.name,
                "shape_id": int(shape.shape_id),
                "shape_type": str(shape.shape_type),
                "is_placeholder": bool(getattr(shape, "is_placeholder", False)),
                "left_inches": emu_to_inches(shape.left) if shape.left is not None else None,
                "top_inches": emu_to_inches(shape.top) if shape.top is not None else None,
                "width_inches": emu_to_inches(shape.width) if shape.width is not None else None,
                "height_inches": emu_to_inches(shape.height) if shape.height is not None else None,
            }

            # --- Table shapes ---
            if shape.has_table:
                table = shape.table
                table_text_parts: list[str] = []
                table_cells: list[dict[str, Any]] = []
                for r_idx in range(len(table.rows)):
                    for c_idx in range(len(table.columns)):
                        cell = table.cell(r_idx, c_idx)
                        cell_text = cell.text or ""
                        table_text_parts.append(cell_text)
                        paragraphs = self._extract_paragraphs(cell.text_frame) if cell.text_frame else []
                        table_cells.append(
                            {
                                "row": r_idx,
                                "col": c_idx,
                                "text": cell_text,
                                "paragraphs": paragraphs,
                            }
                        )

                shape_info["content_type"] = "table"
                shape_info["rows"] = len(table.rows)
                shape_info["cols"] = len(table.columns)
                shape_info["table_cells"] = table_cells
                shape_info["raw_text"] = "\n".join(table_text_parts)
                shape_info["paragraphs"] = []
                items.append(shape_info)

            # --- Group shapes —- recurse into children ---
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                shape_info["content_type"] = "group"
                shape_info["raw_text"] = ""
                shape_info["paragraphs"] = []
                items.append(shape_info)
                # Recursively extract text from group children
                if hasattr(shape, "shapes"):
                    child_items = self._extract_text_items(shape.shapes, depth + 1)
                    for child in child_items:
                        child["parent_group"] = shape.name
                    items.extend(child_items)

            # --- Shapes with text frames (text boxes, autoshapes, placeholders) ---
            elif hasattr(shape, "has_text_frame") and shape.has_text_frame:
                paragraphs = self._extract_paragraphs(shape.text_frame)
                raw_text = shape.text or ""
                shape_info["content_type"] = "text"
                shape_info["raw_text"] = raw_text
                shape_info["paragraphs"] = paragraphs
                items.append(shape_info)

            # --- Picture shapes ---
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_info["content_type"] = "picture"
                shape_info["raw_text"] = ""
                shape_info["paragraphs"] = []
                items.append(shape_info)

            # --- Other shapes (connectors, charts, etc) ---
            else:
                shape_info["content_type"] = "other"
                shape_info["raw_text"] = ""
                shape_info["paragraphs"] = []
                items.append(shape_info)

        return items

    def _extract_paragraphs(self, text_frame) -> list[dict[str, Any]]:
        """Extract paragraph + run structure from a text frame."""
        paragraphs: list[dict[str, Any]] = []
        for paragraph in text_frame.paragraphs:
            runs: list[dict[str, Any]] = []
            for run in paragraph.runs:
                color_hex = None
                try:
                    if run.font.color and run.font.color.rgb:
                        color_hex = str(run.font.color.rgb)
                except Exception:
                    pass

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

            alignment = None
            if paragraph.alignment is not None:
                alignment = self._REVERSE_ALIGN_MAP.get(paragraph.alignment)

            paragraphs.append(
                {
                    "text": paragraph.text,
                    "level": paragraph.level,
                    "alignment": alignment,
                    "runs": runs,
                }
            )
        return paragraphs

    def _find_shape(self, slide, identifier):
        """Find a shape by name (str) or shape_id (int)."""
        # Try as integer shape_id first
        try:
            shape_id = int(identifier)
            for shape in slide.shapes:
                if int(shape.shape_id) == shape_id:
                    return shape
                # Also search inside groups
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP and hasattr(shape, "shapes"):
                    for child in shape.shapes:
                        if int(child.shape_id) == shape_id:
                            return child
        except (ValueError, TypeError):
            pass

        # Try as name
        name = str(identifier)
        for shape in slide.shapes:
            if shape.name == name:
                return shape
            # Also search inside groups
            if shape.shape_type == MSO_SHAPE_TYPE.GROUP and hasattr(shape, "shapes"):
                for child in shape.shapes:
                    if child.name == name:
                        return child

        # Build available list for error
        available = [f"{s.name} (id={s.shape_id})" for s in slide.shapes]
        raise BridgeError(
            code="not_found",
            message=f"Shape '{identifier}' not found.",
            details={"identifier": str(identifier), "available_shapes": available},
        )

    def _detailed_shape_payload(self, shape) -> dict[str, Any]:
        """Return rich details about a shape for get_shape_details."""
        payload: dict[str, Any] = {
            "shape_id": int(shape.shape_id),
            "name": shape.name,
            "shape_type": str(shape.shape_type),
            "left_inches": emu_to_inches(shape.left) if shape.left is not None else None,
            "top_inches": emu_to_inches(shape.top) if shape.top is not None else None,
            "width_inches": emu_to_inches(shape.width) if shape.width is not None else None,
            "height_inches": emu_to_inches(shape.height) if shape.height is not None else None,
            "rotation": float(shape.rotation) if hasattr(shape, "rotation") else 0.0,
            "is_placeholder": bool(getattr(shape, "is_placeholder", False)),
        }

        # Text content
        if hasattr(shape, "has_text_frame") and shape.has_text_frame:
            payload["has_text"] = True
            payload["text"] = shape.text
            payload["paragraphs"] = self._extract_paragraphs(shape.text_frame)
        else:
            payload["has_text"] = False

        # Table content
        if hasattr(shape, "has_table") and shape.has_table:
            table = shape.table
            payload["has_table"] = True
            payload["table_rows"] = len(table.rows)
            payload["table_cols"] = len(table.columns)
            cells = []
            for r_idx in range(len(table.rows)):
                row_cells = []
                for c_idx in range(len(table.columns)):
                    cell = table.cell(r_idx, c_idx)
                    row_cells.append(self._extract_cell_data(cell))
                cells.append(row_cells)
            payload["table_cells"] = cells
        else:
            payload["has_table"] = False

        # Picture
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            payload["is_picture"] = True
            try:
                payload["image_content_type"] = shape.image.content_type
            except Exception:
                payload["image_content_type"] = "unknown"
        else:
            payload["is_picture"] = False

        # Group
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP and hasattr(shape, "shapes"):
            payload["is_group"] = True
            payload["child_count"] = len(shape.shapes)
            payload["children"] = [self._shape_payload(child) for child in shape.shapes]
        else:
            payload["is_group"] = False

        return payload

    def _extract_cell_data(self, cell) -> dict[str, Any]:
        """Extract text and formatting from a table cell."""
        paragraphs = self._extract_paragraphs(cell.text_frame) if cell.text_frame else []
        return {
            "text": cell.text or "",
            "paragraphs": paragraphs,
        }

    def _write_cell(self, cell, data: dict[str, Any]) -> None:
        """Write text and optional formatting to a table cell."""
        text = str(data.get("text", ""))
        cell.text = text

        # Apply formatting to the first paragraph/run
        if cell.text_frame and cell.text_frame.paragraphs:
            paragraph = cell.text_frame.paragraphs[0]
            if paragraph.runs:
                self._apply_run_formatting(paragraph.runs[0], data)

            # Cell fill
            fill_hex = data.get("fill_hex")
            if fill_hex:
                try:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(normalize_color(str(fill_hex)))
                except Exception as exc:
                    logger.warning("could not set cell fill: %s", exc)

    def _replace_in_shapes(self, shapes, find_text: str, replace_text: str, case_sensitive: bool) -> int:
        replacements = 0
        search_pattern = re.compile(
            re.escape(find_text) if case_sensitive else re.escape(find_text), 0 if case_sensitive else re.IGNORECASE
        )

        for shape in shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if search_pattern.search(run.text):
                            old_text = run.text
                            run.text = search_pattern.sub(replace_text, run.text)
                            if run.text != old_text:
                                replacements += 1
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for paragraph in cell.text_frame.paragraphs:
                                for run in paragraph.runs:
                                    if search_pattern.search(run.text):
                                        old_text = run.text
                                        run.text = search_pattern.sub(replace_text, run.text)
                                        if run.text != old_text:
                                            replacements += 1
            elif shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                replacements += self._replace_in_shapes(shape.shapes, find_text, replace_text, case_sensitive)
        return replacements
