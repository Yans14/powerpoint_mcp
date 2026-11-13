"""Shared state and helpers of the COMEngine; the domain mixins build on these."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from com_worker import COMWorker
from engines.base import BaseEngine, EngineSession
from engines.com.common import (
    _ALIGN_MAP,
    _hex_to_bgr_int,
    win32com,
)
from errors import SESSION_NOT_FOUND_HINT, BridgeError, ensure


class COMCore(BaseEngine):
    """Shared state and helpers of the COMEngine; the domain mixins build on these."""

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

    def _list_layout_names(self, session_id: str) -> list[str]:
        return [item["name"] for item in self._get_layouts_threadsafe(session_id)]

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
            raise BridgeError(
                "not_found",
                "Presentation session not found.",
                {"presentation_id": presentation_id},
                SESSION_NOT_FOUND_HINT,
            )
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
            "not_found",
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
