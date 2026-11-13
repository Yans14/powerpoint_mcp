"""Tables."""

from __future__ import annotations

from typing import Any

from lxml import etree

from errors import ensure
from utils.units import to_emu


class OOXMLTablesMixin:
    """Tables."""

    def add_table(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        rows = int(params["rows"])
        cols = int(params["cols"])
        ensure(rows >= 1, "validation_error", "rows must be >= 1")
        ensure(cols >= 1, "validation_error", "cols must be >= 1")

        left = to_emu(params["left"])
        top = to_emu(params["top"])
        width = to_emu(params["width"])
        height = to_emu(params["height"])

        from pptx.util import Emu as EmuVal

        table_shape = slide.shapes.add_table(rows, cols, EmuVal(left), EmuVal(top), EmuVal(width), EmuVal(height))
        table = table_shape.table

        # Optional initial data
        data = params.get("data")
        if data:
            for r_idx, row_data in enumerate(data):
                if r_idx >= rows:
                    break
                for c_idx, cell_text in enumerate(row_data):
                    if c_idx >= cols:
                        break
                    table.cell(r_idx, c_idx).text = str(cell_text)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": table_shape.name,
            "shape_id": int(table_shape.shape_id),
            "rows": rows,
            "cols": cols,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def get_table(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._find_shape(slide, params.get("shape_name") or params.get("shape_id"))
        ensure(shape.has_table, "conflict", "Shape is not a table.")

        table = shape.table
        cells = []
        for r_idx in range(len(table.rows)):
            row_cells = []
            for c_idx in range(len(table.columns)):
                cell = table.cell(r_idx, c_idx)
                cell_data = self._extract_cell_data(cell)
                row_cells.append(cell_data)
            cells.append(row_cells)

        return {
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "rows": len(table.rows),
            "cols": len(table.columns),
            "cells": cells,
        }

    def set_table_cell(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._find_shape(slide, params.get("shape_name") or params.get("shape_id"))
        ensure(shape.has_table, "conflict", "Shape is not a table.")

        table = shape.table
        row = int(params["row"])
        col = int(params["col"])
        ensure(0 <= row < len(table.rows), "validation_error", "row out of bounds")
        ensure(0 <= col < len(table.columns), "validation_error", "col out of bounds")

        cell = table.cell(row, col)
        self._write_cell(cell, params)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "row": row,
            "col": col,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_table_data(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._find_shape(slide, params.get("shape_name") or params.get("shape_id"))
        ensure(shape.has_table, "conflict", "Shape is not a table.")

        table = shape.table
        data = params["data"]

        for r_idx, row_data in enumerate(data):
            if r_idx >= len(table.rows):
                break
            for c_idx, cell_data in enumerate(row_data):
                if c_idx >= len(table.columns):
                    break
                cell = table.cell(r_idx, c_idx)
                if isinstance(cell_data, str):
                    cell.text = cell_data
                elif isinstance(cell_data, dict):
                    self._write_cell(cell, cell_data)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "rows_written": min(len(data), len(table.rows)),
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_table_cell_merge(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        ensure(shape.has_table, "conflict", "Shape does not contain a table.")
        table = shape.table

        start_row = int(params["start_row"])
        start_col = int(params["start_col"])
        end_row = int(params["end_row"])
        end_col = int(params["end_col"])

        ensure(
            0 <= start_row <= end_row < len(table.rows),
            "validation_error",
            f"Row indices must satisfy 0 <= start_row <= end_row < {len(table.rows)}.",
        )
        ensure(
            0 <= start_col <= end_col < len(table.columns),
            "validation_error",
            f"Col indices must satisfy 0 <= start_col <= end_col < {len(table.columns)}.",
        )

        # Use python-pptx's native merge
        start_cell = table.cell(start_row, start_col)
        end_cell = table.cell(end_row, end_col)
        start_cell.merge(end_cell)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "merged_range": f"({start_row},{start_col})->({end_row},{end_col})",
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_table_style(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)
        ensure(shape.has_table, "conflict", "Shape does not contain a table.")

        table = shape.table

        if "first_row" in params:
            table.first_row = bool(params["first_row"])
        if "last_row" in params:
            table.last_row = bool(params["last_row"])
        if "first_col" in params:
            table.first_col = bool(params["first_col"])
        if "last_col" in params:
            table.last_col = bool(params["last_col"])
        if "banded_rows" in params:
            table.horz_banding = bool(params["banded_rows"])
        if "banded_cols" in params:
            table.vert_banding = bool(params["banded_cols"])

        # Hack to set style ID if provided (python-pptx doesn't formally expose style_id assignment easily)
        if "style_id" in params:
            style_id = str(params["style_id"]).strip()
            # The tblPr element holds the style_id reference
            tbl_pr = shape.element.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}tblPr")
            if tbl_pr is not None:
                tbl_style_id = tbl_pr.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId")
                if tbl_style_id is None:
                    # Create it if it doesn't exist
                    tbl_style_id = etree.SubElement(
                        tbl_pr, "{http://schemas.openxmlformats.org/drawingml/2006/main}tableStyleId"
                    )
                tbl_style_id.text = style_id

        self._mark_modified(session)

        return {
            "success": True,
            "shape_name": shape.name,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
