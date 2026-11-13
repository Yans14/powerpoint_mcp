"""Tables."""

from __future__ import annotations

from typing import Any

from engines.com.common import (
    _hex_to_bgr_int,
    _to_points,
)
from errors import ensure


class COMTablesMixin:
    """Tables."""

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
