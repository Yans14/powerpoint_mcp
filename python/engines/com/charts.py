"""Charts."""

from __future__ import annotations

from typing import Any

from engines.com.common import (
    _LEGEND_POS_MAP,
    _XL_CHART_TYPE_MAP,
    _XL_CHART_TYPE_REVERSE,
    _to_points,
)
from errors import ensure


class COMChartsMixin:
    """Charts."""

    def add_chart(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_chart_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_chart_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        chart_type_str = str(params["chart_type"]).lower()
        ensure(chart_type_str in _XL_CHART_TYPE_MAP, "validation_error", f"Unknown chart type: {chart_type_str}")
        xl_type = _XL_CHART_TYPE_MAP[chart_type_str]

        left = _to_points(params.get("left", "1in"))
        top = _to_points(params.get("top", "1in"))
        width = _to_points(params.get("width", "8in"))
        height = _to_points(params.get("height", "5in"))

        # Try AddChart2 (Office 2013+), fallback to AddChart
        try:
            chart_shape = slide.Shapes.AddChart2(-1, xl_type, left, top, width, height)
        except Exception:
            chart_shape = slide.Shapes.AddChart(xl_type, left, top, width, height)

        chart = chart_shape.Chart
        series_data = params.get("series", [])
        categories = params.get("categories", [])
        is_xy = chart_type_str.startswith("xy_")
        is_bubble = chart_type_str.startswith("bubble")

        if series_data:
            chart.ChartData.Activate()
            wb = chart.ChartData.Workbook
            try:
                wb.Application.Visible = False
            except Exception:
                pass
            try:
                ws = wb.Worksheets(1)
                ws.Cells.Clear()

                if is_xy or is_bubble:
                    for s_idx, s in enumerate(series_data):
                        col_offset = s_idx * (3 if is_bubble else 2)
                        s_name = str(s.get("name", f"Series {s_idx + 1}"))
                        ws.Cells(1, col_offset + 1).Value = f"{s_name} X"
                        ws.Cells(1, col_offset + 2).Value = f"{s_name} Y"
                        if is_bubble:
                            ws.Cells(1, col_offset + 3).Value = f"{s_name} Size"
                        for dp_idx, dp in enumerate(s.get("data_points", [])):
                            ws.Cells(dp_idx + 2, col_offset + 1).Value = dp.get("x", 0)
                            ws.Cells(dp_idx + 2, col_offset + 2).Value = dp.get("y", 0)
                            if is_bubble:
                                ws.Cells(dp_idx + 2, col_offset + 3).Value = dp.get("size", 10)
                    max_pts = max((len(s.get("data_points", [])) for s in series_data), default=0)
                    total_cols = len(series_data) * (3 if is_bubble else 2)
                    data_range = ws.Range(ws.Cells(1, 1), ws.Cells(max_pts + 1, total_cols))
                    chart.SetSourceData(data_range)
                else:
                    # Category charts: col A = categories, cols B+ = series values
                    ws.Cells(1, 1).Value = "Category"
                    for s_idx, s in enumerate(series_data):
                        ws.Cells(1, s_idx + 2).Value = str(s.get("name", f"Series {s_idx + 1}"))
                    for c_idx, cat in enumerate(categories):
                        ws.Cells(c_idx + 2, 1).Value = str(cat)
                    for s_idx, s in enumerate(series_data):
                        for v_idx, val in enumerate(s.get("values", [])):
                            ws.Cells(v_idx + 2, s_idx + 2).Value = val
                    num_rows = max(len(categories), max((len(s.get("values", [])) for s in series_data), default=0))
                    data_range = ws.Range(ws.Cells(1, 1), ws.Cells(num_rows + 1, len(series_data) + 1))
                    chart.SetSourceData(data_range)
            finally:
                try:
                    wb.Close(False)
                except Exception:
                    pass

        # Apply optional chart formatting
        if params.get("title"):
            chart.HasTitle = True
            chart.ChartTitle.Text = str(params["title"])
        if params.get("has_legend") is not None:
            chart.HasLegend = bool(params["has_legend"])
        if params.get("legend_position") and chart.HasLegend:
            pos = str(params["legend_position"]).lower()
            if pos in _LEGEND_POS_MAP:
                chart.Legend.Position = _LEGEND_POS_MAP[pos]
        if params.get("has_data_labels") is not None:
            for s_idx in range(1, chart.SeriesCollection().Count + 1):
                chart.SeriesCollection(s_idx).HasDataLabels = bool(params["has_data_labels"])
        if params.get("chart_style"):
            try:
                chart.ChartStyle = int(params["chart_style"])
            except Exception:
                pass

        return {
            "shape_name": str(chart_shape.Name),
            "shape_id": int(chart_shape.Id),
            "chart_type": chart_type_str,
        }

    def get_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._get_chart_data_ts, pid, si, params)
        return {"presentation_id": pid, "slide_index": si, **result}

    def _get_chart_data_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(getattr(shape, "HasChart", False)), "conflict", "Shape is not a chart")
        chart = shape.Chart

        chart_type_int = int(chart.ChartType)
        chart_type_str = _XL_CHART_TYPE_REVERSE.get(chart_type_int, f"unknown_{chart_type_int}")

        result: dict[str, Any] = {
            "shape_name": str(shape.Name),
            "chart_type": chart_type_str,
            "categories": [],
            "series": [],
            "has_legend": bool(chart.HasLegend),
            "has_title": bool(chart.HasTitle),
            "has_data_labels": False,
        }

        if chart.HasTitle:
            result["title"] = str(chart.ChartTitle.Text)

        try:
            chart.ChartData.Activate()
            wb = chart.ChartData.Workbook
            try:
                wb.Application.Visible = False
            except Exception:
                pass
            try:
                ws = wb.Worksheets(1)
                used = ws.UsedRange
                rows = int(used.Rows.Count)
                cols = int(used.Columns.Count)

                # Read categories from column A (rows 2+)
                cats = []
                for r in range(2, rows + 1):
                    val = ws.Cells(r, 1).Value
                    cats.append(str(val) if val is not None else "")
                result["categories"] = cats

                # Read series from columns B+
                series_list = []
                for c in range(2, cols + 1):
                    s_name = ws.Cells(1, c).Value
                    vals = []
                    for r in range(2, rows + 1):
                        v = ws.Cells(r, c).Value
                        vals.append(float(v) if v is not None else None)
                    series_list.append({"name": str(s_name) if s_name else f"Series {c - 1}", "values": vals})
                result["series"] = series_list
            finally:
                try:
                    wb.Close(False)
                except Exception:
                    pass
        except Exception:
            pass

        # Check data labels on first series
        try:
            if chart.SeriesCollection().Count > 0:
                result["has_data_labels"] = bool(chart.SeriesCollection(1).HasDataLabels)
        except Exception:
            pass

        return result

    def update_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._update_chart_data_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _update_chart_data_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(getattr(shape, "HasChart", False)), "conflict", "Shape is not a chart")
        chart = shape.Chart
        chart_type_int = int(chart.ChartType)
        chart_type_str = _XL_CHART_TYPE_REVERSE.get(chart_type_int, f"unknown_{chart_type_int}")

        series_data = params.get("series", [])
        categories = params.get("categories")

        if series_data or categories:
            chart.ChartData.Activate()
            wb = chart.ChartData.Workbook
            try:
                wb.Application.Visible = False
            except Exception:
                pass
            try:
                ws = wb.Worksheets(1)

                if categories:
                    for c_idx, cat in enumerate(categories):
                        ws.Cells(c_idx + 2, 1).Value = str(cat)

                for s_idx, s in enumerate(series_data):
                    col = s_idx + 2
                    if s.get("name"):
                        ws.Cells(1, col).Value = str(s["name"])
                    for v_idx, val in enumerate(s.get("values", [])):
                        ws.Cells(v_idx + 2, col).Value = val

                num_cats = len(categories) if categories else 0
                num_vals = max((len(s.get("values", [])) for s in series_data), default=0) if series_data else 0
                num_rows = max(num_cats, num_vals)
                if num_rows > 0:
                    num_cols = len(series_data) + 1 if series_data else int(ws.UsedRange.Columns.Count)
                    data_range = ws.Range(ws.Cells(1, 1), ws.Cells(num_rows + 1, num_cols))
                    chart.SetSourceData(data_range)
            finally:
                try:
                    wb.Close(False)
                except Exception:
                    pass

        return {"shape_name": str(shape.Name), "chart_type": chart_type_str}

    def set_chart_style(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        shape_name = self.worker.call(self._set_chart_style_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": shape_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_chart_style_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(bool(getattr(shape, "HasChart", False)), "conflict", "Shape is not a chart")
        chart = shape.Chart

        if params.get("has_legend") is not None:
            chart.HasLegend = bool(params["has_legend"])
        if params.get("legend_position") and chart.HasLegend:
            pos = str(params["legend_position"]).lower()
            if pos in _LEGEND_POS_MAP:
                chart.Legend.Position = _LEGEND_POS_MAP[pos]
        if params.get("has_data_labels") is not None:
            for s_idx in range(1, chart.SeriesCollection().Count + 1):
                chart.SeriesCollection(s_idx).HasDataLabels = bool(params["has_data_labels"])
        if params.get("chart_style"):
            try:
                chart.ChartStyle = int(params["chart_style"])
            except Exception:
                pass
        if "title" in params:
            title_val = params["title"]
            if title_val is None or title_val == "":
                chart.HasTitle = False
            else:
                chart.HasTitle = True
                chart.ChartTitle.Text = str(title_val)

        return str(shape.Name)
