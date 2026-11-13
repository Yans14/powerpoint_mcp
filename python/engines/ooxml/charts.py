"""Charts."""

from __future__ import annotations

from typing import Any

from pptx.util import Emu

from engines.ooxml.common import logger
from errors import ensure
from utils.units import to_emu


class OOXMLChartsMixin:
    """Charts."""

    def add_chart(self, params: dict[str, Any]) -> dict[str, Any]:
        from pptx.chart.data import BubbleChartData, CategoryChartData, XyChartData

        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        chart_type_str = str(params["chart_type"]).lower()
        xl_type = self._CHART_TYPE_MAP.get(chart_type_str)
        ensure(
            xl_type is not None,
            "validation_error",
            f"Unknown chart_type '{chart_type_str}'.",
            {"available": sorted(self._CHART_TYPE_MAP.keys())},
        )

        left = to_emu(params.get("left", "1in"))
        top = to_emu(params.get("top", "1in"))
        width = to_emu(params.get("width", "8in"))
        height = to_emu(params.get("height", "5in"))

        series_data = params.get("series", [])
        ensure(len(series_data) > 0, "validation_error", "At least one series is required.")

        # Determine data class based on chart type
        is_xy = chart_type_str.startswith("xy_scatter")
        is_bubble = chart_type_str.startswith("bubble")

        if is_xy:
            chart_data = XyChartData()
            for s in series_data:
                series_obj = chart_data.add_series(str(s.get("name", "Series")))
                data_points = s.get("data_points", [])
                for dp in data_points:
                    series_obj.add_data_point(float(dp["x"]), float(dp["y"]))
        elif is_bubble:
            chart_data = BubbleChartData()
            for s in series_data:
                series_obj = chart_data.add_series(str(s.get("name", "Series")))
                data_points = s.get("data_points", [])
                for dp in data_points:
                    series_obj.add_data_point(float(dp["x"]), float(dp["y"]), float(dp.get("size", 10)))
        else:
            chart_data = CategoryChartData()
            categories = params.get("categories", [])
            if categories:
                chart_data.categories = categories
            for s in series_data:
                chart_data.add_series(str(s.get("name", "Series")), s.get("values", []))

        graphic_frame = slide.shapes.add_chart(xl_type, Emu(left), Emu(top), Emu(width), Emu(height), chart_data)
        chart = graphic_frame.chart

        # Optional styling
        if params.get("has_legend") is not None:
            chart.has_legend = bool(params["has_legend"])

        if chart.has_legend and params.get("legend_position"):
            pos = self._LEGEND_POSITION_MAP.get(str(params["legend_position"]).lower())
            if pos is not None:
                chart.legend.position = pos
                chart.legend.include_in_layout = False

        if params.get("has_data_labels"):
            plot = chart.plots[0]
            plot.has_data_labels = True
            fmt = params.get("data_label_number_format")
            if fmt:
                plot.data_labels.number_format = str(fmt)

        if params.get("chart_style") is not None:
            chart.style = int(params["chart_style"])

        if params.get("title"):
            chart.has_title = True
            chart.chart_title.text_frame.text = str(params["title"])

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": graphic_frame.name,
            "shape_id": int(graphic_frame.shape_id),
            "chart_type": chart_type_str,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def get_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        ensure(shape.has_chart, "conflict", "Shape does not contain a chart.")
        chart = shape.chart

        # Determine chart type name
        chart_type_name = "unknown"
        for name, xl_type in self._CHART_TYPE_MAP.items():
            if chart.chart_type == xl_type:
                chart_type_name = name
                break

        # Extract categories
        categories: list[str] = []
        try:
            plot = chart.plots[0]
            for cat in plot.categories:
                categories.append(str(cat) if cat is not None else "")
        except Exception as exc:
            logger.warning("could not extract chart categories: %s", exc)

        # Extract series
        series_list: list[dict[str, Any]] = []
        try:
            for series in chart.series:
                s_data: dict[str, Any] = {"name": str(series.name) if hasattr(series, "name") else ""}
                # Try to get values
                try:
                    s_data["values"] = [float(v) if v is not None else None for v in series.values]
                except Exception as exc:
                    s_data["values"] = []
                    logger.warning("could not extract series values: %s", exc)
                series_list.append(s_data)
        except Exception as exc:
            logger.warning("could not extract chart series: %s", exc)

        # Chart properties
        result: dict[str, Any] = {
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "chart_type": chart_type_name,
            "categories": categories,
            "series": series_list,
            "has_legend": bool(chart.has_legend),
            "has_title": bool(chart.has_title),
        }

        if chart.has_title:
            try:
                result["title"] = chart.chart_title.text_frame.text
            except Exception:
                result["title"] = ""

        # Data labels
        try:
            plot = chart.plots[0]
            result["has_data_labels"] = bool(plot.has_data_labels)
            if plot.has_data_labels:
                result["data_label_number_format"] = str(plot.data_labels.number_format)
        except Exception:
            result["has_data_labels"] = False

        return result

    def update_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:
        from pptx.chart.data import CategoryChartData, XyChartData

        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        ensure(shape.has_chart, "conflict", "Shape does not contain a chart.")
        chart = shape.chart

        # Determine if XY/scatter type
        chart_type_name = "unknown"
        for name, xl_type in self._CHART_TYPE_MAP.items():
            if chart.chart_type == xl_type:
                chart_type_name = name
                break

        is_xy = chart_type_name.startswith("xy_scatter")
        series_data = params.get("series", [])

        if is_xy:
            new_data = XyChartData()
            for s in series_data:
                series_obj = new_data.add_series(str(s.get("name", "Series")))
                for dp in s.get("data_points", []):
                    series_obj.add_data_point(float(dp["x"]), float(dp["y"]))
        else:
            new_data = CategoryChartData()
            categories = params.get("categories")
            if categories:
                new_data.categories = categories
            else:
                # Keep existing categories
                try:
                    existing_cats = [str(c) for c in chart.plots[0].categories]
                    new_data.categories = existing_cats
                except Exception:
                    pass

            for s in series_data:
                new_data.add_series(str(s.get("name", "Series")), s.get("values", []))

        chart.replace_data(new_data)
        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "chart_type": chart_type_name,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def set_chart_style(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_index = int(params["slide_index"])
        slide = self._get_slide_obj(session, slide_index)

        shape = self._require_shape(slide, params)

        ensure(shape.has_chart, "conflict", "Shape does not contain a chart.")
        chart = shape.chart

        # Legend
        if "has_legend" in params:
            chart.has_legend = bool(params["has_legend"])

        if chart.has_legend and params.get("legend_position"):
            pos = self._LEGEND_POSITION_MAP.get(str(params["legend_position"]).lower())
            if pos is not None:
                chart.legend.position = pos
                chart.legend.include_in_layout = bool(params.get("legend_in_layout", False))

        # Data labels
        if "has_data_labels" in params:
            plot = chart.plots[0]
            plot.has_data_labels = bool(params["has_data_labels"])
            if plot.has_data_labels:
                fmt = params.get("data_label_number_format")
                if fmt:
                    plot.data_labels.number_format = str(fmt)

                dl_position = params.get("data_label_position")
                if dl_position:
                    from pptx.enum.chart import XL_LABEL_POSITION

                    dl_pos_map = {
                        "center": XL_LABEL_POSITION.CENTER,
                        "inside_end": XL_LABEL_POSITION.INSIDE_END,
                        "outside_end": XL_LABEL_POSITION.OUTSIDE_END,
                        "inside_base": XL_LABEL_POSITION.INSIDE_BASE,
                        "above": XL_LABEL_POSITION.ABOVE,
                        "below": XL_LABEL_POSITION.BELOW,
                        "left": XL_LABEL_POSITION.LEFT,
                        "right": XL_LABEL_POSITION.RIGHT,
                        "best_fit": XL_LABEL_POSITION.BEST_FIT,
                    }
                    pos = dl_pos_map.get(str(dl_position).lower())
                    if pos is not None:
                        plot.data_labels.position = pos

        # Chart style
        if "chart_style" in params:
            chart.style = int(params["chart_style"])

        # Title
        if "title" in params:
            title_text = params["title"]
            if title_text is None or title_text == "":
                chart.has_title = False
            else:
                chart.has_title = True
                chart.chart_title.text_frame.text = str(title_text)

        self._mark_modified(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "slide_index": slide_index,
            "shape_name": shape.name,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
