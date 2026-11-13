"""Shapes, images, lines, connectors, grouping and z-order."""

from __future__ import annotations

from typing import Any

from engines.com.common import (
    _MSO_CONNECTOR_MAP,
    _MSO_SHAPE_MAP,
    _ZORDER_MAP,
    _hex_to_bgr_int,
    _to_points,
)
from errors import ensure
from utils.paths import validate_existing_file


class COMShapesMixin:
    """Shapes, images, lines, connectors, grouping and z-order."""

    def get_shape_details(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        details = self.worker.call(self._get_shape_details_ts, pid, si, params)
        return {"presentation_id": pid, "slide_index": si, **details}

    def _get_shape_details_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        return self._detailed_shape_payload_com(shape)

    def add_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_shape_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_shape_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape_type_str = str(params["shape_type"]).lower()
        ensure(shape_type_str in _MSO_SHAPE_MAP, "validation_error", f"Unknown shape type: {shape_type_str}")
        mso_type = _MSO_SHAPE_MAP[shape_type_str]

        left = _to_points(params["left"])
        top = _to_points(params["top"])
        width = _to_points(params["width"])
        height = _to_points(params["height"])

        shape = slide.Shapes.AddShape(mso_type, left, top, width, height)

        if params.get("fill_hex"):
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = _hex_to_bgr_int(str(params["fill_hex"]))
        if params.get("line_hex"):
            shape.Line.ForeColor.RGB = _hex_to_bgr_int(str(params["line_hex"]))
        if params.get("text"):
            shape.TextFrame.TextRange.Text = str(params["text"])

        return {"shape_name": str(shape.Name), "shape_id": int(shape.Id)}

    def add_image(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        image_path = validate_existing_file(
            str(params["image_path"]),
            expected_suffixes=(".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tiff", ".svg"),
        )
        result = self.worker.call(self._add_image_ts, pid, si, str(image_path), params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_image_ts(self, sid: str, si: int, image_path: str, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        left = _to_points(params.get("left", "0in"))
        top = _to_points(params.get("top", "0in"))

        has_w = params.get("width") is not None
        has_h = params.get("height") is not None

        if has_w and has_h:
            width = _to_points(params["width"])
            height = _to_points(params["height"])
            shape = slide.Shapes.AddPicture(image_path, False, True, left, top, width, height)
        else:
            shape = slide.Shapes.AddPicture(image_path, False, True, left, top, -1, -1)
            shape.ScaleWidth(1, True)
            shape.ScaleHeight(1, True)
            if has_w:
                shape.Width = _to_points(params["width"])
            if has_h:
                shape.Height = _to_points(params["height"])

        return {
            "shape_name": str(shape.Name),
            "shape_id": int(shape.Id),
            "left_inches": float(shape.Left) / 72.0,
            "top_inches": float(shape.Top) / 72.0,
            "width_inches": float(shape.Width) / 72.0,
            "height_inches": float(shape.Height) / 72.0,
        }

    def add_line(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_line_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_line_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        x1 = _to_points(params["begin_x"])
        y1 = _to_points(params["begin_y"])
        x2 = _to_points(params["end_x"])
        y2 = _to_points(params["end_y"])

        shape = slide.Shapes.AddLine(x1, y1, x2, y2)

        color_hex = params.get("color_hex", "000000")
        shape.Line.ForeColor.RGB = _hex_to_bgr_int(str(color_hex))
        width_pt = float(params.get("width_pt", 1.0))
        shape.Line.Weight = width_pt

        dash_style = params.get("dash_style")
        if dash_style:
            dash_map = {"solid": 1, "dash": 4, "dot": 3, "dash_dot": 5, "long_dash": 7}
            if dash_style.lower() in dash_map:
                shape.Line.DashStyle = dash_map[dash_style.lower()]

        line_name = str(params.get("line_name", "Connector"))
        try:
            shape.Name = line_name
        except Exception:
            pass

        return {
            "line_name": str(shape.Name),
            "begin": {"x_inches": float(x1) / 72.0, "y_inches": float(y1) / 72.0},
            "end": {"x_inches": float(x2) / 72.0, "y_inches": float(y2) / 72.0},
        }

    def delete_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        deleted_name = self.worker.call(self._delete_shape_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "deleted_shape": deleted_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _delete_shape_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        name = str(shape.Name)
        shape.Delete()
        return name

    def set_shape_properties(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._set_shape_properties_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_shape_properties_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)

        if params.get("left") is not None:
            shape.Left = _to_points(params["left"])
        if params.get("top") is not None:
            shape.Top = _to_points(params["top"])
        if params.get("width") is not None:
            shape.Width = _to_points(params["width"])
        if params.get("height") is not None:
            shape.Height = _to_points(params["height"])
        if params.get("rotation") is not None:
            shape.Rotation = float(params["rotation"])

        fill_hex = params.get("fill_hex")
        if fill_hex is not None:
            if str(fill_hex).lower() == "none":
                shape.Fill.Visible = 0
            else:
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = _hex_to_bgr_int(str(fill_hex))

        line_hex = params.get("line_hex")
        if line_hex is not None:
            if str(line_hex).lower() == "none":
                shape.Line.Visible = 0
            else:
                shape.Line.Visible = -1
                shape.Line.ForeColor.RGB = _hex_to_bgr_int(str(line_hex))

        if params.get("line_width_pt") is not None:
            shape.Line.Weight = float(params["line_width_pt"])
        if params.get("name"):
            shape.Name = str(params["name"])

        return {
            "shape_name": str(shape.Name),
            "left_inches": float(shape.Left) / 72.0,
            "top_inches": float(shape.Top) / 72.0,
            "width_inches": float(shape.Width) / 72.0,
            "height_inches": float(shape.Height) / 72.0,
            "rotation": float(shape.Rotation),
        }

    def clone_shape(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        target_si = int(params.get("target_slide_index", si))
        result = self.worker.call(self._clone_shape_ts, pid, si, target_si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "source_slide_index": si,
            "target_slide_index": target_si,
            "cloned_shape_name": result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _clone_shape_ts(self, sid: str, si: int, target_si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)

        offset_l = _to_points(params["offset_left"]) if params.get("offset_left") else 18  # 0.25in default
        offset_t = _to_points(params["offset_top"]) if params.get("offset_top") else 18

        if si == target_si:
            dup = shape.Duplicate()(1)
            dup.Left = shape.Left + offset_l
            dup.Top = shape.Top + offset_t
            return str(dup.Name)
        else:
            target_slide = self._slide(sid, target_si)
            shape.Copy()
            pasted = target_slide.Shapes.Paste()(1)
            pasted.Left = shape.Left + offset_l
            pasted.Top = shape.Top + offset_t
            return str(pasted.Name)

    def set_shape_z_order(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        action = str(params.get("action", "front")).lower()
        self.worker.call(self._set_shape_z_order_ts, pid, si, action, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "shape_name": str(params.get("shape_name", params.get("shape_id", ""))),
            "action": action,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_shape_z_order_ts(self, sid: str, si: int, action: str, params: dict[str, Any]) -> None:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(action in _ZORDER_MAP, "validation_error", f"Unknown z-order action: {action}")
        shape.ZOrder(_ZORDER_MAP[action])

    def group_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._group_shapes_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _group_shapes_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        shape_names = params.get("shape_names")
        shape_ids = params.get("shape_ids")
        ensure(shape_names or shape_ids, "validation_error", "Provide shape_names or shape_ids")

        names: list[str] = []
        if shape_names:
            names = [str(n) for n in shape_names]
        elif shape_ids:
            for sid_val in shape_ids:
                shape = self._find_shape_com(slide, shape_id=int(sid_val))
                names.append(str(shape.Name))

        ensure(len(names) >= 2, "validation_error", "At least 2 shapes required for grouping")

        shape_range = slide.Shapes.Range(names)
        group = shape_range.Group()
        group_name = str(params.get("group_name", "Group"))
        try:
            group.Name = group_name
        except Exception:
            group_name = str(group.Name)
        return {"group_name": group_name, "shapes_grouped": len(names)}

    def ungroup_shapes(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        ungrouped = self.worker.call(self._ungroup_shapes_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "presentation_id": pid,
            "slide_index": si,
            "ungrouped_shapes": ungrouped,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _ungroup_shapes_ts(self, sid: str, si: int, params: dict[str, Any]) -> list[str]:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        ensure(int(shape.Type) == 6, "conflict", "Shape is not a group")
        shape_range = shape.Ungroup()
        return [str(shape_range(i).Name) for i in range(1, shape_range.Count + 1)]

    def copy_shape_between_decks(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        src_pid = str(params["source_presentation_id"])
        tgt_pid = str(params["target_presentation_id"])
        src_si = int(params["source_slide_index"])
        tgt_si = int(params["target_slide_index"])
        result = self.worker.call(self._copy_shape_between_decks_ts, src_pid, tgt_pid, src_si, tgt_si, params)
        self._mark_dirty(tgt_pid)
        return {
            "success": True,
            "source_presentation_id": src_pid,
            "target_presentation_id": tgt_pid,
            "source_slide_index": src_si,
            "target_slide_index": tgt_si,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": tgt_pid}),
        }

    def _copy_shape_between_decks_ts(
        self, src_sid: str, tgt_sid: str, src_si: int, tgt_si: int, params: dict[str, Any]
    ) -> dict[str, Any]:
        src_slide = self._slide(src_sid, src_si)
        tgt_slide = self._slide(tgt_sid, tgt_si)
        shape = self._require_shape_com(src_slide, params)
        src_name = str(shape.Name)

        shape.Copy()
        pasted = tgt_slide.Shapes.Paste()(1)

        if params.get("offset_left"):
            pasted.Left = pasted.Left + _to_points(params["offset_left"])
        if params.get("offset_top"):
            pasted.Top = pasted.Top + _to_points(params["offset_top"])

        return {"source_shape_name": src_name, "new_shape_name": str(pasted.Name)}

    def set_shape_fill_gradient(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        shape_name = self.worker.call(self._set_shape_fill_gradient_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            "shape_name": shape_name,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _set_shape_fill_gradient_ts(self, sid: str, si: int, params: dict[str, Any]) -> str:
        slide = self._slide(sid, si)
        shape = self._require_shape_com(slide, params)
        fill = shape.Fill
        stops = params.get("stops", [])
        angle = params.get("angle")

        if len(stops) == 2:
            fill.TwoColorGradient(1, 1)  # msoGradientHorizontal, variant 1
            fill.ForeColor.RGB = _hex_to_bgr_int(str(stops[0]["color_hex"]))
            fill.BackColor.RGB = _hex_to_bgr_int(str(stops[1]["color_hex"]))
        elif len(stops) > 2:
            fill.TwoColorGradient(1, 1)
            fill.ForeColor.RGB = _hex_to_bgr_int(str(stops[0]["color_hex"]))
            fill.BackColor.RGB = _hex_to_bgr_int(str(stops[-1]["color_hex"]))
            try:
                gs = fill.GradientStops
                while gs.Count > 2:
                    gs.Delete(gs.Count)
                for stop in stops[1:-1]:
                    gs.Insert(
                        _hex_to_bgr_int(str(stop["color_hex"])),
                        float(stop["position"]),
                    )
            except Exception:
                pass  # Older Office may not support GradientStops

        if angle is not None:
            try:
                fill.GradientAngle = float(angle)
            except Exception:
                pass  # Office 2013+ only

        return str(shape.Name)

    def add_connector(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover
        pid = str(params["presentation_id"])
        si = int(params["slide_index"])
        result = self.worker.call(self._add_connector_ts, pid, si, params)
        self._mark_dirty(pid)
        return {
            "success": True,
            **result,
            "presentation_state": self.get_presentation_state({"presentation_id": pid}),
        }

    def _add_connector_ts(self, sid: str, si: int, params: dict[str, Any]) -> dict[str, Any]:
        slide = self._slide(sid, si)
        conn_type_str = str(params.get("connector_type", "straight")).lower()
        conn_type = _MSO_CONNECTOR_MAP.get(conn_type_str, 1)

        # Create connector with temporary coordinates
        connector = slide.Shapes.AddConnector(conn_type, 0, 0, 100, 100)

        # Connect begin shape
        begin_name = params.get("begin_shape_name")
        begin_id = params.get("begin_shape_id")
        if begin_name or begin_id:
            begin_shape = self._find_shape_com(
                slide, shape_name=begin_name, shape_id=int(begin_id) if begin_id else None
            )
            site = int(params.get("begin_connection_site", 0)) + 1  # COM is 1-based
            connector.ConnectorFormat.BeginConnect(begin_shape, site)

        # Connect end shape
        end_name = params.get("end_shape_name")
        end_id = params.get("end_shape_id")
        if end_name or end_id:
            end_shape = self._find_shape_com(slide, shape_name=end_name, shape_id=int(end_id) if end_id else None)
            site = int(params.get("end_connection_site", 0)) + 1
            connector.ConnectorFormat.EndConnect(end_shape, site)

        if params.get("color_hex"):
            connector.Line.ForeColor.RGB = _hex_to_bgr_int(str(params["color_hex"]))
        if params.get("width_pt"):
            connector.Line.Weight = float(params["width_pt"])

        connector.RerouteConnections()

        return {"shape_name": str(connector.Name), "shape_id": int(connector.Id)}
