"""Module-level helpers and constants shared by the COMEngine mixins."""

from __future__ import annotations

from utils.colors import normalize_color
from utils.units import to_emu

try:  # pragma: no cover - Windows only
    import win32com.client  # type: ignore[import]
except Exception:  # pragma: no cover
    win32com = None  # type: ignore[assignment]


def _hex_to_bgr_int(color_value: str) -> int:
    color = normalize_color(color_value)
    rr = int(color[0:2], 16)
    gg = int(color[2:4], 16)
    bb = int(color[4:6], 16)
    return (bb << 16) | (gg << 8) | rr


def _to_points(value) -> float:
    """Convert any unit string ('2in', '24pt', EMU int) to PowerPoint points."""
    return to_emu(value) / 12700


# MSO AutoShapeType constants
_MSO_SHAPE_MAP: dict[str, int] = {
    "rectangle": 1,
    "rounded_rectangle": 5,
    "oval": 9,
    "diamond": 4,
    "triangle": 7,
    "right_arrow": 33,
    "left_arrow": 34,
    "up_arrow": 35,
    "down_arrow": 36,
    "pentagon": 56,
    "hexagon": 10,
    "chevron": 52,
    "star_5_point": 12,
    "line_inverse": 183,
    "cross": 11,
    "frame": 158,
    "rectangular_callout": 105,
    "rounded_rectangular_callout": 106,
    "cloud_callout": 108,
    "cloud": 179,
}


# MSO Connector type constants
_MSO_CONNECTOR_MAP: dict[str, int] = {
    "straight": 1,
    "elbow": 2,
    "curve": 3,
}


# Excel chart type constants (for COM AddChart2 / AddChart)
_XL_CHART_TYPE_MAP: dict[str, int] = {
    "column_clustered": 51,
    "column_stacked": 52,
    "column_stacked_100": 53,
    "bar_clustered": 57,
    "bar_stacked": 58,
    "bar_stacked_100": 59,
    "line": 4,
    "line_markers": 65,
    "line_stacked": 63,
    "pie": 5,
    "pie_exploded": 69,
    "doughnut": -4120,
    "area": 1,
    "area_stacked": 76,
    "area_stacked_100": 77,
    "xy_scatter": -4169,
    "xy_scatter_lines": 74,
    "xy_scatter_smooth": 72,
    "bubble": 15,
    "radar": -4151,
    "stock_hlc": 88,
    "stock_ohlc": 89,
    "three_d_column": 54,
    "three_d_bar_clustered": 60,
    "three_d_pie": -4102,
    "three_d_line": -4101,
}


_XL_CHART_TYPE_REVERSE: dict[int, str] = {v: k for k, v in _XL_CHART_TYPE_MAP.items()}


# Alignment map (ppAlign constants)
_ALIGN_MAP: dict[str, int] = {
    "left": 1,
    "center": 2,
    "right": 3,
    "justify": 4,
}


# Z-order commands (MsoZOrderCmd)
_ZORDER_MAP: dict[str, int] = {
    "front": 0,  # msoBringToFront
    "back": 1,  # msoSendToBack
    "forward": 2,  # msoBringForward
    "backward": 3,  # msoSendBackward
}


# COM legend position constants
_LEGEND_POS_MAP: dict[str, int] = {
    "bottom": -4107,
    "corner": 2,
    "left": -4131,
    "right": -4152,
    "top": -4160,
}
