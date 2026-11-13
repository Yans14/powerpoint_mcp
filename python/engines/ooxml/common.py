"""Module-level helpers and constants shared by the OOXMLEngine mixins."""

from __future__ import annotations

import logging

logger = logging.getLogger("pptx_mcp.engine.ooxml")


_DEFAULT_RENDER_TIMEOUT_SEC = 90.0


_COLOR_NAMESPACES = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
