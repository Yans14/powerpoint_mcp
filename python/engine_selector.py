from __future__ import annotations

import platform
import sys
from typing import Any


def detect_engine() -> tuple[str, dict[str, Any]]:
    metadata: dict[str, Any] = {
        "python_version": sys.version.split()[0],
        "platform": platform.platform(),
        "preferred_engine": "OOXML",
        "powerpoint_version": "",
    }

    if platform.system() != "Windows":
        metadata["reason"] = "Non-Windows platform. Using OOXML engine."
        return "OOXML", metadata

    try:
        import win32com.client  # type: ignore[import]

        app = win32com.client.Dispatch("PowerPoint.Application")
        metadata["preferred_engine"] = "COM"
        metadata["powerpoint_version"] = str(getattr(app, "Version", ""))
        metadata["reason"] = "PowerPoint COM automation is available."
        app.Quit()
        return "COM", metadata
    except Exception as exc:  # pragma: no cover - platform-specific
        metadata["reason"] = f"COM unavailable: {exc}. Falling back to OOXML."
        return "OOXML", metadata
