from __future__ import annotations

from dataclasses import dataclass
from typing import Any


@dataclass
class BridgeError(Exception):
    code: str
    message: str
    details: dict[str, Any] | None = None
    # One sentence telling the caller (usually a model) what to do next.
    hint: str | None = None

    def __str__(self) -> str:
        return f"{self.code}: {self.message}"

    def to_payload(self) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "code": self.code,
            "message": self.message,
        }
        if self.details:
            payload["details"] = self.details
        if self.hint:
            payload["hint"] = self.hint
        return payload


def ensure(condition: bool, code: str, message: str, details: dict[str, Any] | None = None) -> None:
    if not condition:
        raise BridgeError(code=code, message=message, details=details)


SESSION_NOT_FOUND_HINT = (
    "Call pptx_list_open_presentations to see valid presentation_id values, "
    "or open the file with pptx_open_presentation."
)

# Engines report lookups that failed by listing what does exist under one of these detail keys.
# Deriving the hint from the key keeps both engines consistent without editing every raise site.
_HINTS_BY_DETAIL_KEY = {
    "available_layouts": ("Use one of details.available_layouts (names are case-sensitive), or call pptx_get_layouts."),
    "available_placeholders": (
        "Use one of details.available_placeholders, or call pptx_get_placeholders for this slide."
    ),
    "available_shapes": (
        "Use a shape name or id from details.available_shapes, or call pptx_get_slide_shapes for this slide."
    ),
    "slide_count": (
        "Valid slide_index values run from 1 to details.slide_count; "
        "call pptx_get_presentation_state for the current slides."
    ),
}


def suggest_hint(error: BridgeError) -> str | None:
    """Next-step suggestion for an error that was raised without one."""
    details = error.details or {}
    for key, hint in _HINTS_BY_DETAIL_KEY.items():
        if key in details:
            return hint
    return None
