from __future__ import annotations

from errors import BridgeError


def normalize_color(color_value: str) -> str:
    """Normalize a hex color string to uppercase 6-char format.

    Accepts '#RRGGBB' or 'RRGGBB'. Raises BridgeError on invalid input.
    """
    value = color_value.strip()
    if value.startswith("#"):
        value = value[1:]
    if len(value) != 6:
        raise BridgeError(
            code="validation_error",
            message="Color must be 6 hex characters.",
            details={"color": color_value},
        )
    try:
        int(value, 16)
    except ValueError as exc:
        raise BridgeError(
            code="validation_error",
            message="Color must be valid hex.",
            details={"color": color_value},
        ) from exc
    return value.upper()
