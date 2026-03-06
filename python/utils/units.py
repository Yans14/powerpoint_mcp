from __future__ import annotations

import re

from errors import BridgeError

EMU_PER_INCH = 914400
EMU_PER_POINT = 12700
EMU_PER_CM = 360000
EMU_PER_PIXEL = 9525

_MEASUREMENT_RE = re.compile(r"^(\d+(?:\.\d+)?)(in|pt|cm|px)$", re.IGNORECASE)


def to_emu(value: int | float | str) -> int:
    if isinstance(value, int | float):
        return int(round(value))

    if not isinstance(value, str):
        raise BridgeError(
            code="validation_error",
            message="Measurement must be number (EMU) or string with unit.",
            details={"value": value},
        )

    match = _MEASUREMENT_RE.match(value.strip())
    if not match:
        raise BridgeError(
            code="validation_error",
            message="Invalid measurement string.",
            details={"value": value, "expected": "2in, 24pt, 5cm, 96px"},
        )

    numeric = float(match.group(1))
    unit = match.group(2).lower()

    if unit == "in":
        return int(round(numeric * EMU_PER_INCH))
    if unit == "pt":
        return int(round(numeric * EMU_PER_POINT))
    if unit == "cm":
        return int(round(numeric * EMU_PER_CM))
    if unit == "px":
        return int(round(numeric * EMU_PER_PIXEL))

    raise BridgeError(
        code="validation_error",
        message="Unsupported measurement unit.",
        details={"unit": unit},
    )


def emu_to_inches(value: int | float) -> float:
    return float(value) / EMU_PER_INCH
