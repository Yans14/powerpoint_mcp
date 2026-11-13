"""Parameter validation for bridge methods.

The TypeScript MCP layer validates tool input with zod, but the web UI and the agent plan
executor call the Python service directly. Without a check here, a missing or mistyped
parameter surfaces as a raw ``KeyError``/``ValueError`` from deep inside an engine and is
reported as an internal error. This module rejects such calls up front with a
``validation_error`` that names the offending parameters.

Checks are driven by ``tool_catalog.TOOL_CATALOG`` and are deliberately lenient: parameters the
catalog does not list are passed through untouched, and ``None`` counts as "not provided".
"""

from __future__ import annotations

import math
from collections.abc import Callable
from typing import Any

from errors import BridgeError
from tool_catalog import TOOL_CATALOG


def _is_number(value: Any) -> bool:
    return isinstance(value, int | float) and not isinstance(value, bool) and math.isfinite(value)


def _is_string(value: Any) -> bool:
    return isinstance(value, str)


_TYPE_CHECKS: dict[str, tuple[Callable[[Any], bool], str]] = {
    "string": (_is_string, "a string"),
    "uuid": (_is_string, "a string"),
    "path": (_is_string, "a string"),
    "enum": (_is_string, "a string"),
    "number": (_is_number, "a number"),
    "boolean": (lambda value: isinstance(value, bool), "a boolean"),
    "array": (lambda value: isinstance(value, list), "an array"),
    "object": (lambda value: isinstance(value, dict), "an object"),
    "measurement": (lambda value: _is_number(value) or _is_string(value), 'a number or a string like "2in"'),
}

_PARAM_SPECS: dict[str, list[dict[str, Any]]] = {tool["name"]: tool["params"] for tool in TOOL_CATALOG}


def validate_params(method: str, params: Any) -> None:
    """Raise ``BridgeError(validation_error)`` if ``params`` cannot be a valid call to ``method``.

    Methods absent from the catalog are not checked.
    """
    specs = _PARAM_SPECS.get(method)
    if specs is None:
        return

    if not isinstance(params, dict):
        raise BridgeError(
            code="validation_error",
            message=f"Invalid parameters for {method}: params must be an object.",
        )

    problems: list[dict[str, str]] = []
    for spec in specs:
        name = spec["name"]
        value = params.get(name)

        if value is None:
            if spec.get("required"):
                problems.append({"param": name, "problem": "Required parameter is missing."})
            continue

        check = _TYPE_CHECKS.get(spec.get("type", ""))
        if check is None:
            continue

        is_valid, expected = check
        if not is_valid(value):
            problems.append({"param": name, "problem": f"Expected {expected}, got {type(value).__name__}."})
            continue

        allowed = spec.get("enum_values")
        if allowed and value not in allowed:
            problems.append({"param": name, "problem": f"Must be one of: {', '.join(map(str, allowed))}."})

    if problems:
        summary = "; ".join(f"{item['param']}: {item['problem']}" for item in problems)
        raise BridgeError(
            code="validation_error",
            message=f"Invalid parameters for {method}. {summary}",
            details={"method": method, "errors": problems},
        )
