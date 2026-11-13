"""Parameter validation for bridge methods.

The TypeScript MCP layer validates tool input with zod, but the web UI and the agent plan
executor call the Python service directly. Without a check here, a missing or mistyped parameter
surfaces as a raw ``KeyError``/``ValueError`` from deep inside an engine, and a misspelt optional
parameter is silently ignored. This module validates against the same JSON Schema the MCP client
sees (``tool_catalog.INPUT_SCHEMAS``) and rejects bad calls with a ``validation_error`` that names
the offending parameters.

``None`` counts as "not provided": such keys are dropped before validation and before the engine
sees them.
"""

from __future__ import annotations

from typing import Any

from jsonschema import Draft7Validator

from errors import BridgeError
from tool_catalog import INPUT_SCHEMAS

_VALIDATORS: dict[str, Draft7Validator] = {name: Draft7Validator(schema) for name, schema in INPUT_SCHEMAS.items()}

_MAX_PROBLEM_LENGTH = 200


def validate_params(method: str, params: Any) -> dict[str, Any]:
    """Return ``params`` without null values, or raise ``BridgeError(validation_error)``.

    Methods absent from the catalog are not checked.
    """
    if not isinstance(params, dict):
        raise BridgeError(
            code="validation_error",
            message=f"Invalid parameters for {method}: params must be an object.",
        )

    cleaned = {key: value for key, value in params.items() if value is not None}

    validator = _VALIDATORS.get(method)
    if validator is None:
        return cleaned

    problems: list[dict[str, str]] = []
    for error in sorted(validator.iter_errors(cleaned), key=lambda item: list(item.absolute_path)):
        problems.extend(_describe(error, cleaned, validator.schema))

    if problems:
        summary = "; ".join(f"{item['param']}: {item['problem']}" for item in problems)
        raise BridgeError(
            code="validation_error",
            message=f"Invalid parameters for {method}. {summary}",
            details={"method": method, "errors": problems},
        )
    return cleaned


def _describe(error: Any, params: dict[str, Any], schema: dict[str, Any]) -> list[dict[str, str]]:
    path = [str(part) for part in error.absolute_path]

    if not path and error.validator == "required":
        return [
            {"param": name, "problem": "Required parameter is missing."}
            for name in error.validator_value
            if name not in params
        ]

    if not path and error.validator == "additionalProperties":
        known = set(schema.get("properties", {}))
        return [
            {"param": name, "problem": f"Unknown parameter. Valid parameters: {', '.join(sorted(known))}."}
            for name in params
            if name not in known
        ]

    param = ".".join(path) if path else "(params)"
    problem = error.message
    if error.validator == "anyOf":
        problem = f"{_short(error.instance)} does not match the expected format."
        description = error.schema.get("description")
        if description:
            problem += f" {description}"
    elif error.validator == "enum":
        problem = f"Must be one of: {', '.join(map(str, error.validator_value))}."
    return [{"param": param, "problem": _truncate(problem)}]


def _short(value: Any) -> str:
    return _truncate(repr(value), 60)


def _truncate(text: str, limit: int = _MAX_PROBLEM_LENGTH) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "…"
