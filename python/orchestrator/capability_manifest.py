"""Tool list given to the agent's planner and clarifier.

Generated from the tool catalog rather than written by hand, so parameter names always match
what the executor validates against. Only tools the plan executor can call are listed.
"""

from __future__ import annotations

from typing import Any

from tool_catalog import TOOLS, is_engine_tool

_CONVENTIONS = """=== Parameter conventions ===
- presentation_id: UUID string from pptx_open_presentation or pptx_create_presentation
- slide_index: 1-based integer
- Measurements: "2in", "24pt", "5cm", or raw EMU integer
- Colors: "#RRGGBB" or "RRGGBB"
- shape_name OR shape_id is required for shape operations (not both needed)
- Optional parameters are marked with "?"; do not invent parameters that are not listed
- Use $variable_name to reference values produced by earlier plan steps"""


def _type_label(schema: dict[str, Any]) -> str:
    if "enum" in schema:
        return "|".join(map(str, schema["enum"]))
    variants = schema.get("anyOf")
    if variants:
        kinds = {variant.get("type") for variant in variants}
        if {"number", "string"} <= kinds:
            return "measurement"
        return "string"
    kind = schema.get("type", "any")
    if kind == "array":
        return f"{_type_label(schema.get('items', {}))}[]"
    return {"integer": "int", "number": "number", "boolean": "bool", "string": "str"}.get(kind, kind)


def _signature(tool: dict[str, Any]) -> str:
    schema = tool["input_schema"]
    required = set(schema.get("required", []))
    params = [
        f"{name}{'' if name in required else '?'}: {_type_label(prop)}"
        for name, prop in schema.get("properties", {}).items()
    ]
    return f"- {tool['name']}({', '.join(params)}): {tool['description']}"


def build_capability_manifest() -> str:
    sections: dict[str, list[str]] = {}
    for tool in TOOLS:
        if is_engine_tool(tool):
            sections.setdefault(tool["category"].upper(), []).append(_signature(tool))

    lines = ["=== Available PowerPoint Engine Tools ==="]
    for title, signatures in sections.items():
        lines += ["", f"{title}:", *signatures]
    lines += ["", _CONVENTIONS]
    return "\n".join(lines) + "\n"


CAPABILITY_MANIFEST = build_capability_manifest()
