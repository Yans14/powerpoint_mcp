from __future__ import annotations

import pytest

from engines.com_engine import COMEngine
from engines.ooxml_engine import OOXMLEngine
from orchestrator.capability_manifest import CAPABILITY_MANIFEST
from service import PowerPointService
from tool_catalog import BATCH_TOOL, TOOL_CATALOG, TOOLS, TOOLS_BY_NAME, engine_method_for, is_engine_tool

ENGINE_TOOLS = [tool for tool in TOOLS if is_engine_tool(tool)]


@pytest.mark.parametrize("engine_class", [OOXMLEngine, COMEngine])
def test_every_engine_tool_has_an_engine_method(engine_class):
    missing = [
        tool["name"] for tool in ENGINE_TOOLS if not callable(getattr(engine_class, engine_method_for(tool), None))
    ]

    assert missing == []


def test_service_routes_every_tool():
    service = PowerPointService()
    try:
        assert set(service.method_map) | {BATCH_TOOL} == set(TOOLS_BY_NAME)
        assert service.method_map["pptx_check_positions"] == "checker_dispatch"
        assert service.method_map["pptx_agent_start"] == "agent_dispatch"
        assert service.method_map["pptx_add_slide"] == "add_slide"
    finally:
        service.shutdown()


def _ui_params(tool_name):
    tool = next(item for item in TOOL_CATALOG if item["name"] == tool_name)
    return {param["name"]: param for param in tool["params"]}


def test_web_ui_form_description_is_derived_from_the_schemas():
    connector = _ui_params("pptx_add_connector")
    create = _ui_params("pptx_create_presentation")
    image = _ui_params("pptx_add_image")
    copy = _ui_params("pptx_copy_shape_between_decks")

    assert connector["presentation_id"] == {
        "name": "presentation_id",
        "type": "uuid",
        "required": True,
        "description": connector["presentation_id"]["description"],
        "auto_fill": "presentation_id",
    }
    assert connector["slide_index"]["type"] == "number"
    assert connector["connector_type"]["type"] == "enum"
    assert "straight" in connector["connector_type"]["enum_values"]
    assert connector["connector_type"]["required"] is False
    assert create["width"]["type"] == "measurement"
    assert image["image_path"]["type"] == "path"
    assert copy["target_presentation_id"]["auto_fill"] == "presentation_id"


def test_web_ui_catalog_groups_every_tool_once():
    from tool_catalog import get_catalog_by_category

    groups = get_catalog_by_category()

    assert sorted(tool["name"] for tools in groups.values() for tool in tools) == sorted(TOOLS_BY_NAME)
    assert "Session" in groups and "Checkers" in groups


def test_agent_manifest_lists_engine_tools_with_their_real_parameters():
    for tool in ENGINE_TOOLS:
        assert f"- {tool['name']}(" in CAPABILITY_MANIFEST

    assert "pptx_batch(" not in CAPABILITY_MANIFEST
    assert "pptx_agent_start(" not in CAPABILITY_MANIFEST
    assert "pptx_check_positions(" not in CAPABILITY_MANIFEST
    # The hand-written manifest drifted (e.g. "text" for pptx_add_shape, which takes text_content).
    for tool in ENGINE_TOOLS:
        line = next(item for item in CAPABILITY_MANIFEST.splitlines() if item.startswith(f"- {tool['name']}("))
        for name in tool["input_schema"].get("properties", {}):
            assert f"{name}" in line, (tool["name"], name)
