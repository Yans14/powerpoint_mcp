from __future__ import annotations

import pytest

from errors import BridgeError
from service import PowerPointService


@pytest.fixture
def service():
    svc = PowerPointService()
    yield svc
    svc.shutdown()


@pytest.fixture
def deck(service):
    pid = service.dispatch("pptx_create_presentation", {})["presentation_id"]
    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})
    return pid


def _error(service, method, params):
    with pytest.raises(BridgeError) as excinfo:
        service.dispatch(method, params)
    return excinfo.value


# --- Response size ------------------------------------------------------------------------------


def test_mutations_return_a_slide_count_instead_of_the_full_state(service, deck):
    result = service.dispatch("pptx_add_slide", {"presentation_id": deck, "layout_name": "Blank"})

    assert "presentation_state" not in result
    assert result["slide_count"] == 2
    assert result["success"] is True


def test_include_state_returns_the_full_state(service, deck):
    result = service.dispatch(
        "pptx_add_slide", {"presentation_id": deck, "layout_name": "Blank", "include_state": True}
    )

    assert result["presentation_state"]["slide_count"] == 2
    assert [slide["layout"] for slide in result["presentation_state"]["slides"]] == ["Title Slide", "Blank"]


def test_include_state_false_and_non_boolean_values_keep_the_compact_form(service, deck):
    for value in (False, "true", 1):
        result = service.dispatch(
            "pptx_add_slide", {"presentation_id": deck, "layout_name": "Blank", "include_state": value}
        )
        assert "presentation_state" not in result


def test_state_reads_are_not_compacted(service, deck):
    state = service.dispatch("pptx_get_presentation_state", {"presentation_id": deck})

    assert state["slides"][0]["layout"] == "Title Slide"


# --- Error hints --------------------------------------------------------------------------------


def test_unknown_layout_lists_valid_names_and_points_to_the_lookup_tool(service, deck):
    error = _error(service, "pptx_add_slide", {"presentation_id": deck, "layout_name": "Nope"})

    assert error.code == "not_found"
    assert "Title Slide" in error.details["available_layouts"]
    assert "pptx_get_layouts" in error.hint
    assert error.to_payload()["hint"] == error.hint


def test_unknown_placeholder_and_shape_get_a_hint(service, deck):
    placeholder = _error(
        service,
        "pptx_set_placeholder_text",
        {"presentation_id": deck, "slide_index": 1, "placeholder_name": "Nope", "text_content": "x"},
    )
    shape = _error(service, "pptx_delete_shape", {"presentation_id": deck, "slide_index": 1, "shape_name": "Nope"})

    assert placeholder.code == "not_found" and "pptx_get_placeholders" in placeholder.hint
    assert shape.code == "not_found" and "pptx_get_slide_shapes" in shape.hint


def test_out_of_range_slide_index_gets_a_hint(service, deck):
    error = _error(service, "pptx_get_slide", {"presentation_id": deck, "slide_index": 7})

    assert error.code == "validation_error"
    assert error.details["slide_count"] == 1
    assert "pptx_get_presentation_state" in error.hint


def test_unknown_session_gets_a_hint(service):
    error = _error(service, "pptx_get_presentation_state", {"presentation_id": "missing"})

    assert error.code == "not_found"
    assert "pptx_list_open_presentations" in error.hint


def test_errors_without_a_known_remedy_carry_no_hint(service):
    error = _error(service, "pptx_no_such_method", {})

    assert error.hint is None
    assert "hint" not in error.to_payload()
