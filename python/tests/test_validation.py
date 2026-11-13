from __future__ import annotations

import pytest

from errors import BridgeError
from service import PowerPointService
from validation import validate_params

PID = "00000000-0000-4000-8000-000000000000"


def _errors(method, params):
    with pytest.raises(BridgeError) as excinfo:
        validate_params(method, params)
    assert excinfo.value.code == "validation_error"
    return {item["param"]: item["problem"] for item in excinfo.value.details["errors"]}


def test_missing_required_parameter_is_reported_by_name():
    errors = _errors("pptx_add_slide", {"presentation_id": PID})

    assert list(errors) == ["layout_name"]


def test_all_problems_are_reported_together():
    errors = _errors("pptx_get_slide", {"slide_index": "abc"})

    assert set(errors) == {"presentation_id", "slide_index"}


@pytest.mark.parametrize(
    ("method", "params", "bad_param"),
    [
        ("pptx_get_slide", {"presentation_id": PID, "slide_index": "1"}, "slide_index"),
        ("pptx_get_slide", {"presentation_id": PID, "slide_index": True}, "slide_index"),
        ("pptx_get_slide", {"presentation_id": 42, "slide_index": 1}, "presentation_id"),
        ("pptx_reorder_slides", {"presentation_id": PID, "new_order": "2,1"}, "new_order"),
        ("pptx_create_presentation", {"width": ["10in"]}, "width"),
        ("pptx_find_replace_text", {"presentation_id": PID, "find": "a", "replace": "b", "case_sensitive": 1}, None),
    ],
)
def test_wrong_types_are_rejected(method, params, bad_param):
    errors = _errors(method, params)

    if bad_param is not None:
        assert bad_param in errors


def test_enum_values_are_enforced():
    params = {
        "presentation_id": PID,
        "slide_index": 1,
        "begin_shape_name": "a",
        "end_shape_name": "b",
        "connector_type": "zigzag",
    }

    assert "connector_type" in _errors("pptx_add_connector", params)


def test_valid_calls_pass():
    validate_params("pptx_get_slide", {"presentation_id": PID, "slide_index": 1})
    validate_params("pptx_create_presentation", {"width": "10in", "height": 6858000})
    validate_params("pptx_create_presentation", {})


def test_optional_null_and_unlisted_parameters_are_accepted():
    validate_params("pptx_create_presentation", {"width": None, "not_in_catalog": object()})


def test_unknown_methods_are_not_checked():
    validate_params("pptx_does_not_exist", {"anything": 1})


def test_non_object_params_are_rejected():
    with pytest.raises(BridgeError) as excinfo:
        validate_params("pptx_get_slide", ["not", "a", "dict"])

    assert excinfo.value.code == "validation_error"


def test_service_rejects_bad_input_before_reaching_the_engine():
    service = PowerPointService()
    try:
        with pytest.raises(BridgeError) as excinfo:
            service.dispatch("pptx_add_slide", {"presentation_id": PID})
    finally:
        service.shutdown()

    assert excinfo.value.code == "validation_error"
    assert "layout_name" in excinfo.value.message
