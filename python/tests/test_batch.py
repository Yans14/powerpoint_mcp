from __future__ import annotations

from pathlib import Path

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
    """A presentation with one title slide."""
    pid = service.dispatch("pptx_create_presentation", {})["presentation_id"]
    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})
    return pid


def _op(tool, **arguments):
    return {"tool": tool, "arguments": arguments}


def _slide_count(service, pid):
    return service.dispatch("pptx_get_presentation_state", {"presentation_id": pid})["slide_count"]


def _session_files(service, pid):
    session_dir = Path(service.engine.sessions[pid].working_path).parent
    return sorted(entry.name for entry in session_dir.iterdir())


def _run(service, pid, operations, **extra):
    return service.dispatch("pptx_batch", {"presentation_id": pid, "operations": operations, **extra})


def _failure(service, pid, operations, **extra):
    with pytest.raises(BridgeError) as excinfo:
        _run(service, pid, operations, **extra)
    return excinfo.value


def test_batch_runs_operations_in_order_and_reports_each_result(service, deck):
    result = _run(
        service,
        deck,
        [
            _op("pptx_add_slide", layout_name="Blank"),
            _op(
                "pptx_add_text_box",
                slide_index=2,
                text_content="Hello",
                left="1in",
                top="1in",
                width="3in",
                height="1in",
            ),
            _op("pptx_get_slide_text", slide_index=2),
        ],
    )

    assert result["success"] is True
    assert result["executed"] == 3
    assert result["slide_count"] == 2
    assert "presentation_state" not in result
    assert [item["tool"] for item in result["results"]] == [
        "pptx_add_slide",
        "pptx_add_text_box",
        "pptx_get_slide_text",
    ]
    assert "Hello" in str(result["results"][2]["result"])
    # Per-operation results are compact too.
    assert "presentation_state" not in result["results"][0]["result"]
    assert _session_files(service, deck) == ["presentation.pptx"]


def test_batch_can_return_the_full_state_once(service, deck):
    result = _run(service, deck, [_op("pptx_add_slide", layout_name="Blank")], include_state=True)

    assert [slide["index"] for slide in result["presentation_state"]["slides"]] == [1, 2]
    assert "presentation_state" not in result["results"][0]["result"]


def test_failed_atomic_batch_restores_the_presentation(service, deck, tmp_path):
    out = tmp_path / "saved.pptx"
    service.dispatch("pptx_save_presentation", {"presentation_id": deck, "output_path": str(out)})
    title = service.dispatch("pptx_get_placeholders", {"presentation_id": deck, "slide_index": 1})["placeholders"][0]

    error = _failure(
        service,
        deck,
        [
            _op("pptx_set_placeholder_text", slide_index=1, placeholder_name=title["name"], text_content="Changed"),
            _op("pptx_add_slide", layout_name="Blank"),
            _op("pptx_add_slide", layout_name="No Such Layout"),
            _op("pptx_add_slide", layout_name="Blank"),
        ],
    )

    assert error.code == "not_found"
    assert "operations[2] (pptx_add_slide)" in error.message
    assert "No change from this batch was kept." in error.message
    assert error.details["failed_operation_index"] == 2
    assert error.details["completed_operations"] == 2
    assert error.details["rolled_back"] is True
    assert error.details["error"]["details"]["available_layouts"]
    assert "pptx_get_layouts" in error.hint

    assert _slide_count(service, deck) == 1
    assert "Changed" not in str(service.dispatch("pptx_get_slide_text", {"presentation_id": deck, "slide_index": 1}))
    sessions = service.dispatch("pptx_list_open_presentations", {})["presentations"]
    assert sessions[0]["dirty"] is False
    assert _session_files(service, deck) == ["presentation.pptx"]

    # The session stays usable after the rollback.
    assert _run(service, deck, [_op("pptx_add_slide", layout_name="Blank")])["slide_count"] == 2


def test_failed_non_atomic_batch_keeps_earlier_operations(service, deck):
    error = _failure(
        service,
        deck,
        [_op("pptx_add_slide", layout_name="Blank"), _op("pptx_delete_slide", slide_index=9)],
        atomic=False,
    )

    assert error.details["rolled_back"] is False
    assert "The 1 operation(s) before it were applied and kept." in error.message
    assert _slide_count(service, deck) == 2


def test_engine_without_rollback_support_reports_that_changes_were_kept(service, deck, monkeypatch):
    monkeypatch.setattr(service.engine, "supports_rollback", False)

    error = _failure(
        service,
        deck,
        [_op("pptx_add_slide", layout_name="Blank"), _op("pptx_add_slide", layout_name="No Such Layout")],
    )

    assert error.details["rolled_back"] is False
    assert "applied and kept" in error.message
    assert _slide_count(service, deck) == 2


def test_read_only_batch_takes_no_backup(service, deck, monkeypatch):
    def fail_on_flush(session):
        raise AssertionError("a read-only batch must not write the working file")

    monkeypatch.setattr(service.engine, "flush_session", fail_on_flush)

    result = _run(service, deck, [_op("pptx_get_slide", slide_index=1), _op("pptx_get_layouts")])

    assert result["executed"] == 2


@pytest.mark.parametrize(
    "operation",
    [
        _op("pptx_save_presentation", output_path="/tmp/never-written.pptx"),
        _op("pptx_close_presentation"),
        _op("pptx_batch", operations=[]),
        _op("pptx_agent_status", task_id="t"),
        _op("pptx_get_engine_info"),
        _op("pptx_no_such_tool"),
    ],
)
def test_batch_rejects_tools_that_cannot_be_rolled_back_or_do_not_target_the_presentation(service, deck, operation):
    error = _failure(service, deck, [_op("pptx_add_slide", layout_name="Blank"), operation])

    assert error.code == "validation_error"
    assert "cannot be used inside pptx_batch" in error.message
    # Rejected before anything ran.
    assert _slide_count(service, deck) == 1


@pytest.mark.parametrize(
    "operations",
    [
        [],
        ["pptx_add_slide"],
        [{"arguments": {}}],
        [{"tool": "pptx_add_slide", "arguments": "layout_name=Blank"}],
        [_op("pptx_add_slide", layout_name="Blank", presentation_id="another-presentation")],
        [_op("pptx_add_slide", layout_name="Blank")] * 51,
    ],
)
def test_batch_rejects_malformed_operation_lists(service, deck, operations):
    error = _failure(service, deck, operations)

    assert error.code == "validation_error"
    assert _slide_count(service, deck) == 1


def test_batch_on_unknown_presentation_fails_with_the_usual_error(service):
    error = _failure(service, "missing", [_op("pptx_add_slide", layout_name="Blank")])

    assert error.code == "not_found"
    assert "pptx_list_open_presentations" in error.hint
