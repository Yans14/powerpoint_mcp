from __future__ import annotations

from pathlib import Path

import pytest

from engines.base import MAX_CHECKPOINTS_PER_SESSION
from errors import BridgeError
from orchestrator.snapshot import SnapshotManager
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


def _slide_count(service, pid):
    return service.dispatch("pptx_get_presentation_state", {"presentation_id": pid})["slide_count"]


def _add_slide(service, pid):
    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Blank"})


def _checkpoint(service, pid, **extra):
    return service.dispatch("pptx_create_checkpoint", {"presentation_id": pid, **extra})


def _restore(service, pid, checkpoint_id):
    return service.dispatch("pptx_restore_checkpoint", {"presentation_id": pid, "checkpoint_id": checkpoint_id})


# --- Checkpoint tools ---------------------------------------------------------------------------


def test_restore_returns_to_the_checkpoint_and_can_be_repeated(service, deck):
    checkpoint = _checkpoint(service, deck, label="before appendix")
    assert checkpoint["label"] == "before appendix"

    _add_slide(service, deck)
    _add_slide(service, deck)
    restored = _restore(service, deck, checkpoint["checkpoint_id"])

    assert restored["slide_count"] == 1
    assert _slide_count(service, deck) == 1

    _add_slide(service, deck)
    _restore(service, deck, checkpoint["checkpoint_id"])
    assert _slide_count(service, deck) == 1


def test_checkpoint_captures_edits_not_yet_written_to_disk(service, deck):
    _add_slide(service, deck)
    checkpoint_id = _checkpoint(service, deck)["checkpoint_id"]
    _add_slide(service, deck)

    _restore(service, deck, checkpoint_id)

    assert _slide_count(service, deck) == 2


def test_restored_session_is_dirty_and_saves_the_restored_content(service, deck, tmp_path):
    checkpoint_id = _checkpoint(service, deck)["checkpoint_id"]
    _add_slide(service, deck)
    service.dispatch("pptx_save_presentation", {"presentation_id": deck, "output_path": str(tmp_path / "a.pptx")})

    _restore(service, deck, checkpoint_id)
    sessions = service.dispatch("pptx_list_open_presentations", {})["presentations"]
    reopened = service.dispatch(
        "pptx_save_presentation", {"presentation_id": deck, "output_path": str(tmp_path / "b.pptx")}
    )

    assert sessions[0]["dirty"] is True
    assert reopened["slide_count"] == 1


def test_only_the_most_recent_checkpoints_are_kept(service, deck):
    ids = [_checkpoint(service, deck)["checkpoint_id"] for _ in range(MAX_CHECKPOINTS_PER_SESSION)]

    extra = _checkpoint(service, deck)

    assert extra["dropped_checkpoint_id"] == ids[0]
    assert extra["checkpoint_count"] == MAX_CHECKPOINTS_PER_SESSION
    assert extra["warnings"]
    with pytest.raises(BridgeError) as excinfo:
        _restore(service, deck, ids[0])
    assert excinfo.value.code == "not_found"
    session_dir = Path(service.engine.sessions[deck].working_path).parent
    assert len(list(session_dir.glob("checkpoint-*.pptx"))) == MAX_CHECKPOINTS_PER_SESSION


def test_unknown_checkpoint_lists_the_available_ones(service, deck):
    checkpoint_id = _checkpoint(service, deck)["checkpoint_id"]

    with pytest.raises(BridgeError) as excinfo:
        _restore(service, deck, "nope")

    assert excinfo.value.code == "not_found"
    assert excinfo.value.details["available_checkpoints"] == [checkpoint_id]
    assert "pptx_create_checkpoint" in excinfo.value.hint


def test_checkpoints_are_removed_with_the_session(service, deck):
    _checkpoint(service, deck)
    session_dir = Path(service.engine.sessions[deck].working_path).parent

    service.dispatch("pptx_close_presentation", {"presentation_id": deck})

    assert not session_dir.exists()


def test_engines_without_rollback_refuse_checkpoints_with_a_hint(service, deck, monkeypatch):
    monkeypatch.setattr(service.engine, "supports_rollback", False)

    with pytest.raises(BridgeError) as excinfo:
        _checkpoint(service, deck)

    assert excinfo.value.code == "dependency_missing"
    assert "pptx_save_presentation" in excinfo.value.hint
    assert SnapshotManager(service.engine).create(service.engine.sessions[deck]) is None


# --- Engine-agnostic deck access ----------------------------------------------------------------


def test_checkers_read_decks_through_the_engine_interface(service, deck, monkeypatch):
    real = service.engine.get_pptx
    requested = []

    def tracking_get_pptx(presentation_id):
        requested.append(presentation_id)
        return real(presentation_id)

    monkeypatch.setattr(service.engine, "get_pptx", tracking_get_pptx)
    # Checkers must not depend on OOXML internals: hide them to prove it.
    service.engine.sessions[deck].extra["prs_hidden"] = service.engine.sessions[deck].extra.pop("prs")
    monkeypatch.setattr(service.engine, "_prs", lambda session: session.extra["prs_hidden"])

    result = service.dispatch("pptx_check_content", {"presentation_id": deck})

    assert requested == [deck]
    assert isinstance(result, dict)


def test_get_pptx_reflects_unsaved_edits(service, deck):
    _add_slide(service, deck)

    assert len(service.engine.get_pptx(deck).slides) == 2


def test_get_pptx_on_unknown_session_is_not_found(service):
    with pytest.raises(BridgeError) as excinfo:
        service.engine.get_pptx("missing")

    assert excinfo.value.code == "not_found"
