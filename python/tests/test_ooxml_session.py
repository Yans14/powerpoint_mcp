from __future__ import annotations

import base64
import subprocess
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches

from errors import BridgeError
from orchestrator.snapshot import SnapshotManager
from service import PowerPointService


@pytest.fixture
def service():
    svc = PowerPointService()
    yield svc
    svc.shutdown()


def _is_dirty(service, pid):
    sessions = service.dispatch("pptx_list_open_presentations", {})["presentations"]
    return next(item["dirty"] for item in sessions if item["presentation_id"] == pid)


def _saved_deck(service, tmp_path):
    pid = service.dispatch("pptx_create_presentation", {})["presentation_id"]
    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})
    out = tmp_path / "deck.pptx"
    service.dispatch("pptx_save_presentation", {"presentation_id": pid, "output_path": str(out)})
    return str(out)


def test_create_from_template_applies_requested_size(service, tmp_path):
    template = _saved_deck(service, tmp_path)

    result = service.dispatch("pptx_create_presentation", {"template_path": template, "width": "5in", "height": "4in"})

    assert result["slide_width_emu"] == Inches(5)
    assert result["slide_height_emu"] == Inches(4)
    assert result["slide_count"] == 1


def test_create_from_template_keeps_template_size_by_default(service, tmp_path):
    template = _saved_deck(service, tmp_path)

    result = service.dispatch("pptx_create_presentation", {"template_path": template})

    assert result["slide_width_emu"] == Inches(10)


def test_snapshot_does_not_mark_session_dirty(service, tmp_path, monkeypatch):
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]
    assert _is_dirty(service, pid) is False

    # Stand in for LibreOffice/poppler so the test reaches the flush that precedes rendering.
    monkeypatch.setattr("engines.ooxml_engine.shutil.which", lambda name: f"/fake/{name}")
    monkeypatch.setattr(
        "engines.ooxml_engine.subprocess.run",
        lambda *args, **kwargs: subprocess.CompletedProcess(args, 1, stdout="", stderr="renderer unavailable"),
    )

    with pytest.raises(BridgeError) as excinfo:
        service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, "slide_index": 1})

    assert excinfo.value.code == "engine_error"
    assert _is_dirty(service, pid) is False


# --- Working files ------------------------------------------------------------------------------


def _session(service, pid):
    return service.engine.sessions[pid]


def _slide_count_on_disk(path):
    return len(Presentation(str(path)).slides)


def test_edits_reach_the_working_file_only_when_something_reads_it(service, tmp_path):
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]
    working_path = _session(service, pid).working_path

    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})

    assert _is_dirty(service, pid) is True
    assert service.dispatch("pptx_get_presentation_state", {"presentation_id": pid})["slide_count"] == 2
    assert _slide_count_on_disk(working_path) == 1

    out = tmp_path / "out.pptx"
    service.dispatch("pptx_save_presentation", {"presentation_id": pid, "output_path": str(out)})

    assert _slide_count_on_disk(out) == 2
    assert _slide_count_on_disk(working_path) == 2
    assert _is_dirty(service, pid) is False


def test_get_theme_reads_a_flushed_working_file(service, tmp_path):
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]
    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})

    theme = service.dispatch("pptx_get_theme", {"presentation_id": pid})["theme"]

    assert theme["colors"]
    assert _slide_count_on_disk(_session(service, pid).working_path) == 2


def test_closing_a_presentation_removes_all_its_files(service):
    pid = service.dispatch("pptx_create_presentation", {})["presentation_id"]
    session_dir = Path(_session(service, pid).working_path).parent
    (session_dir / "presentation.pptx.snapshot").write_bytes(b"rollback copy")

    service.dispatch("pptx_close_presentation", {"presentation_id": pid})

    assert not session_dir.exists()


def test_shutdown_removes_the_whole_workspace():
    svc = PowerPointService()
    first = svc.dispatch("pptx_create_presentation", {})["presentation_id"]
    svc.dispatch("pptx_create_presentation", {})
    workspace_dir = Path(_session(svc, first).working_path).parents[1]
    assert len(list(workspace_dir.iterdir())) == 2

    svc.shutdown()

    assert not workspace_dir.exists()


def test_rollback_snapshot_captures_unflushed_edits_and_restores_them(service, tmp_path):
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]
    session = _session(service, pid)
    manager = SnapshotManager(service.engine)

    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})
    snapshot_path = manager.create(session)
    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})

    assert Path(snapshot_path).parent == Path(session.working_path).parent
    assert _slide_count_on_disk(snapshot_path) == 2

    manager.restore(session, snapshot_path)

    assert service.dispatch("pptx_get_presentation_state", {"presentation_id": pid})["slide_count"] == 2
    out = tmp_path / "restored.pptx"
    service.dispatch("pptx_save_presentation", {"presentation_id": pid, "output_path": str(out)})
    assert _slide_count_on_disk(out) == 2


# --- Snapshot rendering -------------------------------------------------------------------------


class _FakeRenderers:
    """Stands in for soffice and pdftoppm: records invocations and writes the files they would."""

    def __init__(self, monkeypatch):
        self.commands: list[list[str]] = []
        self.kwargs: list[dict] = []
        self.pdf_slide_counts: list[int] = []
        monkeypatch.setattr("engines.ooxml_engine.shutil.which", lambda name: name)
        monkeypatch.setattr("engines.ooxml_engine.subprocess.run", self.run)

    def run(self, command, **kwargs):
        self.commands.append(command)
        self.kwargs.append(kwargs)
        if command[0] == "soffice":
            source = Path(command[-1])
            self.pdf_slide_counts.append(_slide_count_on_disk(source))
            (Path(command[command.index("--outdir") + 1]) / f"{source.stem}.pdf").write_bytes(b"%PDF")
        else:
            Path(f"{command[-1]}.jpg").write_bytes(f"slide {command[command.index('-f') + 1]}".encode())
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    def count(self, program):
        return sum(1 for command in self.commands if command[0] == program)


def _decoded(snapshot_base64):
    return base64.b64decode(snapshot_base64).decode()


def test_snapshot_converts_the_deck_once_per_revision(service, tmp_path, monkeypatch):
    renderers = _FakeRenderers(monkeypatch)
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]

    first = service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, "slide_index": 1})
    service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, "slide_index": 1, "width_px": 640})

    assert first["slide_index"] == 1
    assert _decoded(first["snapshot_base64"]) == "slide 1"
    assert renderers.count("soffice") == 1
    assert renderers.count("pdftoppm") == 2
    assert _is_dirty(service, pid) is False

    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})
    service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, "slide_index": 2})

    # The edit invalidated the cached PDF, and the renderer saw the edited deck.
    assert renderers.count("soffice") == 2
    assert renderers.pdf_slide_counts == [1, 2]


def test_snapshot_of_several_slides_in_one_call(service, tmp_path, monkeypatch):
    renderers = _FakeRenderers(monkeypatch)
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]
    service.dispatch("pptx_add_slide", {"presentation_id": pid, "layout_name": "Title Slide"})

    result = service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, "slide_indices": [2, 1]})

    assert "snapshot_base64" not in result
    assert [(item["slide_index"], _decoded(item["snapshot_base64"])) for item in result["snapshots"]] == [
        (2, "slide 2"),
        (1, "slide 1"),
    ]
    assert renderers.count("soffice") == 1
    render_dir = Path(_session(service, pid).working_path).parent / "render"
    assert [entry.name for entry in render_dir.iterdir()] == ["presentation.pdf"]


@pytest.mark.parametrize(
    "selection",
    [{}, {"slide_index": 1, "slide_indices": [1]}, {"slide_indices": []}, {"slide_indices": [5]}],
)
def test_snapshot_rejects_invalid_slide_selection(service, tmp_path, monkeypatch, selection):
    _FakeRenderers(monkeypatch)
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]

    with pytest.raises(BridgeError) as excinfo:
        service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, **selection})

    assert excinfo.value.code == "validation_error"


def test_snapshot_renderers_run_with_a_timeout_and_without_the_protocol_stdin(service, tmp_path, monkeypatch):
    renderers = _FakeRenderers(monkeypatch)
    monkeypatch.setenv("PPTX_RENDER_TIMEOUT_SEC", "12")
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]

    service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, "slide_index": 1})

    assert len(renderers.kwargs) == 2
    for kwargs in renderers.kwargs:
        assert kwargs["stdin"] is subprocess.DEVNULL
        assert 0 < kwargs["timeout"] <= 12


def test_snapshot_reports_a_renderer_that_hangs_as_a_timeout(service, tmp_path, monkeypatch):
    pid = service.dispatch("pptx_open_presentation", {"file_path": _saved_deck(service, tmp_path)})["presentation_id"]

    def hang(command, **kwargs):
        raise subprocess.TimeoutExpired(command, kwargs["timeout"])

    monkeypatch.setattr("engines.ooxml_engine.shutil.which", lambda name: name)
    monkeypatch.setattr("engines.ooxml_engine.subprocess.run", hang)

    with pytest.raises(BridgeError) as excinfo:
        service.dispatch("pptx_get_slide_snapshot", {"presentation_id": pid, "slide_index": 1})

    assert excinfo.value.code == "timeout"
    assert "LibreOffice PDF export" in excinfo.value.message
