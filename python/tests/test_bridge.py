from __future__ import annotations

import json
import logging
import os
import signal
import subprocess
import sys
import uuid
from pathlib import Path

import pytest

import bridge
from errors import BridgeError

PYTHON_DIR = Path(__file__).resolve().parents[1]
BRIDGE_SCRIPT = PYTHON_DIR / "bridge.py"


class _FakeService:
    def __init__(self, outcome=None):
        self.outcome = outcome
        self.shutdown_calls = 0

    def dispatch(self, method, params):
        if isinstance(self.outcome, Exception):
            raise self.outcome
        return self.outcome

    def shutdown(self):
        self.shutdown_calls += 1


def _request(method="pptx_anything", **params):
    return json.dumps({"id": "req-1", "method": method, "params": params})


# --- handle_line -------------------------------------------------------------------------------


def test_handle_line_returns_result():
    response, stop = bridge.handle_line(_FakeService({"value": 1}), _request())

    assert response == {"id": "req-1", "result": {"value": 1}}
    assert stop is False


def test_handle_line_maps_bridge_error():
    service = _FakeService(BridgeError(code="not_found", message="missing", details={"x": 1}))

    response, _ = bridge.handle_line(service, _request())

    assert response == {"id": "req-1", "error": {"code": "not_found", "message": "missing", "details": {"x": 1}}}


def test_unhandled_error_is_logged_with_traceback_but_not_returned(caplog):
    service = _FakeService(RuntimeError("boom"))

    with caplog.at_level(logging.ERROR, logger="pptx_mcp.bridge"):
        response, stop = bridge.handle_line(service, _request())

    assert stop is False
    assert response["error"]["code"] == "internal_error"
    assert response["error"]["message"] == "RuntimeError: boom"
    assert response["error"]["details"] == {"error_type": "RuntimeError", "request_id": "req-1"}
    assert "Traceback" not in json.dumps(response)
    assert "Traceback" in caplog.text and "boom" in caplog.text


def test_malformed_request_is_a_validation_error():
    response, stop = bridge.handle_line(_FakeService(), "{not json")

    assert stop is False
    assert response["id"] == "unknown"
    assert response["error"]["code"] == "validation_error"


def test_shutdown_request_cleans_up_before_replying():
    service = _FakeService()

    response, stop = bridge.handle_line(service, _request("__shutdown__"))

    assert response == {"id": "req-1", "result": {"ok": True}}
    assert stop is True
    assert service.shutdown_calls == 1


def test_request_log_line_carries_id_method_status_and_duration(caplog):
    with caplog.at_level(logging.INFO, logger="pptx_mcp.bridge"):
        bridge.handle_line(_FakeService({}), _request("pptx_get_engine_info"))

    assert "request id=req-1 method=pptx_get_engine_info status=ok duration_ms=" in caplog.text


# --- stdout protection --------------------------------------------------------------------------


def test_only_protocol_replies_reach_stdout():
    script = (
        "import os, subprocess, sys\n"
        "import bridge\n"
        "out = bridge.claim_protocol_stream()\n"
        "print('python-level noise')\n"
        "os.write(1, b'fd-level noise\\n')\n"
        "subprocess.run([sys.executable, '-c', 'print(\"child noise\")'], check=True)\n"
        "out.write('reply\\n')\n"
        "out.flush()\n"
    )

    result = subprocess.run(
        [sys.executable, "-c", script], cwd=PYTHON_DIR, capture_output=True, text=True, timeout=30, check=True
    )

    assert result.stdout == "reply\n"
    for noise in ("python-level noise", "fd-level noise", "child noise"):
        assert noise in result.stderr


# --- real process -------------------------------------------------------------------------------


class _BridgeProcess:
    def __init__(self, workspace_root: Path, **env):
        self.workspace_root = workspace_root
        self.process = subprocess.Popen(
            [sys.executable, str(BRIDGE_SCRIPT)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env={**os.environ, "PPTX_WORKSPACE_DIR": str(workspace_root), **env},
        )

    def call(self, method, **params):
        request_id = uuid.uuid4().hex
        line = json.dumps({"id": request_id, "method": method, "params": params}, ensure_ascii=False)
        self.process.stdin.write(line.encode("utf-8") + b"\n")
        self.process.stdin.flush()
        response = json.loads(self.process.stdout.readline())
        assert response["id"] == request_id
        assert "error" not in response, response
        return response["result"]

    def finish(self, timeout=15):
        """Close stdin and wait for exit. Returns (remaining stdout, stderr)."""
        stdout, stderr = self.process.communicate(timeout=timeout)
        return stdout.decode("utf-8"), stderr.decode("utf-8")

    def workspaces(self):
        return [entry.name for entry in self.workspace_root.glob("bridge-*")]

    def kill(self):
        if self.process.poll() is None:
            self.process.kill()
            self.process.communicate()


@pytest.fixture
def start_bridge(tmp_path):
    started: list[_BridgeProcess] = []

    def _start(**env):
        process = _BridgeProcess(tmp_path / "workspaces", **env)
        started.append(process)
        return process

    yield _start
    for process in started:
        process.kill()


def test_bridge_round_trips_non_ascii_text(start_bridge):
    process = start_bridge()
    text = "Été à Zürich — 日本語 ✓"

    pid = process.call("pptx_create_presentation")["presentation_id"]
    process.call("pptx_add_slide", presentation_id=pid, layout_name="Title Slide")
    placeholder = process.call("pptx_get_placeholders", presentation_id=pid, slide_index=1)["placeholders"][0]
    process.call(
        "pptx_set_placeholder_text",
        presentation_id=pid,
        slide_index=1,
        placeholder_name=placeholder["name"],
        text_content=text,
    )

    slide = process.call("pptx_get_slide", presentation_id=pid, slide_index=1)
    assert text in json.dumps(slide, ensure_ascii=False)


def test_bridge_removes_its_workspace_when_stdin_closes(start_bridge):
    process = start_bridge()
    process.call("pptx_create_presentation")
    assert len(process.workspaces()) == 1

    stdout, stderr = process.finish()

    assert process.process.returncode == 0
    assert stdout == ""
    assert process.workspaces() == []
    assert "method=pptx_create_presentation status=ok" in stderr


@pytest.mark.skipif(sys.platform == "win32", reason="SIGTERM cannot be handled on Windows")
def test_bridge_removes_its_workspace_on_sigterm(start_bridge):
    process = start_bridge()
    process.call("pptx_create_presentation")

    process.process.send_signal(signal.SIGTERM)
    process.process.wait(timeout=15)

    assert process.process.returncode == 0
    assert process.workspaces() == []


def test_bridge_sweeps_workspaces_of_dead_processes_at_startup(start_bridge, tmp_path):
    finished = subprocess.Popen([sys.executable, "-c", "pass"])
    finished.wait()
    stale = tmp_path / "workspaces" / f"bridge-{finished.pid}-leftover"
    stale.mkdir(parents=True)
    (stale / "presentation.pptx").write_bytes(b"stale")

    process = start_bridge()
    process.call("pptx_get_engine_info")

    assert not stale.exists()


def test_log_level_is_configurable(start_bridge):
    process = start_bridge(PPTX_LOG_LEVEL="WARNING")
    process.call("pptx_get_engine_info")

    _, stderr = process.finish()

    assert "method=pptx_get_engine_info" not in stderr
