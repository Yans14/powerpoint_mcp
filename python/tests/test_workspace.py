from __future__ import annotations

import os
import subprocess
import sys

import pytest

from workspace import Workspace, sweep_stale_workspaces


@pytest.fixture
def root(tmp_path, monkeypatch):
    root = tmp_path / "workspaces"
    monkeypatch.setenv("PPTX_WORKSPACE_DIR", str(root))
    return root


def _dead_pid() -> int:
    process = subprocess.Popen([sys.executable, "-c", "pass"])
    process.wait()
    return process.pid


def test_workspace_is_created_on_first_use_and_removed_on_cleanup(root):
    workspace = Workspace()
    assert not root.exists()

    session_dir = workspace.new_dir("session-")

    assert session_dir.is_dir()
    assert session_dir.parent == workspace.path
    assert workspace.path.parent == root
    assert workspace.path.name.startswith(f"bridge-{os.getpid()}-")

    owned = workspace.path
    workspace.cleanup()
    assert not owned.exists()


def test_sweep_removes_only_workspaces_whose_owner_is_gone(root):
    root.mkdir()
    dead = root / f"bridge-{_dead_pid()}-aaaa"
    alive = root / f"bridge-{os.getppid()}-bbbb"
    # A directory with our own pid predates this process: sweeping happens before we create ours.
    recycled_pid = root / f"bridge-{os.getpid()}-cccc"
    unrelated = root / "something-else"
    for directory in (dead, alive, recycled_pid, unrelated):
        directory.mkdir()

    removed = sweep_stale_workspaces()

    assert removed == 2
    assert not dead.exists()
    assert not recycled_pid.exists()
    assert alive.exists()
    assert unrelated.exists()


def test_sweep_without_a_root_directory_does_nothing(root):
    assert sweep_stale_workspaces() == 0
