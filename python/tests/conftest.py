from __future__ import annotations

import pytest


@pytest.fixture(autouse=True)
def _isolated_workspace(tmp_path, monkeypatch):
    """Keep session working files of every test under its own tmp_path, not the shared temp root."""
    monkeypatch.setenv("PPTX_WORKSPACE_DIR", str(tmp_path / "pptx-workspaces"))
