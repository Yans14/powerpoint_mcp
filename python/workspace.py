"""Scratch space for a bridge process.

Everything the server writes outside the user's own files (session working copies, rollback
snapshots, render output) lives under one directory per process, so it can be removed in one go
at shutdown. A process that is killed cannot clean up after itself; the next bridge to start
sweeps the directories whose owner is gone.

Layout: ``<tmp>/pptx-mcp/bridge-<pid>-<random>/``. Set ``PPTX_WORKSPACE_DIR`` to relocate the root.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import sys
import tempfile
import uuid
from pathlib import Path

logger = logging.getLogger("pptx_mcp.workspace")

_DIR_PATTERN = re.compile(r"^bridge-(\d+)-")


def workspace_root() -> Path:
    override = os.environ.get("PPTX_WORKSPACE_DIR")
    return Path(override) if override else Path(tempfile.gettempdir()) / "pptx-mcp"


class Workspace:
    """Private directory owned by the current process, created on first use."""

    def __init__(self) -> None:
        self._path: Path | None = None

    @property
    def path(self) -> Path:
        if self._path is None:
            root = workspace_root()
            root.mkdir(parents=True, exist_ok=True)
            self._path = Path(tempfile.mkdtemp(prefix=f"bridge-{os.getpid()}-", dir=root))
        return self._path

    def new_dir(self, prefix: str) -> Path:
        directory = self.path / f"{prefix}{uuid.uuid4().hex}"
        directory.mkdir()
        return directory

    def cleanup(self) -> None:
        if self._path is not None:
            shutil.rmtree(self._path, ignore_errors=True)
            self._path = None


def sweep_stale_workspaces() -> int:
    """Delete workspaces left behind by bridge processes that no longer exist.

    Call at startup, before this process creates its own workspace: a directory carrying our own
    pid can then only be a leftover from an earlier process that had the same pid (common in
    containers, where pids repeat across runs).
    """
    root = workspace_root()
    if not root.is_dir():
        return 0

    removed = 0
    for entry in root.iterdir():
        match = _DIR_PATTERN.match(entry.name)
        if not match or not entry.is_dir():
            continue
        owner_pid = int(match.group(1))
        if owner_pid != os.getpid() and _process_alive(owner_pid):
            continue
        shutil.rmtree(entry, ignore_errors=True)
        removed += 1

    if removed:
        logger.info("removed %d stale workspace(s) from %s", removed, root)
    return removed


def _process_alive(pid: int) -> bool:
    """Best-effort liveness check. Answers True when unsure, so a live workspace is never deleted."""
    if sys.platform == "win32":  # pragma: no cover - Windows only
        return _windows_process_alive(pid)

    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except OSError:
        return True
    return True


def _windows_process_alive(pid: int) -> bool:  # pragma: no cover - Windows only
    # os.kill(pid, 0) is not a probe on Windows: it terminates the target process.
    import ctypes
    from ctypes import wintypes

    process_query_limited_information = 0x1000
    error_invalid_parameter = 87
    still_active = 259

    try:
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel32.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]

        handle = kernel32.OpenProcess(process_query_limited_information, False, pid)
        if not handle:
            return ctypes.get_last_error() != error_invalid_parameter
        try:
            exit_code = wintypes.DWORD()
            if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                return True
            return exit_code.value == still_active
        finally:
            kernel32.CloseHandle(handle)
    except Exception:
        return True
