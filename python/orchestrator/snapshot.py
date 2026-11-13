from __future__ import annotations

import shutil
from pathlib import Path

from engines.base import BaseEngine, EngineSession


class SnapshotManager:
    """Rollback points for agent plans: a copy of the session's working file, kept beside it.

    The copy lives in the session's own directory, so it disappears with the session even when
    the task that created it is never cancelled.
    """

    def __init__(self, engine: BaseEngine) -> None:
        self.engine = engine

    def create(self, session: EngineSession) -> str:
        # Edits may still be in memory only; the snapshot must capture them.
        self.engine.flush_session(session)
        snapshot_path = session.working_path + ".snapshot"
        shutil.copy2(session.working_path, snapshot_path)
        return snapshot_path

    def restore(self, session: EngineSession, snapshot_path: str) -> None:
        shutil.copy2(snapshot_path, session.working_path)
        self.engine.reload_session(session)
        session.dirty = True

    def remove(self, snapshot_path: str) -> None:
        Path(snapshot_path).unlink(missing_ok=True)
