from __future__ import annotations

from pathlib import Path

from engines.base import BaseEngine, EngineSession


class SnapshotManager:
    """Rollback points for agent plans: a copy of the session's working file, kept beside it.

    The copy lives in the session's own directory, so it disappears with the session even when
    the task that created it is never cancelled.
    """

    def __init__(self, engine: BaseEngine) -> None:
        self.engine = engine

    def create(self, session: EngineSession) -> str | None:
        """Write a rollback copy. Returns None when the engine cannot restore one."""
        if not self.engine.supports_rollback:
            return None
        snapshot_path = session.working_path + ".snapshot"
        self.engine.copy_working_file(session, Path(snapshot_path))
        return snapshot_path

    def restore(self, session: EngineSession, snapshot_path: str) -> None:
        self.engine.restore_working_file(session, Path(snapshot_path))
        session.dirty = True

    def remove(self, snapshot_path: str) -> None:
        Path(snapshot_path).unlink(missing_ok=True)
