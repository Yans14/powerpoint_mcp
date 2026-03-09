from __future__ import annotations

import shutil
from pathlib import Path

from pptx import Presentation

from engines.base import EngineSession


class SnapshotManager:
    def create(self, session: EngineSession) -> str:
        snapshot_path = session.working_path + ".snapshot"
        shutil.copy2(session.working_path, snapshot_path)
        return snapshot_path

    def restore(self, session: EngineSession, snapshot_path: str) -> None:
        shutil.copy2(snapshot_path, session.working_path)
        session.extra["prs"] = Presentation(session.working_path)
        session.dirty = True

    def remove(self, snapshot_path: str) -> None:
        Path(snapshot_path).unlink(missing_ok=True)
