from __future__ import annotations

import shutil
import uuid
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any

from errors import SESSION_NOT_FOUND_HINT, BridgeError, ensure

if TYPE_CHECKING:
    from pptx.presentation import Presentation

MAX_SNAPSHOT_SLIDES = 20
MAX_CHECKPOINTS_PER_SESSION = 10


def snapshot_slide_indices(params: dict[str, Any]) -> tuple[list[int], bool]:
    """Slides requested from pptx_get_slide_snapshot.

    Returns the 1-based indices and whether the caller used the single-slide form
    (``slide_index``), which keeps the original flat response shape.
    """
    single = params.get("slide_index")
    many = params.get("slide_indices")
    ensure(
        (single is None) != (many is None),
        "validation_error",
        "Provide exactly one of slide_index or slide_indices.",
    )
    if single is not None:
        return [int(single)], True

    ensure(
        isinstance(many, list) and 1 <= len(many) <= MAX_SNAPSHOT_SLIDES,
        "validation_error",
        f"slide_indices must list between 1 and {MAX_SNAPSHOT_SLIDES} slides.",
    )
    return [int(index) for index in many], False


@dataclass
class EngineSession:
    id: str
    original_path: str
    working_path: str
    engine: str
    dirty: bool = False
    # Bumped on every in-memory edit; flushed_revision is the one last written to working_path.
    revision: int = 0
    flushed_revision: int = 0
    opened_at: datetime = field(default_factory=lambda: datetime.now(UTC))
    extra: dict[str, Any] = field(default_factory=dict)

    @property
    def original_path_safe(self) -> str:
        return self.original_path

    @property
    def working_path_obj(self) -> Path:
        return Path(self.working_path)


class BaseEngine(ABC):
    name: str
    sessions: dict[str, EngineSession]
    # True when flush_session/reload_session really round-trip the working file, which is what
    # restoring a session to an earlier state relies on.
    supports_rollback = False

    @abstractmethod
    def get_engine_info(self, params: dict[str, Any] | None = None) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def create_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def open_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def save_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def close_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def list_open_presentations(self, params: dict[str, Any] | None = None) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_presentation_state(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_layouts(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_layout_detail(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_masters(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_theme(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def add_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def duplicate_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def delete_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def reorder_slides(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def move_slide(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_slide_background(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_slide_snapshot(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_placeholders(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_placeholder_text(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_placeholder_image(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def clear_placeholder(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_placeholder_text(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_placeholder_rich_text(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def add_text_box(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_slide_text(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_shape_details(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def add_table(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_table(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_table_cell(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_table_data(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def add_shape(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def delete_shape(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_slide_notes(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_shape_text(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_slide_xml(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_shape_properties(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def clone_shape(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def group_shapes(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def ungroup_shapes(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_shape_z_order(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def add_image(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def add_line(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def find_replace_text(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    # --- Phase 5: Chart tools ---

    @abstractmethod
    def add_chart(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def update_chart_data(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_chart_style(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    # --- Phase 6: Agent workflow tools ---

    @abstractmethod
    def copy_shape_between_decks(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def get_slide_shapes(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_table_cell_merge(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    # --- Phase 7: Formatting & Fidelity tools ---

    @abstractmethod
    def set_paragraph_spacing(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_text_box_properties(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_table_style(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def set_shape_fill_gradient(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def add_connector(self, params: dict[str, Any]) -> dict[str, Any]:
        raise NotImplementedError

    @abstractmethod
    def shutdown(self) -> None:
        raise NotImplementedError

    # --- Read access for checkers and the agent ---

    @abstractmethod
    def get_pptx(self, presentation_id: str) -> Presentation:
        """A python-pptx ``Presentation`` reflecting the session's current content.

        Checkers and the agent's verification read decks through this instead of reaching into
        engine internals, so they work with every engine. Callers must treat it as read-only:
        depending on the engine it is either the live object or a detached copy.
        """
        raise NotImplementedError

    # --- Working-file synchronisation (used by checkpoints, batches and agent rollback) ---

    def flush_session(self, session: EngineSession) -> None:
        """Make ``session.working_path`` reflect every edit made so far.

        No-op for engines that keep the working file current themselves.
        """
        return None

    def reload_session(self, session: EngineSession) -> None:
        """Pick up a working file that was replaced on disk, discarding in-memory state."""
        return None

    def require_rollback_support(self) -> None:
        if not self.supports_rollback:
            raise BridgeError(
                code="dependency_missing",
                message=f"Checkpoints and rollback are not supported by the {self.name} engine.",
                details={"engine": self.name},
                hint="Save a copy with pptx_save_presentation before risky edits and reopen it to go back.",
            )

    def copy_working_file(self, session: EngineSession, destination: Path) -> None:
        """Write the session's current content to ``destination``."""
        self.require_rollback_support()
        self.flush_session(session)
        shutil.copy2(session.working_path, destination)

    def restore_working_file(self, session: EngineSession, source: Path) -> None:
        """Replace the session's content with a copy previously written by copy_working_file."""
        self.require_rollback_support()
        shutil.copy2(source, session.working_path)
        self.reload_session(session)

    # --- Checkpoint tools (pptx_create_checkpoint / pptx_restore_checkpoint) ---

    def _session_for(self, params: dict[str, Any]) -> EngineSession:
        presentation_id = str(params["presentation_id"])
        session = self.sessions.get(presentation_id)
        if session is None:
            raise BridgeError(
                code="not_found",
                message="Presentation session not found.",
                details={"presentation_id": presentation_id},
                hint=SESSION_NOT_FOUND_HINT,
            )
        return session

    def create_checkpoint(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._session_for(params)
        self.require_rollback_support()
        checkpoints: dict[str, dict[str, Any]] = session.extra.setdefault("checkpoints", {})

        dropped = None
        if len(checkpoints) >= MAX_CHECKPOINTS_PER_SESSION:
            dropped = next(iter(checkpoints))
            Path(checkpoints.pop(dropped)["path"]).unlink(missing_ok=True)

        checkpoint_id = uuid.uuid4().hex[:12]
        # Kept beside the working copy, so the session's cleanup removes it too.
        path = Path(session.working_path).with_name(f"checkpoint-{checkpoint_id}.pptx")
        self.copy_working_file(session, path)
        label = str(params.get("label") or "")
        checkpoints[checkpoint_id] = {
            "path": str(path),
            "label": label,
            "created_at": datetime.now(UTC).isoformat(),
        }

        result: dict[str, Any] = {
            "success": True,
            "presentation_id": session.id,
            "checkpoint_id": checkpoint_id,
            "label": label,
            "checkpoint_count": len(checkpoints),
        }
        if dropped:
            result["dropped_checkpoint_id"] = dropped
            result["warnings"] = [
                f"Only {MAX_CHECKPOINTS_PER_SESSION} checkpoints are kept per presentation; "
                f"the oldest ({dropped}) was removed."
            ]
        return result

    def restore_checkpoint(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._session_for(params)
        self.require_rollback_support()
        checkpoints: dict[str, dict[str, Any]] = session.extra.get("checkpoints", {})
        checkpoint_id = str(params["checkpoint_id"])
        checkpoint = checkpoints.get(checkpoint_id)
        if checkpoint is None:
            raise BridgeError(
                code="not_found",
                message=f"Checkpoint '{checkpoint_id}' not found.",
                details={"checkpoint_id": checkpoint_id, "available_checkpoints": list(checkpoints)},
                hint="Use a checkpoint_id returned by pptx_create_checkpoint for this presentation.",
            )

        self.restore_working_file(session, Path(checkpoint["path"]))
        session.dirty = True
        return {
            "success": True,
            "presentation_id": session.id,
            "checkpoint_id": checkpoint_id,
            "label": checkpoint["label"],
            # The checkpoint is kept: it can be restored again.
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }
