"""Session lifecycle: create, open, save, close, list."""

from __future__ import annotations

import io
import os
import shutil
import tempfile
import uuid
from pathlib import Path
from typing import Any

from engines.base import EngineSession
from utils.paths import validate_existing_file, validate_output_file
from utils.units import to_emu


class COMSessionMixin:
    """Session lifecycle: create, open, save, close, list."""

    def get_engine_info(
        self, params: dict[str, Any] | None = None
    ) -> dict[str, Any]:  # pragma: no cover - Windows only
        return {
            "engine": self.name,
            **self.metadata,
        }

    def create_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        session_id = str(uuid.uuid4())
        width = params.get("width")
        height = params.get("height")
        template_path = params.get("template_path")

        if template_path:
            source = validate_existing_file(str(template_path), expected_suffixes=(".pptx", ".potx"))
            fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
            os.close(fd)
            Path(tmp_path).unlink(missing_ok=True)
            shutil.copy2(source, tmp_path)
            self.worker.call(self._open_presentation_for_session, session_id, tmp_path)
            if width is not None or height is not None:
                # An explicit size overrides the template's page setup.
                self.worker.call(self._apply_page_size, session_id, width, height)
            original_path = str(source)
        else:
            fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
            os.close(fd)
            Path(tmp_path).unlink(missing_ok=True)
            self.worker.call(self._create_blank_presentation_for_session, session_id, tmp_path, width, height)
            original_path = ""

        self.sessions[session_id] = EngineSession(
            id=session_id,
            original_path=original_path,
            working_path=tmp_path,
            engine=self.name,
            dirty=True,
        )

        return {
            "success": True,
            "presentation_id": session_id,
            "engine": self.name,
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def _create_blank_presentation_for_session(self, session_id: str, tmp_path: str, width, height) -> None:
        prs = self._app.Presentations.Add()

        if width is not None:
            prs.PageSetup.SlideWidth = to_emu(width) / 12700
        if height is not None:
            prs.PageSetup.SlideHeight = to_emu(height) / 12700

        prs.SaveAs(tmp_path, 24)
        self._presentations[session_id] = prs

    def _apply_page_size(self, session_id: str, width, height) -> None:
        prs = self._presentations[session_id]

        if width is not None:
            prs.PageSetup.SlideWidth = to_emu(width) / 12700
        if height is not None:
            prs.PageSetup.SlideHeight = to_emu(height) / 12700

        prs.Save()

    def _open_presentation_for_session(self, session_id: str, path: str) -> None:
        prs = self._app.Presentations.Open(path, WithWindow=False)
        self._presentations[session_id] = prs

    def open_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        source = validate_existing_file(str(params["file_path"]), expected_suffixes=(".pptx", ".potx"))

        fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-session-")
        os.close(fd)
        Path(tmp_path).unlink(missing_ok=True)
        shutil.copy2(source, tmp_path)

        session_id = str(uuid.uuid4())
        self.worker.call(self._open_presentation_for_session, session_id, tmp_path)

        self.sessions[session_id] = EngineSession(
            id=session_id,
            original_path=str(source),
            working_path=tmp_path,
            engine=self.name,
            dirty=False,
        )

        layout_names = self.worker.call(self._list_layout_names, session_id)

        return {
            "success": True,
            "presentation_id": session_id,
            "slide_count": self.worker.call(self._slide_count, session_id),
            "layout_names": layout_names,
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def save_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        session = self._get_session(str(params["presentation_id"]))
        output_path = validate_output_file(str(params["output_path"]), expected_suffixes=(".pptx",))
        self.worker.call(self._save_copy_as, session.id, str(output_path))
        session.dirty = False

        return {
            "success": True,
            "presentation_id": session.id,
            "saved_path": str(output_path),
            "file_size_bytes": output_path.stat().st_size,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def get_pptx(self, presentation_id: str):  # pragma: no cover - Windows only
        from pptx import Presentation

        self._get_session(presentation_id)
        fd, tmp_path = tempfile.mkstemp(suffix=".pptx", prefix="pptx-export-")
        os.close(fd)
        Path(tmp_path).unlink(missing_ok=True)
        try:
            # SaveCopyAs leaves the open presentation (and its dirty state) untouched.
            self.worker.call(self._export_copy, presentation_id, tmp_path)
            with open(tmp_path, "rb") as handle:
                # Load fully into memory so the temporary file can be removed right away.
                return Presentation(io.BytesIO(handle.read()))
        finally:
            Path(tmp_path).unlink(missing_ok=True)

    def _export_copy(self, session_id: str, output_path: str) -> None:
        self._require_presentation(session_id).SaveCopyAs(output_path)

    def _save_copy_as(self, session_id: str, output_path: str) -> None:
        prs = self._require_presentation(session_id)
        prs.Save()
        prs.SaveCopyAs(output_path)

    def close_presentation(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        session = self._get_session(presentation_id)
        self.worker.call(self._close_presentation_threadsafe, presentation_id)
        self.sessions.pop(presentation_id, None)
        try:
            Path(session.working_path).unlink(missing_ok=True)
        except Exception:
            pass

        return {
            "success": True,
            "presentation_id": presentation_id,
            "closed": True,
        }

    def _close_presentation_threadsafe(self, session_id: str) -> None:
        prs = self._presentations.pop(session_id, None)
        if prs is not None:
            prs.Close()

    def list_open_presentations(self, params: dict[str, Any] | None = None) -> dict[str, Any]:  # pragma: no cover
        return {
            "presentations": [
                {
                    "presentation_id": session.id,
                    "original_path": session.original_path,
                    "working_path": session.working_path,
                    "engine": session.engine,
                    "dirty": session.dirty,
                    "opened_at": session.opened_at.isoformat(),
                }
                for session in self.sessions.values()
            ]
        }

    def get_presentation_state(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        self._get_session(presentation_id)
        slides = self.worker.call(self._presentation_state_slides, presentation_id)
        return {
            "presentation_id": presentation_id,
            "engine": self.name,
            "slide_count": len(slides),
            "slides": slides,
        }

    def _presentation_state_slides(self, session_id: str) -> list[dict[str, Any]]:
        prs = self._require_presentation(session_id)
        items: list[dict[str, Any]] = []
        for i in range(1, prs.Slides.Count + 1):
            slide = prs.Slides(i)
            items.append(
                {
                    "index": i,
                    "title": self._slide_title(slide),
                    "layout": slide.CustomLayout.Name if slide.CustomLayout else "",
                    "shape_count": slide.Shapes.Count,
                }
            )
        return items
