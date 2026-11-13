"""Session lifecycle: create, open, save, close, list."""

from __future__ import annotations

import shutil
import uuid
from typing import Any

from pptx import Presentation

from engines.base import EngineSession
from utils.paths import validate_existing_file, validate_output_file
from utils.units import to_emu


class OOXMLSessionMixin:
    """Session lifecycle: create, open, save, close, list."""

    def get_engine_info(self, params: dict[str, Any] | None = None) -> dict[str, Any]:
        return {
            "engine": self.name,
            **self.metadata,
        }

    def create_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        width = params.get("width")
        height = params.get("height")
        template_path = params.get("template_path")

        if template_path:
            template = validate_existing_file(str(template_path), expected_suffixes=(".pptx", ".potx"))
            session_dir, tmp_path = self._new_session_files()
            shutil.copy2(template, tmp_path)
            prs = Presentation(str(tmp_path))
            if width is not None or height is not None:
                # An explicit size overrides the template's page setup.
                if width is not None:
                    prs.slide_width = to_emu(width)
                if height is not None:
                    prs.slide_height = to_emu(height)
                prs.save(tmp_path)
            original_path = str(template)
        else:
            prs = Presentation()
            if width is not None:
                prs.slide_width = to_emu(width)
            if height is not None:
                prs.slide_height = to_emu(height)
            session_dir, tmp_path = self._new_session_files()
            prs.save(tmp_path)
            prs = Presentation(tmp_path)
            original_path = ""

        session_id = str(uuid.uuid4())
        session = EngineSession(
            id=session_id,
            original_path=original_path,
            working_path=tmp_path,
            engine=self.name,
            dirty=True,
            extra={"prs": prs, "session_dir": session_dir},
        )
        self.sessions[session_id] = session

        return {
            "success": True,
            "presentation_id": session_id,
            "engine": self.name,
            "slide_width_emu": prs.slide_width,
            "slide_height_emu": prs.slide_height,
            "slide_count": len(prs.slides),
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def open_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        source = validate_existing_file(str(params["file_path"]), expected_suffixes=(".pptx", ".potx"))

        session_dir, tmp_path = self._new_session_files()
        shutil.copy2(source, tmp_path)

        session_id = str(uuid.uuid4())
        prs = Presentation(tmp_path)
        session = EngineSession(
            id=session_id,
            original_path=str(source),
            working_path=tmp_path,
            engine=self.name,
            dirty=False,
            extra={"prs": prs, "session_dir": session_dir},
        )
        self.sessions[session_id] = session

        return {
            "success": True,
            "presentation_id": session_id,
            "slide_count": len(prs.slides),
            "layout_names": [layout.name for layout in prs.slide_layouts],
            "presentation_state": self.get_presentation_state({"presentation_id": session_id}),
        }

    def save_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        output_path = validate_output_file(str(params["output_path"]), expected_suffixes=(".pptx",))
        # Always write, even when no edit was recorded: saving must never depend on bookkeeping.
        self._prs(session).save(session.working_path)
        session.flushed_revision = session.revision
        shutil.copy2(session.working_path, output_path)
        session.dirty = False

        return {
            "success": True,
            "presentation_id": session.id,
            "saved_path": str(output_path),
            "file_size_bytes": output_path.stat().st_size,
            "presentation_state": self.get_presentation_state({"presentation_id": session.id}),
        }

    def close_presentation(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        self.sessions.pop(session.id, None)
        self._remove_session_files(session)

        return {
            "success": True,
            "presentation_id": session.id,
            "closed": True,
        }

    def list_open_presentations(self, params: dict[str, Any] | None = None) -> dict[str, Any]:
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

    def get_presentation_state(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        prs = self._prs(session)

        slides = []
        for index, slide in enumerate(prs.slides, start=1):
            slides.append(
                {
                    "index": index,
                    "title": self._slide_title(slide),
                    "layout": slide.slide_layout.name,
                    "shape_count": len(slide.shapes),
                }
            )

        return {
            "presentation_id": session.id,
            "engine": self.name,
            "slide_count": len(prs.slides),
            "slides": slides,
        }

    def get_pptx(self, presentation_id: str) -> Presentation:
        return self._prs(self._get_session(presentation_id))
