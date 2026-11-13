"""Slide rendering to JPEG."""

from __future__ import annotations

import base64
import os
import tempfile
from pathlib import Path
from typing import Any

from engines.base import snapshot_slide_indices
from errors import ensure


class COMSnapshotMixin:
    """Slide rendering to JPEG."""

    def get_slide_snapshot(self, params: dict[str, Any]) -> dict[str, Any]:  # pragma: no cover - Windows only
        presentation_id = str(params["presentation_id"])
        slide_indices, single = snapshot_slide_indices(params)
        width_px = int(params.get("width_px") or 1280)

        snapshots = [
            self.worker.call(self._export_slide_snapshot_threadsafe, presentation_id, slide_index, width_px)
            for slide_index in slide_indices
        ]

        if single:
            return {
                "presentation_id": presentation_id,
                "slide_index": slide_indices[0],
                "mime_type": "image/jpeg",
                "snapshot_base64": snapshots[0],
                "width_px": width_px,
            }

        return {
            "presentation_id": presentation_id,
            "mime_type": "image/jpeg",
            "width_px": width_px,
            "snapshots": [
                {"slide_index": slide_index, "snapshot_base64": snapshot}
                for slide_index, snapshot in zip(slide_indices, snapshots, strict=True)
            ],
        }

    def _export_slide_snapshot_threadsafe(self, session_id: str, slide_index: int, width_px: int) -> str:
        prs = self._require_presentation(session_id)
        ensure(1 <= slide_index <= prs.Slides.Count, "validation_error", "slide_index out of bounds")
        slide = prs.Slides(slide_index)

        width_points = prs.PageSetup.SlideWidth
        height_points = prs.PageSetup.SlideHeight
        ratio = float(height_points) / float(width_points) if width_points else 0.5625
        height_px = max(1, int(width_px * ratio))

        fd, tmp_path = tempfile.mkstemp(suffix=".jpg", prefix="pptx-snapshot-")
        os.close(fd)
        try:
            slide.Export(tmp_path, "JPG", width_px, height_px)
            encoded = base64.b64encode(Path(tmp_path).read_bytes()).decode("ascii")
            return encoded
        finally:
            Path(tmp_path).unlink(missing_ok=True)
