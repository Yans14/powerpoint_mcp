"""Slide rendering to JPEG."""

from __future__ import annotations

import base64
import os
import shutil
import subprocess
import time
import uuid
from pathlib import Path
from typing import Any

from engines.base import EngineSession, snapshot_slide_indices
from engines.ooxml.common import _DEFAULT_RENDER_TIMEOUT_SEC
from errors import BridgeError, ensure


class OOXMLSnapshotMixin:
    """Slide rendering to JPEG."""

    def get_slide_snapshot(self, params: dict[str, Any]) -> dict[str, Any]:
        session = self._get_session(str(params["presentation_id"]))
        slide_indices, single = snapshot_slide_indices(params)
        width_px = int(params.get("width_px") or 1280)
        for slide_index in slide_indices:
            self._assert_slide_index(self._prs(session), slide_index)

        soffice = shutil.which("soffice")
        pdftoppm = shutil.which("pdftoppm")

        if not soffice or not pdftoppm:
            raise BridgeError(
                code="dependency_missing",
                message="pptx_get_slide_snapshot requires LibreOffice (soffice) and poppler (pdftoppm) in OOXML mode.",
                details={
                    "missing": [
                        dependency
                        for dependency, ok in {
                            "soffice": bool(soffice),
                            "pdftoppm": bool(pdftoppm),
                        }.items()
                        if not ok
                    ],
                    "install_hint": "Install LibreOffice and poppler-utils to enable slide snapshots.",
                },
            )

        timeout_sec = self._render_timeout_sec()
        deadline = time.monotonic() + timeout_sec
        pdf_path = self._rendered_pdf(session, soffice, deadline, timeout_sec)
        images = [
            self._render_slide_image(pdf_path, pdftoppm, slide_index, width_px, deadline, timeout_sec)
            for slide_index in slide_indices
        ]

        if single:
            return {
                "presentation_id": session.id,
                "slide_index": slide_indices[0],
                "mime_type": "image/jpeg",
                "snapshot_base64": images[0],
                "width_px": width_px,
            }

        return {
            "presentation_id": session.id,
            "mime_type": "image/jpeg",
            "width_px": width_px,
            "snapshots": [
                {"slide_index": slide_index, "snapshot_base64": image}
                for slide_index, image in zip(slide_indices, images, strict=True)
            ],
        }

    @staticmethod
    def _render_timeout_sec() -> float:
        try:
            value = float(os.environ.get("PPTX_RENDER_TIMEOUT_SEC", ""))
        except ValueError:
            return _DEFAULT_RENDER_TIMEOUT_SEC
        return value if value > 0 else _DEFAULT_RENDER_TIMEOUT_SEC

    def _rendered_pdf(self, session: EngineSession, soffice: str, deadline: float, timeout_sec: float) -> Path:
        """PDF rendering of the whole deck, converted at most once per revision of the session."""
        cached = session.extra.get("render_cache")
        if cached and cached["revision"] == session.revision and Path(cached["pdf_path"]).exists():
            return Path(cached["pdf_path"])

        # Rendering is a read: flushing writes the working copy without marking the session dirty.
        self.flush_session(session)

        source_path = Path(session.working_path)
        render_dir = source_path.parent / "render"
        shutil.rmtree(render_dir, ignore_errors=True)
        render_dir.mkdir(parents=True)
        session.extra.pop("render_cache", None)

        self._run_renderer(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(render_dir), str(source_path)],
            "LibreOffice PDF export",
            deadline,
            timeout_sec,
        )

        pdf_path = render_dir / f"{source_path.stem}.pdf"
        ensure(pdf_path.exists(), "engine_error", "Expected PDF output was not generated.")
        session.extra["render_cache"] = {"revision": session.revision, "pdf_path": str(pdf_path)}
        return pdf_path

    def _render_slide_image(
        self, pdf_path: Path, pdftoppm: str, slide_index: int, width_px: int, deadline: float, timeout_sec: float
    ) -> str:
        image_prefix = pdf_path.parent / f"slide-{slide_index}-{uuid.uuid4().hex}"
        image_path = Path(f"{image_prefix}.jpg")
        try:
            self._run_renderer(
                [
                    pdftoppm,
                    "-f",
                    str(slide_index),
                    "-l",
                    str(slide_index),
                    "-jpeg",
                    "-singlefile",
                    "-scale-to-x",
                    str(width_px),
                    str(pdf_path),
                    str(image_prefix),
                ],
                "pdftoppm rendering",
                deadline,
                timeout_sec,
            )
            ensure(image_path.exists(), "engine_error", "Slide snapshot image was not generated.")
            return base64.b64encode(image_path.read_bytes()).decode("ascii")
        finally:
            image_path.unlink(missing_ok=True)

    @staticmethod
    def _run_renderer(command: list[str], label: str, deadline: float, timeout_sec: float) -> None:
        """Run an external renderer within what is left of the snapshot's time budget."""
        timeout_error = BridgeError(
            code="timeout",
            message=f"Slide rendering did not finish within {timeout_sec:g} s (stopped during: {label}).",
            details={"timeout_sec": timeout_sec, "hint": "Raise PPTX_RENDER_TIMEOUT_SEC for very large decks."},
        )
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise timeout_error

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                # Never let a child read the bridge's stdin: it carries the protocol stream.
                stdin=subprocess.DEVNULL,
                timeout=remaining,
            )
        except subprocess.TimeoutExpired:
            raise timeout_error from None

        if result.returncode != 0:
            raise BridgeError(
                code="engine_error",
                message=f"{label} failed.",
                details={"stdout": result.stdout, "stderr": result.stderr},
            )
