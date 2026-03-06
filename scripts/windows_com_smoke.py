#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import json
import platform
import sys
import traceback
from collections.abc import Callable
from dataclasses import asdict, dataclass, field
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

REPO_ROOT = Path(__file__).resolve().parents[1]
PYTHON_ROOT = REPO_ROOT / "python"
if str(PYTHON_ROOT) not in sys.path:
    sys.path.insert(0, str(PYTHON_ROOT))

from errors import BridgeError  # noqa: E402
from service import PowerPointService  # noqa: E402


class SmokeFailure(Exception):
    """Raised when a smoke check fails."""


@dataclass
class StepResult:
    name: str
    status: str
    details: dict[str, Any] = field(default_factory=dict)
    error: dict[str, Any] | None = None


MANUAL_CHECKLIST = [
    "PowerPoint opened a visible window during automation and no crash occurred.",
    "No repair dialog appears when opening the saved output .pptx.",
    "Placeholder text inherits template formatting (font, size, color) rather than default text-box styling.",
    "Slide move/reorder operations appear in the expected sequence in PowerPoint UI.",
    "The generated snapshot image matches the expected slide canvas without clipping or overflow.",
]


def _run_step(name: str, func: Callable[[], Any], steps: list[StepResult]) -> Any:
    try:
        result = func()
        details: dict[str, Any]
        if isinstance(result, dict):
            details = {"keys": sorted(result.keys())}
        else:
            details = {"result_type": type(result).__name__}
        steps.append(StepResult(name=name, status="passed", details=details))
        return result
    except BridgeError as exc:
        steps.append(
            StepResult(
                name=name,
                status="failed",
                error={
                    "code": exc.code,
                    "message": exc.message,
                    "details": exc.details or {},
                },
            )
        )
        raise
    except Exception as exc:
        steps.append(
            StepResult(
                name=name,
                status="failed",
                error={
                    "code": "internal_error",
                    "message": str(exc),
                    "traceback": traceback.format_exc(),
                },
            )
        )
        raise


def _select_layout(layouts: list[dict[str, Any]], layout_name: str | None) -> str:
    if not layouts:
        raise SmokeFailure("No layouts returned by pptx_get_layouts.")

    if layout_name:
        exact = [layout for layout in layouts if layout.get("name") == layout_name]
        if not exact:
            available = [layout.get("name", "") for layout in layouts]
            raise SmokeFailure(f"Requested layout '{layout_name}' not found. Available layouts: {available}")
        return str(exact[0]["name"])

    preferred_names = ("Title Slide", "Title and Content")
    for preferred in preferred_names:
        for layout in layouts:
            if layout.get("name") == preferred:
                return preferred

    return str(layouts[0].get("name", ""))


def _select_text_placeholder(
    service: PowerPointService,
    presentation_id: str,
    slide_index: int,
    placeholders: list[dict[str, Any]],
    steps: list[StepResult],
) -> str:
    if not placeholders:
        raise SmokeFailure("No placeholders found on the added slide.")

    for placeholder in placeholders:
        placeholder_name = str(placeholder.get("name", ""))
        if not placeholder_name:
            continue

        try:
            _run_step(
                name=f"set_placeholder_text:{placeholder_name}",
                func=lambda p=placeholder_name: service.dispatch(
                    "pptx_set_placeholder_text",
                    {
                        "presentation_id": presentation_id,
                        "slide_index": slide_index,
                        "placeholder_name": p,
                        "text_content": "COM smoke validation text",
                    },
                ),
                steps=steps,
            )
            return placeholder_name
        except BridgeError as exc:
            if exc.code in {"conflict", "validation_error", "not_found"}:
                continue
            raise

    raise SmokeFailure("No text-capable placeholder accepted pptx_set_placeholder_text.")


def _write_report_markdown(report: dict[str, Any], markdown_path: Path) -> None:
    lines: list[str] = []
    lines.append("# Windows COM Smoke Report")
    lines.append("")
    lines.append(f"- Generated at (UTC): `{report['generated_at_utc']}`")
    lines.append(f"- Host: `{report['host']}`")
    lines.append(f"- Engine: `{report['engine']}`")
    lines.append(f"- Status: `{report['status']}`")
    lines.append("")

    if report.get("saved_pptx"):
        lines.append(f"- Saved presentation: `{report['saved_pptx']}`")
    if report.get("saved_snapshot_jpg"):
        lines.append(f"- Saved snapshot: `{report['saved_snapshot_jpg']}`")
    lines.append("")

    lines.append("## Automated Steps")
    lines.append("")
    for step in report.get("steps", []):
        prefix = "[x]" if step.get("status") == "passed" else "[ ]"
        line = f"- {prefix} `{step.get('name', '')}`"
        error = step.get("error")
        if error:
            line += f" - `{error.get('code', 'error')}: {error.get('message', '')}`"
        lines.append(line)
    lines.append("")

    lines.append("## Manual Parity Checklist")
    lines.append("")
    for item in report.get("manual_checklist", []):
        lines.append(f"- [ ] {item}")
    lines.append("")

    markdown_path.write_text("\n".join(lines), encoding="utf-8")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run automated Windows COM smoke checks and emit a parity checklist report.",
    )
    parser.add_argument(
        "--input-pptx",
        type=str,
        default="",
        help="Optional absolute path to an existing .pptx file to open. If omitted, a new presentation is created.",
    )
    parser.add_argument(
        "--layout-name",
        type=str,
        default="",
        help="Optional exact layout name to use for pptx_add_slide.",
    )
    parser.add_argument(
        "--output-dir",
        type=str,
        default=str(REPO_ROOT / "artifacts" / "com-smoke"),
        help="Directory for generated .pptx, snapshot, and report files.",
    )
    parser.add_argument(
        "--snapshot-width-px",
        type=int,
        default=1280,
        help="Width for pptx_get_slide_snapshot.",
    )
    parser.add_argument(
        "--skip-snapshot",
        action="store_true",
        help="Skip snapshot validation step.",
    )
    parser.add_argument(
        "--allow-ooxml",
        action="store_true",
        help="Allow execution even when COM is unavailable. Useful only for local script debugging.",
    )
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    timestamp = datetime.now(UTC).strftime("%Y%m%dT%H%M%SZ")

    output_dir = Path(args.output_dir).expanduser().resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    report_json = output_dir / f"com_smoke_report_{timestamp}.json"
    report_md = output_dir / f"com_smoke_report_{timestamp}.md"
    output_pptx = output_dir / f"com_smoke_output_{timestamp}.pptx"
    output_snapshot = output_dir / f"com_smoke_snapshot_{timestamp}.jpg"

    service = PowerPointService()
    presentation_id: str | None = None
    steps: list[StepResult] = []
    status = "passed"
    failure: dict[str, Any] | None = None
    engine = ""

    try:
        engine_info = _run_step(
            "pptx_get_engine_info",
            lambda: service.dispatch("pptx_get_engine_info", {}),
            steps,
        )
        engine = str(engine_info.get("engine", ""))

        if engine != "COM" and not args.allow_ooxml:
            raise SmokeFailure(
                "COM engine is not active. Run this script on a licensed Windows host with PowerPoint installed "
                "or pass --allow-ooxml for local script debugging."
            )

        if args.input_pptx:
            input_path = Path(args.input_pptx).expanduser().resolve()
            if not input_path.is_file():
                raise SmokeFailure(f"input-pptx does not exist: {input_path}")
            if input_path.suffix.lower() != ".pptx":
                raise SmokeFailure("input-pptx must point to a .pptx file")

            open_result = _run_step(
                "pptx_open_presentation",
                lambda: service.dispatch("pptx_open_presentation", {"file_path": str(input_path)}),
                steps,
            )
            presentation_id = str(open_result["presentation_id"])
        else:
            create_result = _run_step(
                "pptx_create_presentation",
                lambda: service.dispatch(
                    "pptx_create_presentation",
                    {
                        "width": "10in",
                        "height": "5.625in",
                    },
                ),
                steps,
            )
            presentation_id = str(create_result["presentation_id"])

        layouts_result = _run_step(
            "pptx_get_layouts",
            lambda: service.dispatch("pptx_get_layouts", {"presentation_id": presentation_id}),
            steps,
        )
        layout_name = _select_layout(layouts_result.get("layouts", []), args.layout_name or None)

        add_result = _run_step(
            "pptx_add_slide",
            lambda: service.dispatch(
                "pptx_add_slide",
                {
                    "presentation_id": presentation_id,
                    "layout_name": layout_name,
                },
            ),
            steps,
        )
        added_slide_index = int(add_result["added_slide_index"])

        placeholders_result = _run_step(
            "pptx_get_placeholders",
            lambda: service.dispatch(
                "pptx_get_placeholders",
                {
                    "presentation_id": presentation_id,
                    "slide_index": added_slide_index,
                },
            ),
            steps,
        )

        selected_placeholder = _select_text_placeholder(
            service=service,
            presentation_id=presentation_id,
            slide_index=added_slide_index,
            placeholders=list(placeholders_result.get("placeholders", [])),
            steps=steps,
        )

        _run_step(
            "pptx_get_placeholder_text",
            lambda: service.dispatch(
                "pptx_get_placeholder_text",
                {
                    "presentation_id": presentation_id,
                    "slide_index": added_slide_index,
                    "placeholder_name": selected_placeholder,
                },
            ),
            steps,
        )

        _run_step(
            "pptx_clear_placeholder",
            lambda: service.dispatch(
                "pptx_clear_placeholder",
                {
                    "presentation_id": presentation_id,
                    "slide_index": added_slide_index,
                    "placeholder_name": selected_placeholder,
                },
            ),
            steps,
        )

        if not args.skip_snapshot:
            snapshot_result = _run_step(
                "pptx_get_slide_snapshot",
                lambda: service.dispatch(
                    "pptx_get_slide_snapshot",
                    {
                        "presentation_id": presentation_id,
                        "slide_index": added_slide_index,
                        "width_px": int(args.snapshot_width_px),
                    },
                ),
                steps,
            )
            snapshot_b64 = str(snapshot_result.get("snapshot_base64", ""))
            if snapshot_b64:
                output_snapshot.write_bytes(base64.b64decode(snapshot_b64))
            else:
                raise SmokeFailure("Snapshot response did not include snapshot_base64.")

        duplicate_result = _run_step(
            "pptx_duplicate_slide",
            lambda: service.dispatch(
                "pptx_duplicate_slide",
                {
                    "presentation_id": presentation_id,
                    "source_index": added_slide_index,
                    "target_position": added_slide_index + 1,
                },
            ),
            steps,
        )
        duplicated_index = int(duplicate_result["duplicated_slide_index"])

        _run_step(
            "pptx_move_slide",
            lambda: service.dispatch(
                "pptx_move_slide",
                {
                    "presentation_id": presentation_id,
                    "from_index": duplicated_index,
                    "to_index": added_slide_index,
                },
            ),
            steps,
        )

        state_result = _run_step(
            "pptx_get_presentation_state",
            lambda: service.dispatch("pptx_get_presentation_state", {"presentation_id": presentation_id}),
            steps,
        )
        slide_count = int(state_result["slide_count"])
        if slide_count > 1:
            new_order = list(range(2, slide_count + 1)) + [1]
            _run_step(
                "pptx_reorder_slides",
                lambda: service.dispatch(
                    "pptx_reorder_slides",
                    {
                        "presentation_id": presentation_id,
                        "new_order": new_order,
                    },
                ),
                steps,
            )

        post_reorder_state = _run_step(
            "pptx_get_presentation_state:post_reorder",
            lambda: service.dispatch("pptx_get_presentation_state", {"presentation_id": presentation_id}),
            steps,
        )
        final_slide_count = int(post_reorder_state["slide_count"])
        if final_slide_count > 1:
            _run_step(
                "pptx_delete_slide",
                lambda: service.dispatch(
                    "pptx_delete_slide",
                    {
                        "presentation_id": presentation_id,
                        "slide_index": final_slide_count,
                    },
                ),
                steps,
            )

        _run_step(
            "pptx_save_presentation",
            lambda: service.dispatch(
                "pptx_save_presentation",
                {
                    "presentation_id": presentation_id,
                    "output_path": str(output_pptx),
                },
            ),
            steps,
        )

    except Exception as exc:
        status = "failed"
        failure = {
            "type": type(exc).__name__,
            "message": str(exc),
        }

    finally:
        if presentation_id:
            try:
                _run_step(
                    "pptx_close_presentation",
                    lambda: service.dispatch("pptx_close_presentation", {"presentation_id": presentation_id}),
                    steps,
                )
            except Exception:
                pass

        try:
            service.shutdown()
        except Exception:
            pass

    report: dict[str, Any] = {
        "generated_at_utc": datetime.now(UTC).isoformat(),
        "status": status,
        "host": {
            "platform": platform.platform(),
            "system": platform.system(),
            "release": platform.release(),
            "python_version": platform.python_version(),
        },
        "engine": engine,
        "saved_pptx": str(output_pptx) if output_pptx.exists() else "",
        "saved_snapshot_jpg": str(output_snapshot) if output_snapshot.exists() else "",
        "steps": [asdict(step) for step in steps],
        "manual_checklist": MANUAL_CHECKLIST,
    }

    if failure:
        report["failure"] = failure

    report_json.write_text(json.dumps(report, indent=2), encoding="utf-8")
    _write_report_markdown(report, report_md)

    print(f"Smoke report JSON: {report_json}")
    print(f"Smoke report Markdown checklist: {report_md}")
    if report["saved_pptx"]:
        print(f"Saved presentation: {report['saved_pptx']}")
    if report["saved_snapshot_jpg"]:
        print(f"Saved snapshot: {report['saved_snapshot_jpg']}")

    return 0 if status == "passed" else 1


if __name__ == "__main__":
    raise SystemExit(main())
