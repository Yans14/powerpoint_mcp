"""pptx_batch: several operations on one presentation in a single bridge call.

Operations run in order through the normal dispatcher, so each one is validated and behaves
exactly as if it had been called on its own. In atomic mode the presentation is restored to its
pre-batch state when an operation fails, provided the engine supports it.
"""

from __future__ import annotations

import logging
import shutil
from pathlib import Path
from typing import TYPE_CHECKING, Any

from errors import BridgeError, ensure
from tool_catalog import TOOL_CATALOG

if TYPE_CHECKING:
    from service import PowerPointService

logger = logging.getLogger("pptx_mcp.batch")

MAX_BATCH_OPERATIONS = 50

# Saving and closing cannot be undone by restoring the working file, so they stay outside batches.
_EXCLUDED_TOOLS = {"pptx_batch", "pptx_save_presentation", "pptx_close_presentation"}
_TOOLS = {tool["name"]: tool for tool in TOOL_CATALOG}


def is_batchable(tool_name: str) -> bool:
    """A tool can run in a batch if it acts on a single presentation and can be rolled back."""
    spec = _TOOLS.get(tool_name)
    if spec is None or tool_name in _EXCLUDED_TOOLS or tool_name.startswith("pptx_agent_"):
        return False
    return any(param["name"] == "presentation_id" for param in spec["params"])


def run_batch(service: PowerPointService, params: dict[str, Any]) -> dict[str, Any]:
    presentation_id = str(params["presentation_id"])
    operations = params["operations"]
    atomic = params.get("atomic") is not False

    ensure(
        isinstance(operations, list) and 1 <= len(operations) <= MAX_BATCH_OPERATIONS,
        "validation_error",
        f"operations must list between 1 and {MAX_BATCH_OPERATIONS} operations.",
    )
    calls = [_parse_operation(index, operation, presentation_id) for index, operation in enumerate(operations)]

    engine = service.engine
    # Also fails early, with the usual error, when the presentation is not open.
    service.dispatch("pptx_get_presentation_state", {"presentation_id": presentation_id})
    session = engine.sessions[presentation_id]

    mutates = any(_TOOLS[tool]["mutating"] for tool, _ in calls)
    restorable = atomic and mutates and engine.supports_rollback
    backup_path = Path(session.working_path + ".batch")
    was_dirty = session.dirty

    if restorable:
        engine.flush_session(session)
        shutil.copy2(session.working_path, backup_path)

    results: list[dict[str, Any]] = []
    try:
        for index, (tool, arguments) in enumerate(calls):
            try:
                results.append({"tool": tool, "result": service.dispatch(tool, arguments)})
            except Exception as exc:
                if restorable:
                    shutil.copy2(backup_path, session.working_path)
                    engine.reload_session(session)
                    session.dirty = was_dirty
                raise _batch_failure(exc, index, tool, restorable, mutates) from exc
    finally:
        backup_path.unlink(missing_ok=True)

    return {
        "success": True,
        "presentation_id": presentation_id,
        "executed": len(results),
        "results": results,
        "presentation_state": engine.get_presentation_state({"presentation_id": presentation_id}),
    }


def _parse_operation(index: int, operation: Any, presentation_id: str) -> tuple[str, dict[str, Any]]:
    where = f"operations[{index}]"
    ensure(isinstance(operation, dict), "validation_error", f"{where} must be an object with a tool name.")

    tool = operation.get("tool")
    arguments = operation.get("arguments") or {}
    ensure(isinstance(tool, str), "validation_error", f"{where}.tool must be a tool name.")
    ensure(isinstance(arguments, dict), "validation_error", f"{where}.arguments must be an object.")

    if not is_batchable(tool):
        raise BridgeError(
            code="validation_error",
            message=f"{where}: {tool} cannot be used inside pptx_batch.",
            details={"operation_index": index, "tool": tool},
            hint=(
                "Batches accept tools that take presentation_id, except pptx_save_presentation, "
                "pptx_close_presentation, pptx_batch and the pptx_agent_* tools. Call those separately."
            ),
        )

    ensure(
        arguments.get("presentation_id", presentation_id) == presentation_id,
        "validation_error",
        f"{where}: every operation runs on the batch's presentation_id; remove presentation_id from its arguments.",
    )
    # Per-operation state would repeat the same payload for every step; the batch reports it once.
    arguments = {key: value for key, value in arguments.items() if key != "include_state"}
    return tool, {**arguments, "presentation_id": presentation_id}


def _batch_failure(exc: Exception, index: int, tool: str, restored: bool, mutates: bool) -> BridgeError:
    if isinstance(exc, BridgeError):
        cause = exc
    else:
        logger.exception("unhandled error in batch operation %d (%s)", index, tool)
        cause = BridgeError(code="internal_error", message=f"{type(exc).__name__}: {exc}")

    if restored or not mutates:
        outcome = "No change from this batch was kept."
    elif index == 0:
        outcome = "Nothing had been changed yet."
    else:
        outcome = f"The {index} operation(s) before it were applied and kept."

    return BridgeError(
        code=cause.code,
        message=f"Batch stopped at operations[{index}] ({tool}): {cause.message} {outcome}",
        details={
            "failed_operation_index": index,
            "tool": tool,
            "completed_operations": index,
            "rolled_back": restored,
            "error": cause.to_payload(),
        },
        hint=cause.hint,
    )
