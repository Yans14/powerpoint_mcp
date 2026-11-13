"""Line-delimited JSON-RPC bridge between the TypeScript MCP server and the PowerPoint engines.

Requests arrive on stdin, one JSON object per line; each gets exactly one JSON line in reply.
Diagnostics go to stderr only.

Environment:
    PPTX_LOG_LEVEL   Level for the ``pptx_mcp.*`` loggers (DEBUG, INFO, WARNING, ...). Default INFO.
"""

from __future__ import annotations

import json
import logging
import os
import signal
import sys
import time
from typing import Any, TextIO

from pydantic import ValidationError

from errors import BridgeError
from models.ooxml import BridgeRequest
from workspace import sweep_stale_workspaces

logger = logging.getLogger("pptx_mcp.bridge")


def claim_protocol_stream() -> TextIO:
    """Reserve the process's stdout for protocol replies and send every other write to stderr.

    Replies are parsed line by line by the parent, so a stray ``print`` from a dependency, or
    output from a child process that inherited fd 1, would corrupt them. The original stdout is
    duplicated to a private descriptor and fd 1 is repointed at stderr, which covers Python-level
    and C-level writers alike.
    """
    sys.stdout.flush()
    protocol_fd = os.dup(1)
    os.dup2(2, 1)
    sys.stdout = sys.stderr
    return os.fdopen(protocol_fd, "w", encoding="utf-8", newline="\n")


def configure_logging() -> None:
    level = getattr(logging, os.environ.get("PPTX_LOG_LEVEL", "INFO").upper(), None)
    logging.basicConfig(
        stream=sys.stderr,
        level=logging.WARNING,
        format="%(asctime)s %(levelname)s %(name)s %(message)s",
    )
    # Third-party libraries stay at WARNING; only our own loggers follow PPTX_LOG_LEVEL.
    logging.getLogger("pptx_mcp").setLevel(level if isinstance(level, int) else logging.INFO)


def handle_line(service: Any, line: str) -> tuple[dict[str, Any], bool]:
    """Process one request line. Returns the reply and whether the bridge should stop."""
    request_id = "unknown"
    method = "?"
    started = time.perf_counter()

    try:
        request = BridgeRequest.model_validate(json.loads(line))
        request_id, method = request.id, request.method

        if method == "__shutdown__":
            # Clean up before acknowledging: the parent may kill us right after the reply.
            service.shutdown()
            return {"id": request_id, "result": {"ok": True}}, True

        response: dict[str, Any] = {"id": request_id, "result": service.dispatch(method, request.params)}
        status = "ok"
    except BridgeError as exc:
        response = {"id": request_id, "error": exc.to_payload()}
        status = exc.code
    except (json.JSONDecodeError, ValidationError) as exc:
        response = {
            "id": request_id,
            "error": {"code": "validation_error", "message": f"Malformed bridge request: {exc}"},
        }
        status = "validation_error"
    except Exception as exc:
        # The traceback belongs in the server log, not in a reply that ends up in a model's context.
        logger.exception("unhandled error id=%s method=%s", request_id, method)
        response = {
            "id": request_id,
            "error": {
                "code": "internal_error",
                "message": f"{type(exc).__name__}: {exc}",
                "details": {"error_type": type(exc).__name__, "request_id": request_id},
            },
        }
        status = "internal_error"

    duration_ms = (time.perf_counter() - started) * 1000
    logger.info("request id=%s method=%s status=%s duration_ms=%.1f", request_id, method, status, duration_ms)
    return response, False


def _exit_on_signal(signum: int, frame: Any) -> None:
    # Unwind through main()'s finally block so session files are removed.
    raise SystemExit(0)


def main() -> None:
    protocol_out = claim_protocol_stream()
    configure_logging()
    signal.signal(signal.SIGTERM, _exit_on_signal)

    # Imported only now: the engines pull in heavy third-party packages, and anything they print
    # while loading must not reach the protocol stream.
    from service import PowerPointService

    sweep_stale_workspaces()
    service = PowerPointService()
    logger.info("bridge started engine=%s pid=%d", service.engine.name, os.getpid())

    shut_down = False
    try:
        # Read bytes and decode explicitly: the parent always sends UTF-8, whereas a text-mode
        # stdin would follow the platform locale (cp1252 on Windows) and mangle non-ASCII text.
        for raw_line in sys.stdin.buffer:
            line = raw_line.decode("utf-8", errors="replace").strip()
            if not line:
                continue

            response, shut_down = handle_line(service, line)
            protocol_out.write(json.dumps(response, ensure_ascii=True, default=str) + "\n")
            protocol_out.flush()
            if shut_down:
                return
    finally:
        # Covers the parent closing stdin or sending SIGTERM without a __shutdown__ request,
        # so session working copies are not leaked in the temp directory.
        if not shut_down:
            service.shutdown()


if __name__ == "__main__":
    main()
