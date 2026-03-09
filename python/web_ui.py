"""FastAPI web UI for testing the PowerPoint MCP service."""

from __future__ import annotations

import asyncio
import json
import logging
import time
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException, WebSocket, WebSocketDisconnect
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from errors import BridgeError
from service import PowerPointService
from tool_catalog import get_catalog_by_category

# ---------------------------------------------------------------------------
# Logging -> WebSocket broadcaster
# ---------------------------------------------------------------------------


class _WSLogHandler(logging.Handler):
    """Broadcasts log records to connected WebSocket clients."""

    def __init__(self) -> None:
        super().__init__()
        self.connections: set[WebSocket] = set()
        self._loop: asyncio.AbstractEventLoop | None = None

    def emit(self, record: logging.LogRecord) -> None:
        if not self.connections or not self._loop:
            return
        entry = json.dumps(
            {
                "timestamp": record.created,
                "level": record.levelname,
                "logger": record.name,
                "message": self.format(record),
            }
        )
        asyncio.run_coroutine_threadsafe(self._broadcast(entry), self._loop)

    async def _broadcast(self, message: str) -> None:
        dead: set[WebSocket] = set()
        for ws in self.connections:
            try:
                await ws.send_text(message)
            except Exception:
                dead.add(ws)
        self.connections -= dead


_ws_handler = _WSLogHandler()
_ws_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
logging.getLogger().addHandler(_ws_handler)
logging.getLogger().setLevel(logging.INFO)

logger = logging.getLogger("web_ui")

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = FastAPI(title="PowerPoint MCP Test UI")
service = PowerPointService()

WEB_DIR = Path(__file__).resolve().parent.parent / "web"

# ---------------------------------------------------------------------------
# Models
# ---------------------------------------------------------------------------


class DispatchRequest(BaseModel):
    method: str
    params: dict[str, Any] = {}


# ---------------------------------------------------------------------------
# API routes
# ---------------------------------------------------------------------------


@app.get("/api/tools")
def get_tools() -> dict[str, Any]:
    return get_catalog_by_category()


@app.get("/api/engine")
def get_engine() -> dict[str, Any]:
    try:
        return {"success": True, "result": service.dispatch("pptx_get_engine_info", {})}
    except BridgeError as exc:
        return {"success": False, "error": exc.to_payload()}


@app.get("/api/sessions")
def get_sessions() -> dict[str, Any]:
    try:
        return {"success": True, "result": service.dispatch("pptx_list_open_presentations", {})}
    except BridgeError as exc:
        return {"success": False, "error": exc.to_payload()}


@app.post("/api/dispatch")
def dispatch_tool(request: DispatchRequest) -> dict[str, Any]:
    logger.info("dispatch %s %s", request.method, json.dumps(request.params, default=str)[:300])
    start = time.monotonic()
    try:
        result = service.dispatch(request.method, request.params)
        duration = round((time.monotonic() - start) * 1000, 2)
        logger.info("dispatch %s OK (%.1fms)", request.method, duration)
        return {"success": True, "result": result, "duration_ms": duration}
    except BridgeError as exc:
        duration = round((time.monotonic() - start) * 1000, 2)
        logger.warning("dispatch %s FAILED: %s", request.method, exc)
        raise HTTPException(
            status_code=400, detail={"success": False, "error": exc.to_payload(), "duration_ms": duration}
        ) from exc
    except Exception as exc:
        duration = round((time.monotonic() - start) * 1000, 2)
        logger.exception("dispatch %s ERROR", request.method)
        raise HTTPException(
            status_code=500,
            detail={
                "success": False,
                "error": {"code": "internal_error", "message": str(exc)},
                "duration_ms": duration,
            },
        ) from exc


# ---------------------------------------------------------------------------
# WebSocket logs
# ---------------------------------------------------------------------------


@app.websocket("/ws/logs")
async def websocket_logs(websocket: WebSocket) -> None:
    await websocket.accept()
    _ws_handler.connections.add(websocket)
    _ws_handler._loop = asyncio.get_event_loop()
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        _ws_handler.connections.discard(websocket)


# ---------------------------------------------------------------------------
# Static / SPA
# ---------------------------------------------------------------------------


@app.get("/")
def serve_index() -> FileResponse:
    return FileResponse(WEB_DIR / "index.html")


if WEB_DIR.is_dir():
    app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")
