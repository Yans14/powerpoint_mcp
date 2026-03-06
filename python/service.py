from __future__ import annotations

from typing import Any

from engine_selector import detect_engine
from engines.base import BaseEngine
from engines.com_engine import COMEngine
from engines.ooxml_engine import OOXMLEngine
from errors import BridgeError
from handlers.discovery import DISCOVERY_METHODS
from handlers.placeholders import PLACEHOLDER_METHODS
from handlers.session import SESSION_METHODS
from handlers.slides import SLIDE_METHODS


class PowerPointService:
    def __init__(self) -> None:
        selected_engine, metadata = detect_engine()

        if selected_engine == "COM":
            try:
                self.engine: BaseEngine = COMEngine(metadata)
            except BridgeError as exc:
                fallback_metadata = {
                    **metadata,
                    "fallback_reason": exc.message,
                    "selected_engine": "OOXML",
                }
                self.engine = OOXMLEngine(fallback_metadata)
        else:
            self.engine = OOXMLEngine(metadata)

        self.method_map: dict[str, str] = {
            **SESSION_METHODS,
            **DISCOVERY_METHODS,
            **SLIDE_METHODS,
            **PLACEHOLDER_METHODS,
        }

    def dispatch(self, method: str, params: dict[str, Any]) -> Any:
        method_name = self.method_map.get(method)
        if not method_name:
            raise BridgeError(
                code="not_found",
                message=f"Unknown method '{method}'.",
                details={"method": method},
            )

        target = getattr(self.engine, method_name, None)
        if target is None:
            raise BridgeError(
                code="engine_error",
                message=f"Engine method '{method_name}' is not implemented.",
                details={"method": method, "engine": self.engine.name},
            )

        return target(params)

    def shutdown(self) -> None:
        self.engine.shutdown()
