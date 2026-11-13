from __future__ import annotations

from typing import Any

from batch import run_batch
from checkers.checker_service import CheckerService
from engine_selector import detect_engine
from engines.base import BaseEngine
from engines.com_engine import COMEngine
from engines.ooxml_engine import OOXMLEngine
from errors import BridgeError, suggest_hint
from handlers.agent import AGENT_METHODS
from handlers.charts import CHART_METHODS
from handlers.checkers import CHECKER_METHODS
from handlers.discovery import DISCOVERY_METHODS
from handlers.placeholders import PLACEHOLDER_METHODS
from handlers.session import SESSION_METHODS
from handlers.shapes import SHAPE_METHODS
from handlers.slides import SLIDE_METHODS
from handlers.tables import TABLE_METHODS
from validation import validate_params


def _without_full_state(result: Any) -> Any:
    """Replace a result's ``presentation_state`` with just its ``slide_count``."""
    if not isinstance(result, dict) or not isinstance(result.get("presentation_state"), dict):
        return result

    state = result["presentation_state"]
    compact = {key: value for key, value in result.items() if key != "presentation_state"}
    compact.setdefault("slide_count", state.get("slide_count"))
    return compact


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
            **SHAPE_METHODS,
            **TABLE_METHODS,
            **CHART_METHODS,
            **AGENT_METHODS,
            **CHECKER_METHODS,
        }

        # Checker service always available (no LLM needed)
        self.checker_service = CheckerService(engine=self.engine)

        # Agent orchestrator is optional (requires LLM API key)
        self.orchestrator = None
        from orchestrator.config import AgentConfig

        config = AgentConfig.from_env()
        if config is not None:
            try:
                from orchestrator.agent import AgentOrchestrator

                core_map = {
                    **SESSION_METHODS,
                    **DISCOVERY_METHODS,
                    **SLIDE_METHODS,
                    **PLACEHOLDER_METHODS,
                    **SHAPE_METHODS,
                    **TABLE_METHODS,
                    **CHART_METHODS,
                }
                self.orchestrator = AgentOrchestrator(
                    engine=self.engine,
                    config=config,
                    method_map=core_map,
                )
            except (BridgeError, ImportError):
                pass  # LLM package or dependency not installed; agent tools disabled

    def dispatch(self, method: str, params: dict[str, Any]) -> Any:
        """Run one bridge method.

        Mutating methods answer with a ``slide_count`` instead of the full ``presentation_state``
        unless the caller passes ``include_state: true``: the per-slide listing repeated after
        every edit is mostly noise on large decks.
        """
        include_state = False
        if isinstance(params, dict) and "include_state" in params:
            params = dict(params)
            include_state = params.pop("include_state") is True

        try:
            result = self._dispatch(method, params)
        except BridgeError as exc:
            if exc.hint is None:
                exc.hint = suggest_hint(exc)
            raise

        return result if include_state else _without_full_state(result)

    def _dispatch(self, method: str, params: dict[str, Any]) -> Any:
        if method == "pptx_batch":
            validate_params(method, params)
            return run_batch(self, params)

        method_name = self.method_map.get(method)
        if not method_name:
            raise BridgeError(
                code="not_found",
                message=f"Unknown method '{method}'.",
                details={"method": method},
            )

        # Callers other than the TypeScript layer (web UI, tests) do not pre-validate input.
        validate_params(method, params)

        # Route agent methods
        if method_name == "agent_dispatch":
            if self.orchestrator is None:
                raise BridgeError(
                    code="dependency_missing",
                    message=(
                        "Agent orchestrator not configured. "
                        "Set PPTX_LLM_PROVIDER and ANTHROPIC_API_KEY or OPENAI_API_KEY."
                    ),
                )
            return self.orchestrator.dispatch(method, params)

        # Route checker methods
        if method_name == "checker_dispatch":
            return self.checker_service.dispatch(method, params)

        # Standard engine dispatch
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
