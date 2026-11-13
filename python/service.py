from __future__ import annotations

from typing import Any

from batch import run_batch
from checkers.checker_service import CheckerService
from engine_selector import detect_engine
from engines.base import BaseEngine
from engines.com_engine import COMEngine
from engines.ooxml_engine import OOXMLEngine
from errors import BridgeError, suggest_hint
from tool_catalog import BATCH_TOOL, TOOLS, engine_method_for, is_engine_tool
from validation import validate_params

# Tools an engine method serves directly; the agent's plans may only call these.
_ENGINE_TOOL_MAP = {tool["name"]: engine_method_for(tool) for tool in TOOLS if is_engine_tool(tool)}


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

        # Derived from the generated catalog: every tool except pptx_batch, handled separately.
        self.method_map: dict[str, str] = {
            tool["name"]: engine_method_for(tool) for tool in TOOLS if tool["name"] != BATCH_TOOL
        }

        # Checker service always available (no LLM needed)
        self.checker_service = CheckerService(engine=self.engine)

        # The embedded LLM agent is opt-in (PPTX_ENABLE_AGENT=1) and needs the agent extra.
        self.orchestrator = None
        self.agent_unavailable_reason = "The agent is disabled. Set PPTX_ENABLE_AGENT=1 to enable it."
        from orchestrator.config import AgentConfig, agent_enabled

        if agent_enabled():
            config = AgentConfig.from_env()
            if config is None:
                self.agent_unavailable_reason = (
                    "PPTX_ENABLE_AGENT=1 is set but no API key was found. Set PPTX_LLM_API_KEY, or "
                    "ANTHROPIC_API_KEY / OPENAI_API_KEY / AZURE_OPENAI_API_KEY for the chosen PPTX_LLM_PROVIDER."
                )
            else:
                try:
                    from orchestrator.agent import AgentOrchestrator

                    self.orchestrator = AgentOrchestrator(
                        engine=self.engine,
                        config=config,
                        method_map=dict(_ENGINE_TOOL_MAP),
                    )
                except (BridgeError, ImportError) as exc:
                    self.agent_unavailable_reason = (
                        f"The agent could not start ({exc}). Install its packages: "
                        "pip install -r requirements-agent.txt"
                    )

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
        if method == BATCH_TOOL:
            return run_batch(self, validate_params(method, params))

        method_name = self.method_map.get(method)
        if not method_name:
            raise BridgeError(
                code="not_found",
                message=f"Unknown method '{method}'.",
                details={"method": method},
            )

        # Callers other than the TypeScript layer (web UI, tests) do not pre-validate input.
        params = validate_params(method, params)

        # Route agent methods
        if method_name == "agent_dispatch":
            if self.orchestrator is None:
                raise BridgeError(code="dependency_missing", message=self.agent_unavailable_reason)
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
