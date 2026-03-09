from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

from engines.base import BaseEngine
from errors import BridgeError
from orchestrator.models import AgentTask, PlanStep, StepStatus


class PlanExecutor:
    def __init__(self, engine: BaseEngine, method_map: dict[str, str]) -> None:
        self.engine = engine
        self.method_map = method_map

    def execute(self, task: AgentTask, steps: list[PlanStep]) -> None:
        for step in steps:
            step.status = StepStatus.RUNNING
            task.updated_at = datetime.now(UTC)
            try:
                resolved = self._resolve_params(step.params, task.context)
                result = self._call(step.tool_name, resolved)
                step.result = result
                step.status = StepStatus.DONE
                self._update_context(task.context, task, step, result)
            except BridgeError as exc:
                step.status = StepStatus.FAILED
                step.error = f"{exc.code}: {exc.message}"
                raise

    def _resolve_params(self, params: dict[str, Any], context: dict[str, Any]) -> dict[str, Any]:
        resolved: dict[str, Any] = {}
        for key, value in params.items():
            if isinstance(value, str) and value.startswith("$"):
                var_name = value[1:]
                if var_name not in context:
                    raise BridgeError(
                        code="internal_error",
                        message=(
                            f"Plan step references undefined context variable "
                            f"'${var_name}'. Available: {list(context.keys())}"
                        ),
                    )
                resolved[key] = context[var_name]
            elif isinstance(value, dict):
                resolved[key] = self._resolve_params(value, context)
            elif isinstance(value, list):
                resolved[key] = [
                    self._resolve_params(item, context) if isinstance(item, dict) else item for item in value
                ]
            else:
                resolved[key] = value
        return resolved

    def _call(self, tool_name: str, params: dict[str, Any]) -> Any:
        method_name = self.method_map.get(tool_name)
        if not method_name:
            raise BridgeError(
                code="not_found",
                message=f"Tool '{tool_name}' not found in method map.",
                details={"tool_name": tool_name},
            )
        target = getattr(self.engine, method_name, None)
        if target is None:
            raise BridgeError(
                code="engine_error",
                message=f"Engine method '{method_name}' not implemented.",
            )
        return target(params)

    def _update_context(
        self,
        context: dict[str, Any],
        task: AgentTask,
        step: PlanStep,
        result: Any,
    ) -> None:
        if not isinstance(result, dict):
            return
        if step.captures and "presentation_id" in result:
            captured_id = str(result["presentation_id"])
            context[step.captures] = captured_id
            if captured_id != task.presentation_id:
                task.opened_presentation_ids.append(captured_id)
