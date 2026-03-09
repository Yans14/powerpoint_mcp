from __future__ import annotations

import json
from typing import Any

from pydantic import BaseModel

from orchestrator.capability_manifest import CAPABILITY_MANIFEST
from orchestrator.llm_client import LLMClient
from orchestrator.models import PlanStep, StepStatus
from orchestrator.prompts import PLANNING_SYSTEM, PLANNING_USER


class RawPlanStep(BaseModel):
    step_id: str
    description: str
    tool_name: str
    params: dict[str, Any]
    captures: str | None = None


class RawPlan(BaseModel):
    plan_summary: str
    steps: list[RawPlanStep]


class LLMPlanner:
    def __init__(self, client: LLMClient, max_steps: int = 100) -> None:
        self.client = client
        self.max_steps = max_steps

    def generate_plan(
        self,
        query: str,
        answers: dict[str, str],
        prs_state: dict[str, Any],
        theme: dict[str, Any],
    ) -> list[PlanStep]:
        system = PLANNING_SYSTEM.format(
            capability_manifest=CAPABILITY_MANIFEST,
            max_steps=self.max_steps,
        )
        user = PLANNING_USER.format(
            query=query,
            answers_json=json.dumps(answers, indent=2),
            state_json=json.dumps(prs_state, indent=2),
            theme_json=json.dumps(theme, indent=2),
        )
        raw = self.client.call_structured(system, user, RawPlan)
        return [
            PlanStep(
                step_id=s.step_id,
                description=s.description,
                tool_name=s.tool_name,
                params=s.params,
                captures=s.captures,
                status=StepStatus.PENDING,
            )
            for s in raw.steps
        ]
