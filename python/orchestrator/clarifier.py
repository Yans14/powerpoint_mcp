from __future__ import annotations

import json
from typing import Any

from pydantic import BaseModel

from orchestrator.capability_manifest import CAPABILITY_MANIFEST
from orchestrator.llm_client import LLMClient
from orchestrator.models import Question
from orchestrator.prompts import ANALYSIS_SYSTEM, ANALYSIS_USER


class RawQuestion(BaseModel):
    question_id: str
    text: str
    category: str
    choices: list[str]
    required: bool = True


class RawAnalysis(BaseModel):
    detected_intent: str
    complexity: str
    estimated_steps: int
    analysis_notes: str
    plan_preview: str
    questions: list[RawQuestion]


class LLMClarifier:
    def __init__(self, client: LLMClient) -> None:
        self.client = client

    def analyze(
        self,
        query: str,
        prs_state: dict[str, Any],
        theme: dict[str, Any],
        layouts: dict[str, Any],
    ) -> tuple[dict[str, Any], list[Question]]:
        system = ANALYSIS_SYSTEM.format(capability_manifest=CAPABILITY_MANIFEST)
        user = ANALYSIS_USER.format(
            query=query,
            state_json=json.dumps(prs_state, indent=2),
            theme_json=json.dumps(theme, indent=2),
            layouts_json=json.dumps(layouts, indent=2),
        )
        raw = self.client.call_structured(system, user, RawAnalysis)
        analysis = {
            "detected_intent": raw.detected_intent,
            "complexity": raw.complexity,
            "estimated_steps": raw.estimated_steps,
            "analysis_notes": raw.analysis_notes,
            "plan_preview": raw.plan_preview,
        }
        questions = [
            Question(
                question_id=q.question_id,
                text=q.text,
                category=q.category,
                choices=q.choices,
                required=q.required,
            )
            for q in raw.questions
        ]
        return analysis, questions
