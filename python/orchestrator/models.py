from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime
from enum import Enum
from typing import Any


class AgentState(str, Enum):
    ANALYZING = "analyzing"
    PLANNING = "planning"
    CLARIFYING = "clarifying"
    READY = "ready"
    EXECUTING = "executing"
    VERIFYING = "verifying"
    COMPLETE = "complete"
    FAILED = "failed"
    CANCELLED = "cancelled"


class StepStatus(str, Enum):
    PENDING = "pending"
    RUNNING = "running"
    DONE = "done"
    FAILED = "failed"
    SKIPPED = "skipped"


@dataclass
class Question:
    question_id: str
    text: str
    category: str
    choices: list[str]
    required: bool = True

    def to_dict(self) -> dict[str, Any]:
        return {
            "question_id": self.question_id,
            "text": self.text,
            "category": self.category,
            "choices": self.choices,
            "required": self.required,
        }


@dataclass
class PlanStep:
    step_id: str
    description: str
    tool_name: str
    params: dict[str, Any]
    captures: str | None = None
    status: StepStatus = StepStatus.PENDING
    result: Any | None = None
    error: str | None = None

    def to_dict(self) -> dict[str, Any]:
        return {
            "step_id": self.step_id,
            "description": self.description,
            "tool_name": self.tool_name,
            "status": self.status.value,
            "error": self.error,
        }


@dataclass
class AgentTask:
    task_id: str
    presentation_id: str
    query: str
    state: AgentState
    analysis: dict[str, Any]
    plan: list[PlanStep]
    clarifying_questions: list[Question]
    user_answers: dict[str, str]
    executed_steps: list[PlanStep]
    verification_results: dict[str, Any]
    snapshot_path: str | None
    opened_presentation_ids: list[str]
    context: dict[str, Any]
    created_at: datetime = field(default_factory=lambda: datetime.now(UTC))
    updated_at: datetime = field(default_factory=lambda: datetime.now(UTC))
    error: str | None = None
