from __future__ import annotations

import uuid
from datetime import UTC, datetime
from typing import Any

from engines.base import BaseEngine
from errors import BridgeError
from orchestrator.clarifier import LLMClarifier
from orchestrator.config import AgentConfig
from orchestrator.executor import PlanExecutor
from orchestrator.llm_client import LLMClient
from orchestrator.models import AgentState, AgentTask, StepStatus
from orchestrator.planner import LLMPlanner
from orchestrator.snapshot import SnapshotManager


class AgentOrchestrator:
    def __init__(
        self,
        engine: BaseEngine,
        config: AgentConfig,
        method_map: dict[str, str],
    ) -> None:
        self.engine = engine
        self.config = config
        self.tasks: dict[str, AgentTask] = {}
        llm = LLMClient(config)
        self.clarifier = LLMClarifier(llm)
        self.planner = LLMPlanner(llm, max_steps=config.max_steps)
        self.executor = PlanExecutor(engine, method_map)
        self.snapshot_mgr = SnapshotManager()

    def dispatch(self, method: str, params: dict[str, Any]) -> Any:
        handlers = {
            "pptx_agent_start": self._agent_start,
            "pptx_agent_respond": self._agent_respond,
            "pptx_agent_execute": self._agent_execute,
            "pptx_agent_status": self._agent_status,
            "pptx_agent_rollback": self._agent_rollback,
            "pptx_agent_cancel": self._agent_cancel,
        }
        handler = handlers.get(method)
        if handler is None:
            raise BridgeError(code="not_found", message=f"Unknown agent method: {method}")
        return handler(params)

    def _agent_start(self, params: dict[str, Any]) -> dict[str, Any]:
        presentation_id = str(params["presentation_id"])
        query = str(params["query"])
        skip_questions = bool(params.get("skip_questions", False))

        prs_state = self.engine.get_presentation_state({"presentation_id": presentation_id})
        theme = self.engine.get_theme({"presentation_id": presentation_id})
        layouts = self.engine.get_layouts({"presentation_id": presentation_id})

        analysis, questions = self.clarifier.analyze(query, prs_state, theme, layouts)

        if skip_questions:
            questions = []

        task_id = str(uuid.uuid4())
        task = AgentTask(
            task_id=task_id,
            presentation_id=presentation_id,
            query=query,
            state=AgentState.CLARIFYING if questions else AgentState.PLANNING,
            analysis=analysis,
            plan=[],
            clarifying_questions=questions,
            user_answers={},
            executed_steps=[],
            verification_results={},
            snapshot_path=None,
            opened_presentation_ids=[],
            context={"presentation_id": presentation_id},
        )

        if not questions:
            task.plan = self.planner.generate_plan(
                query=query,
                answers={},
                prs_state=prs_state,
                theme=theme,
            )
            task.state = AgentState.READY

        self.tasks[task_id] = task

        return {
            "task_id": task_id,
            "state": task.state.value,
            "analysis": analysis,
            "questions": [q.to_dict() for q in questions],
            "plan_preview": analysis.get("plan_preview", ""),
            "estimated_steps": analysis.get("estimated_steps", 0),
        }

    def _agent_respond(self, params: dict[str, Any]) -> dict[str, Any]:
        task = self._get_task(str(params["task_id"]))
        if task.state != AgentState.CLARIFYING:
            raise BridgeError(
                code="conflict",
                message=f"Task is in state '{task.state.value}', not 'clarifying'.",
            )

        for answer in params["answers"]:
            task.user_answers[str(answer["question_id"])] = str(answer["answer"])

        required_ids = {q.question_id for q in task.clarifying_questions if q.required}
        unanswered = required_ids - set(task.user_answers.keys())

        if unanswered:
            remaining = [q.to_dict() for q in task.clarifying_questions if q.question_id in unanswered]
            return {
                "task_id": task.task_id,
                "state": "clarifying",
                "remaining_questions": remaining,
            }

        task.state = AgentState.PLANNING
        prs_state = self.engine.get_presentation_state({"presentation_id": task.presentation_id})
        theme = self.engine.get_theme({"presentation_id": task.presentation_id})

        task.plan = self.planner.generate_plan(
            query=task.query,
            answers=task.user_answers,
            prs_state=prs_state,
            theme=theme,
        )
        task.state = AgentState.READY
        task.updated_at = datetime.now(UTC)

        return {
            "task_id": task.task_id,
            "state": "ready",
            "plan_preview": task.analysis.get("plan_preview", ""),
            "step_count": len(task.plan),
            "steps": [s.to_dict() for s in task.plan],
        }

    def _agent_execute(self, params: dict[str, Any]) -> dict[str, Any]:
        task = self._get_task(str(params["task_id"]))

        if task.state != AgentState.READY:
            raise BridgeError(
                code="conflict",
                message=f"Task must be in READY state to execute. Current: {task.state.value}",
            )

        session = self.engine.sessions[task.presentation_id]
        snapshot_path = self.snapshot_mgr.create(session)
        task.snapshot_path = snapshot_path
        task.state = AgentState.EXECUTING
        task.updated_at = datetime.now(UTC)

        try:
            self.executor.execute(task, task.plan)
            task.executed_steps = task.plan[:]
        except BridgeError:
            task.state = AgentState.FAILED
            task.error = "Step failure during execution"
            return {
                "task_id": task.task_id,
                "state": "failed",
                "error": task.error,
                "steps_executed": sum(1 for s in task.plan if s.status == StepStatus.DONE),
                "steps_total": len(task.plan),
                "execution_log": [s.to_dict() for s in task.plan],
                "rollback_available": True,
            }

        task.state = AgentState.VERIFYING
        verification = self._run_verification(task)
        task.verification_results = verification
        task.state = AgentState.COMPLETE
        task.updated_at = datetime.now(UTC)

        return {
            "task_id": task.task_id,
            "state": "complete",
            "steps_executed": sum(1 for s in task.plan if s.status == StepStatus.DONE),
            "steps_total": len(task.plan),
            "execution_log": [s.to_dict() for s in task.plan],
            "verification": verification,
            "rollback_available": task.snapshot_path is not None,
            "summary": task.analysis.get("plan_preview", "Transformation complete."),
        }

    def _run_verification(self, task: AgentTask) -> dict[str, Any]:
        try:
            final_state = self.engine.get_presentation_state({"presentation_id": task.presentation_id})
        except Exception:
            final_state = {}

        try:
            from checkers.content_checker import ContentChecker
            from checkers.position_checker import PositionChecker

            prs = self.engine.sessions[task.presentation_id].extra["prs"]
            slide_count = final_state.get("slide_count", 0)
            slide_indices = list(range(1, slide_count + 1))

            pos_result = PositionChecker().check(
                prs, slide_indices, check_overlaps=True, check_bounds=True, check_alignment=False, tolerance_px=5
            )
            content_result = ContentChecker().check(prs, check_empty=True, check_default_text=True)
        except Exception as exc:
            pos_result = {"error": str(exc)}
            content_result = {"error": str(exc)}

        return {
            "final_state": final_state,
            "position_check": pos_result,
            "content_check": content_result,
        }

    def _agent_status(self, params: dict[str, Any]) -> dict[str, Any]:
        task = self._get_task(str(params["task_id"]))
        done = sum(1 for s in task.plan if s.status == StepStatus.DONE)
        return {
            "task_id": task.task_id,
            "state": task.state.value,
            "presentation_id": task.presentation_id,
            "query": task.query,
            "progress": {"current_step": done, "total_steps": len(task.plan)},
            "error": task.error,
            "created_at": task.created_at.isoformat(),
            "updated_at": task.updated_at.isoformat(),
            "rollback_available": task.snapshot_path is not None,
        }

    def _agent_rollback(self, params: dict[str, Any]) -> dict[str, Any]:
        task = self._get_task(str(params["task_id"]))
        if task.snapshot_path is None:
            raise BridgeError(
                code="conflict",
                message="No snapshot available for this task. Execute the plan first.",
            )
        session = self.engine.sessions[task.presentation_id]
        self.snapshot_mgr.restore(session, task.snapshot_path)

        for pid in task.opened_presentation_ids:
            try:
                self.engine.close_presentation({"presentation_id": pid})
            except Exception:
                pass

        task.state = AgentState.READY
        task.executed_steps = []
        for step in task.plan:
            step.status = StepStatus.PENDING
            step.result = None
            step.error = None
        task.opened_presentation_ids = []
        task.error = None
        task.updated_at = datetime.now(UTC)

        return {
            "task_id": task.task_id,
            "state": "ready",
            "message": "Rollback complete. Presentation restored to pre-execution state.",
        }

    def _agent_cancel(self, params: dict[str, Any]) -> dict[str, Any]:
        task = self._get_task(str(params["task_id"]))
        if task.snapshot_path:
            self.snapshot_mgr.remove(task.snapshot_path)
        for pid in task.opened_presentation_ids:
            try:
                self.engine.close_presentation({"presentation_id": pid})
            except Exception:
                pass
        del self.tasks[task.task_id]
        return {"task_id": task.task_id, "cancelled": True}

    def _get_task(self, task_id: str) -> AgentTask:
        task = self.tasks.get(task_id)
        if task is None:
            raise BridgeError(
                code="not_found",
                message=f"Agent task '{task_id}' not found.",
                details={"task_id": task_id},
            )
        return task
