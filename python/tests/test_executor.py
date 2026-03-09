from unittest.mock import MagicMock

import pytest

from errors import BridgeError
from orchestrator.executor import PlanExecutor
from orchestrator.models import AgentState, AgentTask, PlanStep, StepStatus


def _make_executor(method_map=None):
    engine = MagicMock()
    method_map = method_map or {"pptx_get_slide": "get_slide"}
    return PlanExecutor(engine, method_map), engine


def _make_task(context=None):
    return AgentTask(
        task_id="t1",
        presentation_id="p1",
        query="test",
        state=AgentState.EXECUTING,
        analysis={},
        plan=[],
        clarifying_questions=[],
        user_answers={},
        executed_steps=[],
        verification_results={},
        snapshot_path=None,
        opened_presentation_ids=[],
        context=context or {"presentation_id": "p1"},
    )


def test_resolves_variable_references():
    executor, engine = _make_executor({"pptx_open_presentation": "open_presentation"})
    engine.open_presentation.return_value = {"presentation_id": "new-id"}
    step = PlanStep(
        "s001",
        "Open template",
        "pptx_open_presentation",
        {"file_path": "/tmp/template.pptx"},
        captures="template_id",
    )
    task = _make_task()
    executor.execute(task, [step])
    assert task.context["template_id"] == "new-id"
    assert step.status == StepStatus.DONE


def test_variable_substitution_in_params():
    executor, engine = _make_executor({"pptx_get_slide": "get_slide"})
    engine.get_slide.return_value = {}
    step = PlanStep(
        "s001",
        "Get slide",
        "pptx_get_slide",
        {"presentation_id": "$template_id", "slide_index": 1},
    )
    task = _make_task({"presentation_id": "p1", "template_id": "p2"})
    executor.execute(task, [step])
    engine.get_slide.assert_called_once_with({"presentation_id": "p2", "slide_index": 1})


def test_raises_on_unknown_variable():
    executor, engine = _make_executor()
    step = PlanStep("s001", "Desc", "pptx_get_slide", {"presentation_id": "$nonexistent"})
    task = _make_task()
    with pytest.raises(BridgeError) as exc_info:
        executor.execute(task, [step])
    assert "nonexistent" in exc_info.value.message


def test_marks_step_failed_on_engine_error():
    executor, engine = _make_executor({"pptx_get_slide": "get_slide"})
    engine.get_slide.side_effect = BridgeError(code="not_found", message="Slide not found")
    step = PlanStep("s001", "Desc", "pptx_get_slide", {"presentation_id": "p1", "slide_index": 99})
    task = _make_task()
    with pytest.raises(BridgeError):
        executor.execute(task, [step])
    assert step.status == StepStatus.FAILED
    assert "not_found" in step.error
