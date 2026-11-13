from __future__ import annotations

from unittest.mock import MagicMock

import pytest

from errors import BridgeError
from orchestrator.config import DEFAULT_MODELS, AgentConfig, agent_enabled
from orchestrator.executor import PlanExecutor
from orchestrator.models import AgentState, AgentTask, PlanStep, StepStatus
from service import PowerPointService

_LLM_ENV = ("PPTX_ENABLE_AGENT", "PPTX_LLM_PROVIDER", "PPTX_LLM_API_KEY", "PPTX_LLM_MODEL", "ANTHROPIC_API_KEY")


@pytest.fixture(autouse=True)
def _clean_llm_env(monkeypatch):
    for name in _LLM_ENV:
        monkeypatch.delenv(name, raising=False)


def _agent_error(monkeypatch, **env):
    for name, value in env.items():
        monkeypatch.setenv(name, value)
    service = PowerPointService()
    try:
        with pytest.raises(BridgeError) as excinfo:
            service.dispatch("pptx_agent_status", {"task_id": "t"})
        return service.orchestrator, excinfo.value
    finally:
        service.shutdown()


def test_an_api_key_alone_does_not_switch_the_agent_on(monkeypatch):
    orchestrator, error = _agent_error(monkeypatch, ANTHROPIC_API_KEY="sk-test")

    assert orchestrator is None
    assert error.code == "dependency_missing"
    assert "PPTX_ENABLE_AGENT=1" in error.message


def test_enabling_the_agent_without_a_key_says_which_key_is_missing(monkeypatch):
    orchestrator, error = _agent_error(monkeypatch, PPTX_ENABLE_AGENT="1")

    assert orchestrator is None
    assert "ANTHROPIC_API_KEY" in error.message


@pytest.mark.parametrize(
    ("value", "expected"), [("1", True), ("true", True), (" On ", True), ("0", False), ("", False)]
)
def test_agent_switch_values(monkeypatch, value, expected):
    monkeypatch.setenv("PPTX_ENABLE_AGENT", value)

    assert agent_enabled() is expected


def test_default_model_is_current_and_overridable(monkeypatch):
    monkeypatch.setenv("ANTHROPIC_API_KEY", "sk-test")
    assert AgentConfig.from_env().model == DEFAULT_MODELS["anthropic"] == "claude-sonnet-5-5"

    monkeypatch.setenv("PPTX_LLM_MODEL", "claude-opus-5-5")
    assert AgentConfig.from_env().model == "claude-opus-5-5"


def _task():
    return AgentTask(
        task_id="t1",
        presentation_id="p1",
        query="q",
        state=AgentState.EXECUTING,
        analysis={},
        plan=[],
        clarifying_questions=[],
        user_answers={},
        executed_steps=[],
        verification_results={},
        snapshot_path=None,
        opened_presentation_ids=[],
        context={"presentation_id": "p1"},
    )


def test_plan_execution_stops_when_it_exceeds_its_time_budget(monkeypatch):
    clock = iter([0.0, 0.0, 5.0])
    monkeypatch.setattr("orchestrator.executor.time.monotonic", lambda: next(clock))
    engine = MagicMock()
    engine.get_engine_info.return_value = {}
    executor = PlanExecutor(engine, {"pptx_get_engine_info": "get_engine_info"}, timeout_sec=1)
    steps = [PlanStep(f"s{i}", "info", "pptx_get_engine_info", {}) for i in range(2)]

    with pytest.raises(BridgeError) as excinfo:
        executor.execute(_task(), steps)

    assert excinfo.value.code == "timeout"
    assert steps[0].status == StepStatus.DONE
    assert steps[1].status == StepStatus.FAILED
    assert engine.get_engine_info.call_count == 1


def test_plan_steps_with_unknown_parameters_are_rejected():
    engine = MagicMock()
    executor = PlanExecutor(engine, {"pptx_get_slide": "get_slide"})
    step = PlanStep("s1", "read", "pptx_get_slide", {"presentation_id": "p1", "slide_index": 1, "verbose": True})

    with pytest.raises(BridgeError) as excinfo:
        executor.execute(_task(), [step])

    assert excinfo.value.code == "validation_error"
    engine.get_slide.assert_not_called()
