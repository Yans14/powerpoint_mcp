from orchestrator.models import AgentState, AgentTask, PlanStep, Question, StepStatus


def _make_task(**kwargs) -> AgentTask:
    defaults = dict(
        task_id="task-abc",
        presentation_id="pres-uuid",
        query="convert branding",
        state=AgentState.CLARIFYING,
        analysis={},
        plan=[],
        clarifying_questions=[],
        user_answers={},
        executed_steps=[],
        verification_results={},
        snapshot_path=None,
        opened_presentation_ids=[],
        context={"presentation_id": "pres-uuid"},
    )
    return AgentTask(**{**defaults, **kwargs})


def test_agent_state_values():
    assert AgentState.ANALYZING.value == "analyzing"
    assert AgentState.COMPLETE.value == "complete"
    assert AgentState.FAILED.value == "failed"


def test_question_to_dict():
    q = Question("q1", "What template?", "template", ["A: New", "B: Old"], required=True)
    d = q.to_dict()
    assert d["question_id"] == "q1"
    assert d["choices"] == ["A: New", "B: Old"]


def test_plan_step_default_status():
    step = PlanStep("s001", "Open file", "pptx_open_presentation", {"file_path": "/tmp/a.pptx"})
    assert step.status == StepStatus.PENDING
    assert step.result is None
    assert step.error is None


def test_plan_step_to_dict_includes_status():
    step = PlanStep("s001", "Desc", "pptx_get_slide", {})
    step.status = StepStatus.DONE
    d = step.to_dict()
    assert d["status"] == "done"


def test_task_construction():
    task = _make_task()
    assert task.state == AgentState.CLARIFYING
    assert task.context["presentation_id"] == "pres-uuid"
