from __future__ import annotations

import asyncio

import pytest

from docx_harness import Agent, ToolRegistry, WordDocument, adraft, adraft_many, create_agent, tool

from conftest import ScriptedLLM, assistant, call


def paragraphs(*texts):
    return call("insert_paragraphs", paragraphs=[{"text": t} for t in texts])


async def test_loop_runs_tools_until_text_reply(doc):
    llm = ScriptedLLM([
        assistant(call("get_outline")),
        assistant(paragraphs("Hello", "World")),
        assistant(content="Added two paragraphs."),
    ])
    events = []
    agent = create_agent(doc, llm=llm, on_event=events.append)
    result = await agent.arun("write hello world")

    assert result.status == "completed"
    assert result.final_text == "Added two paragraphs."
    assert result.steps == 3
    assert doc.text() == "Hello\nWorld"
    assert [m["role"] for m in result.messages] == ["system", "user", "assistant", "tool", "assistant", "tool", "assistant"]
    assert result.messages[3]["tool_call_id"] == llm.requests[1]["messages"][2]["tool_calls"][0]["id"]
    assert [e.type for e in events].count("tool_call") == 2
    # every available tool is offered outside plan mode (a blank document has
    # no tables and tracks no changes, so those groups are left out)
    assert len(llm.requests[0]["tools"]) == len(agent.tools.to_openai_schema()) < len(agent.tools)


async def test_tool_errors_go_back_to_the_model(doc):
    llm = ScriptedLLM([
        assistant(call("update_paragraph", block_id="b42", text="x")),
        assistant(content="Could not find it."),
    ])
    result = await create_agent(doc, llm=llm).arun("edit")
    assert result.messages[3]["content"].startswith("Error: No block with id 'b42'")
    assert result.status == "completed"


async def test_max_steps(doc):
    llm = ScriptedLLM([assistant(call("get_outline"), content="looking")] * 3)
    result = await create_agent(doc, llm=llm, max_steps=3).arun("loop forever")
    assert result.status == "max_steps"
    assert result.steps == 3
    assert result.final_text == "looking"


async def test_plan_mode_is_read_only_then_executes(doc):
    llm = ScriptedLLM([
        # planning: one allowed read, one forbidden write
        assistant(call("get_outline"), paragraphs("sneaky")),
        assistant(content="1. Add a greeting paragraph."),
        # execution
        assistant(paragraphs("Greetings")),
        assistant(content="Done."),
    ])
    seen = []

    def approve(plan):
        seen.append((plan.text, doc.text()))
        return True

    result = await create_agent(doc, llm=llm).arun("greet", plan_mode=True, approve=approve)

    # at approval time nothing had been written
    assert seen == [("1. Add a greeting paragraph.", "")]
    assert doc.text() == "Greetings"
    assert result.status == "completed"
    assert result.plan.text == "1. Add a greeting paragraph."
    assert result.steps == 4

    plan_request, _, execute_request, _ = llm.requests
    plan_tools = {t["function"]["name"] for t in plan_request["tools"]}
    assert "insert_paragraphs" not in plan_tools and "read_blocks" in plan_tools
    assert "PLAN MODE" in plan_request["messages"][0]["content"]
    assert "PLAN MODE" not in execute_request["messages"][0]["content"]
    assert "insert_paragraphs" in {t["function"]["name"] for t in execute_request["tools"]}
    # the blocked call was answered with an error, and planning context is carried over
    tool_messages = [m for m in execute_request["messages"] if m["role"] == "tool"]
    assert any("not available in plan mode" in m["content"] for m in tool_messages)
    assert execute_request["messages"][-1]["content"].startswith("The plan is approved")


async def test_plan_rejected_changes_nothing(doc, tmp_path):
    llm = ScriptedLLM([assistant(content="1. Delete everything.")])
    result = await create_agent(doc, llm=llm).arun("x", plan_mode=True, approve=lambda plan: False)
    assert result.status == "rejected"
    assert llm.turns == [] and len(llm.requests) == 1

    out = tmp_path / "never.docx"
    rejected = await adraft("x", output_path=out, plan=True, approve=lambda plan: False,
                            llm=ScriptedLLM([assistant(content="plan")]))
    assert rejected.status == "rejected"
    assert not out.exists()


async def test_plan_revision_with_async_approver(doc):
    llm = ScriptedLLM([
        assistant(content="1. Add one paragraph."),
        assistant(content="1. Add two paragraphs."),
        assistant(paragraphs("a", "b")),
        assistant(content="Done."),
    ])
    decisions = iter(["make it two paragraphs", True])

    async def approve(plan):
        return next(decisions)

    result = await create_agent(doc, llm=llm).arun("write", plan_mode=True, approve=approve)
    assert result.plan.text == "1. Add two paragraphs."
    assert doc.text() == "a\nb"
    revise_request = llm.requests[1]["messages"]
    assert "make it two paragraphs" in revise_request[-1]["content"]
    assert revise_request[-2]["content"] == "1. Add one paragraph."


async def test_plan_auto_approved_without_callback(doc):
    llm = ScriptedLLM([assistant(content="plan"), assistant(paragraphs("x")), assistant(content="ok")])
    result = await create_agent(doc, llm=llm).arun("write", plan_mode=True)
    assert result.status == "completed" and doc.text() == "x"


async def test_read_only_calls_run_concurrently_and_writes_in_order():
    running = 0
    peak = 0
    log = []

    @tool(read_only=True)
    async def slow_read(n: int) -> str:
        """Read."""
        nonlocal running, peak
        running += 1
        peak = max(peak, running)
        await asyncio.sleep(0.02)
        running -= 1
        return f"read {n}"

    @tool
    async def slow_write(n: int) -> str:
        """Write."""
        nonlocal running, peak
        running += 1
        peak = max(peak, running)
        await asyncio.sleep(0.01 * (3 - n))  # later calls finish faster if run concurrently
        running -= 1
        log.append(n)
        return f"wrote {n}"

    reads = [call("slow_read", _id=f"r{n}", n=n) for n in range(3)]
    writes = [call("slow_write", _id=f"w{n}", n=n) for n in range(3)]

    llm = ScriptedLLM([assistant(*reads), assistant(content="done")])
    result = await Agent(llm, ToolRegistry([slow_read, slow_write])).arun("go")
    assert peak == 3
    assert [m["content"] for m in result.messages if m["role"] == "tool"] == ["read 0", "read 1", "read 2"]

    peak = 0
    llm = ScriptedLLM([assistant(*writes), assistant(content="done")])
    result = await Agent(llm, ToolRegistry([slow_read, slow_write])).arun("go")
    assert peak == 1
    assert log == [0, 1, 2]
    assert [m["tool_call_id"] for m in result.messages if m["role"] == "tool"] == ["w0", "w1", "w2"]


def test_sync_wrappers(doc):
    llm = ScriptedLLM([assistant(content="plan"), assistant(paragraphs("x")), assistant(content="ok")])
    agent = create_agent(doc, llm=llm)
    plan = agent.plan("write")           # each sync call uses its own event loop
    result = agent.execute(plan)
    assert result.final_text == "ok" and doc.text() == "x"


async def test_sync_wrapper_inside_event_loop_explains_itself(doc):
    agent = create_agent(doc, llm=ScriptedLLM([]))
    with pytest.raises(RuntimeError, match="async methods"):
        agent.run("x")


async def test_adraft_edits_existing_file_and_adraft_many(tmp_path):
    source = tmp_path / "in.docx"
    original = WordDocument.open()
    original.insert_paragraph("Old text")
    original.save(source)

    out = tmp_path / "out.docx"
    llm = ScriptedLLM([assistant(call("replace_text", find="Old", replace="New")), assistant(content="ok")])
    await adraft("update", source, out, llm=llm)
    assert WordDocument.open(out).text() == "New text"
    assert WordDocument.open(source).text() == "Old text"

    jobs = [
        {"instruction": "a", "output_path": tmp_path / "a.docx",
         "llm": ScriptedLLM([assistant(paragraphs("A")), assistant(content="ok")])},
        {"instruction": "b", "output_path": tmp_path / "b.docx",
         "llm": ScriptedLLM([])},  # fails: no turns
        {"instruction": "c", "output_path": tmp_path / "c.docx",
         "llm": ScriptedLLM([assistant(paragraphs("C")), assistant(content="ok")])},
    ]
    results = await adraft_many(jobs, concurrency=2)
    assert results[0].status == "completed" and results[2].status == "completed"
    assert isinstance(results[1], AssertionError)
    assert WordDocument.open(tmp_path / "a.docx").text() == "A"
    assert WordDocument.open(tmp_path / "c.docx").text() == "C"


async def test_duplicate_reads_are_short_circuited():
    hits = []

    @tool(read_only=True)
    def peek(n: int = 0) -> str:
        """Peek."""
        hits.append(n)
        return f"value {n}"

    @tool
    def poke() -> str:
        """Poke."""
        return "poked"

    llm = ScriptedLLM([
        assistant(call("peek", _id="a", n=1)),
        assistant(call("peek", _id="b", n=1), call("peek", _id="c", n=2)),   # first is a repeat
        assistant(call("poke", _id="d")),
        assistant(call("peek", _id="e", n=1)),                                # fresh again after a write
        assistant(content="done"),
    ])
    result = await Agent(llm, ToolRegistry([peek, poke])).arun("go")
    tool_results = {m["tool_call_id"]: m["content"] for m in result.messages if m["role"] == "tool"}
    assert hits == [1, 2, 1]
    assert tool_results["a"] == "value 1" and tool_results["e"] == "value 1"
    assert "already made this exact call" in tool_results["b"]


async def test_planning_step_budget_forces_a_plan(doc):
    llm = ScriptedLLM(
        [assistant(call("read_blocks", count=n)) for n in (1, 2)]
        + [assistant(content="1. The plan."), assistant(content="executed")]
    )
    agent = create_agent(doc, llm=llm)
    agent.max_plan_steps = 3
    result = await agent.arun("x", plan_mode=True)
    assert result.plan.text == "1. The plan."
    assert llm.requests[0]["tools"] and llm.requests[1]["tools"]
    assert llm.requests[2]["tools"] is None          # last planning step: no tools offered
    assert llm.requests[3]["tools"]                  # execution has tools again


async def test_tools_offered_follow_the_state_step_by_step(doc):
    llm = ScriptedLLM([
        assistant(call("insert_table", rows=[["a", "b"]])),
        assistant(content="done"),
    ])
    await create_agent(doc, llm=llm).arun("add a table")
    names = [{t["function"]["name"] for t in request["tools"]} for request in llm.requests]
    assert "set_table_cell" not in names[0] and "set_table_cell" in names[1]


def test_format_usage():
    from docx_harness.agent import format_usage

    usage = {"prompt_tokens": 48_210, "completion_tokens": 812, "requests": 12, "cost": 0.00423}
    assert format_usage(usage, 12) == "12 steps · 48.2k in / 812 out · $0.0042"
    assert format_usage({**usage, "cost": 2.5}, 1) == "1 step · 48.2k in / 812 out · $2.50"
    del usage["cost"]
    assert format_usage(usage) == "48.2k in / 812 out"
    assert format_usage(None) == format_usage({"prompt_tokens": 0, "completion_tokens": 0, "requests": 0}) == ""


# ---------------------------------------------------------------- retries


def _fake_client(replies, **kwargs):
    from types import SimpleNamespace

    from docx_harness import OpenAICompatibleClient

    calls = []

    async def create(**request):
        calls.append(request)
        return replies.pop(0)

    client = OpenAICompatibleClient("some/model", api_key="k", **kwargs)
    client._get_client = lambda: SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)))
    return client, calls


def _ok():
    from types import SimpleNamespace

    message = SimpleNamespace(content="hi", tool_calls=None)
    usage = SimpleNamespace(prompt_tokens=10, completion_tokens=5)
    return SimpleNamespace(usage=usage, choices=[SimpleNamespace(message=message)])


def _failed(code=502, message="Upstream provider error"):
    from types import SimpleNamespace

    return SimpleNamespace(usage=None, choices=None, error={"code": code, "message": message})


@pytest.fixture
def sleeps(monkeypatch):
    import docx_harness.agent.llm as llm_module

    delays = []

    async def fake_sleep(seconds):
        delays.append(seconds)

    monkeypatch.setattr(llm_module.asyncio, "sleep", fake_sleep)
    return delays


async def test_reply_without_choices_is_retried(sleeps):
    client, calls = _fake_client([_failed(), _failed(), _ok()])
    reply = await client.acomplete([{"role": "user", "content": "x"}])
    assert reply == {"role": "assistant", "content": "hi"}
    assert len(calls) == 3
    assert len(sleeps) == 2 and sleeps[1] > sleeps[0]
    assert client.usage == {"prompt_tokens": 10, "completion_tokens": 5, "requests": 1}


async def test_retries_give_up_with_provider_message(sleeps):
    from docx_harness import LLMError

    client, calls = _fake_client([_failed() for _ in range(3)], max_retries=2)
    with pytest.raises(LLMError, match="Upstream provider error") as info:
        await client.acomplete([{"role": "user", "content": "x"}])
    assert info.value.code == 502
    assert len(calls) == 3 and len(sleeps) == 2


async def test_permanent_error_is_not_retried(sleeps):
    from docx_harness import LLMError

    client, calls = _fake_client([_failed(401, "No auth credentials found")])
    with pytest.raises(LLMError, match="No auth credentials"):
        await client.acomplete([{"role": "user", "content": "x"}])
    assert len(calls) == 1 and sleeps == []


async def test_retry_and_timeout_settings_reach_the_sdk(monkeypatch):
    import openai

    from docx_harness import OpenAICompatibleClient

    seen = {}
    monkeypatch.setattr(openai, "AsyncOpenAI", lambda **kwargs: seen.update(kwargs) or object())
    OpenAICompatibleClient("m", api_key="k", max_retries=7, timeout=30)._get_client()
    assert seen["max_retries"] == 7 and seen["timeout"] == 30
    seen.clear()
    OpenAICompatibleClient("m", api_key="k")._get_client()
    assert seen["max_retries"] == 4 and seen["timeout"] == 120.0
