from __future__ import annotations

import asyncio

import pytest

from word_agent import Agent, ToolRegistry, WordDocument, adraft, adraft_many, create_agent, tool

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
    # every tool is offered outside plan mode
    assert len(llm.requests[0]["tools"]) == len(agent.tools)


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
