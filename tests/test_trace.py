from __future__ import annotations

import json

from docx_harness import JsonlTrace, WordDocument, cli, create_agent

from conftest import ScriptedLLM, assistant, call


def read(path):
    return [json.loads(line) for line in path.read_text().splitlines()]


def test_trace_records_a_run(tmp_path):
    path = tmp_path / "nested" / "run.jsonl"
    trace = JsonlTrace(path)
    seen = []
    llm = ScriptedLLM([
        assistant(call("insert_paragraphs", paragraphs=[{"text": "Hello"}])),
        assistant(content="Added a greeting."),
    ])
    agent = create_agent(WordDocument.open(), llm=llm, on_event=trace.chain(seen.append))
    run = trace.start("say hello", model="scripted")
    result = agent.run("say hello")
    trace.end(result, {"prompt_tokens": 10, "completion_tokens": 2, "requests": 2})

    records = read(path)
    assert [r["type"] for r in records] == ["run_start", "assistant", "tool_call", "tool_result", "assistant", "run_end"]
    assert {r["run"] for r in records} == {run}
    assert all(isinstance(r["elapsed_ms"], int) and r["ts"] for r in records)
    assert records[0]["instruction"] == "say hello" and records[0]["model"] == "scripted"
    assert records[2]["name"] == records[3]["name"] == "insert_paragraphs"
    assert records[-1]["status"] == "completed" and records[-1]["steps"] == 2
    assert records[-1]["final_text"] == "Added a greeting." and records[-1]["usage"]["requests"] == 2
    assert len(seen) == 4  # the chained callback still gets every event

    # a second run appends to the same file under a new id
    trace.start("again")
    trace.end(status="error", error="boom")
    second = read(path)[len(records):]
    assert [r["type"] for r in second] == ["run_start", "run_end"] and second[0]["run"] != run
    assert second[1]["error"] == "boom"


def test_cli_trace(tmp_path, monkeypatch):
    llm = ScriptedLLM([assistant(content="Nothing to do.")])
    monkeypatch.setattr(cli, "OpenRouterClient", lambda model=None: llm)
    path = tmp_path / "t.jsonl"
    assert cli.main(["noop", "-o", str(tmp_path / "out.docx"), "--trace", str(path)]) == 0
    records = read(path)
    assert [r["type"] for r in records] == ["run_start", "assistant", "run_end"]
    assert records[0]["source"] == "cli" and records[0]["instruction"] == "noop"
