from __future__ import annotations

import io
import json
import threading
import urllib.error
import urllib.request

import pytest

from word_agent import WordDocument, WordPackage
from word_agent.ui.server import make_server

from conftest import ScriptedLLM, assistant, call

HEADERS = {"X-Word-Agent": "1"}


@pytest.fixture
def ui():
    llm = ScriptedLLM([])
    server = make_server(port=0, llm_factory=lambda: llm)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base = f"http://127.0.0.1:{server.server_address[1]}"

    def request(path, data=None, headers=HEADERS):
        req = urllib.request.Request(base + path, data=data, headers=headers if data is not None else {})
        with urllib.request.urlopen(req, timeout=10) as response:
            return response.read()

    def post_stream(path, payload=None):
        body = json.dumps(payload).encode() if payload is not None else b""
        return [json.loads(line) for line in request(path, body).decode().splitlines()]

    yield llm, request, post_stream
    server.shutdown()
    server.server_close()


def paragraphs(*texts):
    return call("insert_paragraphs", paragraphs=[{"text": t} for t in texts])


def test_page_and_empty_state(ui):
    _, request, _ = ui
    assert b"word-agent" in request("/")
    assert json.loads(request("/state")) == {"name": "document.docx", "blocks": [], "pending": None}


def test_chat_streams_events_and_keeps_conversation(ui):
    llm, request, post_stream = ui
    llm.turns = [assistant(paragraphs("Hello")), assistant(content="Added a greeting.")]
    events = post_stream("/chat", {"message": "say hello"})
    assert [e["type"] for e in events] == ["tool", "reply", "state"]
    assert events[0]["name"] == "insert_paragraphs"
    assert events[1]["text"] == "Added a greeting."
    assert [b["text"] for b in events[2]["blocks"]] == ["Hello"]

    # second turn continues the same conversation and the same document
    llm.turns = [assistant(call("replace_text", find="Hello", replace="Goodbye")), assistant(content="Changed.")]
    events = post_stream("/chat", {"message": "make it goodbye"})
    assert [b["text"] for b in events[-1]["blocks"]] == ["Goodbye"]
    second_turn = llm.requests[2]["messages"]
    assert [m["role"] for m in second_turn] == ["system", "user", "assistant", "tool", "assistant", "user"]
    assert second_turn[1]["content"] == "say hello" and second_turn[-1]["content"] == "make it goodbye"

    downloaded = WordDocument(WordPackage.from_bytes(request("/download")))
    assert downloaded.text() == "Goodbye"


def test_plan_feedback_approve_and_discard(ui):
    llm, request, post_stream = ui
    llm.turns = [assistant(content="1. Add one line.")]
    events = post_stream("/chat", {"message": "write", "plan": True})
    assert [e["type"] for e in events] == ["plan", "state"]
    assert events[-1]["pending"] == "1. Add one line." and events[-1]["blocks"] == []

    # a message while a plan is pending is feedback on the plan
    llm.turns = [assistant(content="1. Add two lines.")]
    events = post_stream("/chat", {"message": "two lines please"})
    assert events[0] == {"type": "plan", "text": "1. Add two lines."}
    assert events[-1]["blocks"] == []

    llm.turns = [assistant(paragraphs("a", "b")), assistant(content="Done.")]
    events = post_stream("/approve")
    assert [e["type"] for e in events] == ["tool", "reply", "state"]
    assert [b["text"] for b in events[-1]["blocks"]] == ["a", "b"] and events[-1]["pending"] is None

    llm.turns = [assistant(content="1. Delete it all.")]
    post_stream("/chat", {"message": "wipe", "plan": True})
    state = json.loads(request("/discard", b""))
    assert state["pending"] is None and len(state["blocks"]) == 2
    assert post_stream("/approve")[0]["type"] == "error"


def test_errors_are_reported_in_the_stream(ui):
    _, _, post_stream = ui  # the scripted model has no turns, so the call fails
    events = post_stream("/chat", {"message": "hi"})
    assert events[0]["type"] == "error" and events[-1]["type"] == "state"


def test_upload_and_reset(ui):
    _, request, _ = ui
    source = WordDocument.open()
    source.insert_paragraph("From upload", "Heading 1")
    state = json.loads(request("/upload", source.package.to_bytes(), {**HEADERS, "X-Filename": "my%20file.docx"}))
    assert state["name"] == "my file.docx"
    assert state["blocks"][0]["text"] == "From upload" and state["blocks"][0]["style"] == "heading 1"

    with pytest.raises(urllib.error.HTTPError) as bad:
        request("/upload", b"not a docx")
    assert bad.value.code == 400
    assert json.loads(request("/state"))["name"] == "my file.docx"   # a bad upload leaves the session alone

    assert json.loads(request("/reset", b""))["blocks"] == []


def test_cross_site_requests_are_refused(ui):
    _, request, _ = ui
    with pytest.raises(urllib.error.HTTPError) as missing_header:
        request("/reset", b"", headers={})
    assert missing_header.value.code == 403
    with pytest.raises(urllib.error.HTTPError) as wrong_host:
        request("/reset", b"", headers={**HEADERS, "Host": "evil.example"})
    assert wrong_host.value.code == 403
