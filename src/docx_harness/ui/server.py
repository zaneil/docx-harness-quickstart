"""A small local chat UI for trying the agent.

    docx-harness-ui            # then open http://127.0.0.1:8765

Chat on the left, a live preview of the document on the right. Start from the
blank template or upload a .docx, keep chatting to refine it, download the
result. This is a test tool: one in-memory session, standard library only,
bound to localhost.
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import threading
import uuid
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib import resources
from datetime import datetime
from pathlib import Path
from typing import Callable
from urllib.parse import quote, unquote

from ..agent import AgentEvent, LLMClient, OpenRouterClient, format_usage
from ..agent.llm import DEFAULT_MODEL
from ..api import create_agent
from ..document import PackageError, WordDocument, WordPackage
from ..trace import JsonlTrace

MAX_UPLOAD_BYTES = 50 * 1024 * 1024
ALLOWED_HOSTS = {"127.0.0.1", "localhost"}


class Session:
    """The one document and conversation the UI is working on."""

    def __init__(self, llm_factory: Callable[[], LLMClient], trace_dir: str | os.PathLike | None = None):
        self.llm_factory = llm_factory
        self.trace_dir = Path(trace_dir) if trace_dir else None
        self.trace: JsonlTrace | None = None
        self.lock = threading.Lock()
        self.track_changes = False
        self.reset()

    def reset(self, data: bytes | None = None, name: str | None = None) -> None:
        package = WordPackage.from_bytes(data) if data else WordPackage.open()
        self.doc = WordDocument(package, track_changes=self.track_changes)
        self.name = name or "document.docx"
        self.history: list[dict] | None = None
        self.pending = None  # a Plan awaiting approval
        if self.trace_dir:
            # A new file per conversation; it appears with the first message.
            stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
            self.trace = JsonlTrace(self.trace_dir / f"{stamp}_{uuid.uuid4().hex[:6]}.jsonl")

    def state(self) -> dict:
        return {
            "name": self.name,
            "blocks": self.doc.read(count=1_000_000, include_formatting=True, resolve_styles=True)["blocks"],
            "pending": self.pending.text if self.pending else None,
            "track_changes": self.track_changes,
            "revisions": len(self.doc.revisions()),
        }


def make_handler(session: Session, model_label: str):
    page = resources.files("docx_harness.ui").joinpath("chat.html").read_text(encoding="utf-8")

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *args) -> None:  # keep the terminal quiet
            pass

        # -------------------------------------------------------- responses

        def _send(self, status: int, body: bytes, content_type: str, extra: dict | None = None) -> None:
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            for key, value in (extra or {}).items():
                self.send_header(key, value)
            self.end_headers()
            self.wfile.write(body)

        def _json(self, payload: dict, status: int = 200) -> None:
            self._send(status, json.dumps(payload).encode(), "application/json")

        def _stream(self, work) -> None:
            """Run ``work(emit)`` and send each emitted event as one JSON line."""
            self.send_response(200)
            self.send_header("Content-Type", "application/x-ndjson")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()

            def emit(event: dict) -> None:
                try:
                    self.wfile.write((json.dumps(event) + "\n").encode())
                    self.wfile.flush()
                except OSError:
                    pass  # the browser went away; let the work finish

            with session.lock:
                try:
                    asyncio.run(work(emit))
                except Exception as exc:  # noqa: BLE001 - shown in the chat
                    emit({"type": "error", "text": str(exc) or type(exc).__name__})
                    if session.trace and session.trace.run:
                        session.trace.end(status="error", error=str(exc) or type(exc).__name__)
                emit({"type": "state", **session.state()})

        def _agent(self, emit):
            def on_event(event: AgentEvent) -> None:
                if event.type == "tool_call":
                    emit({"type": "tool", "name": event.data["name"], "arguments": event.data.get("arguments") or ""})
                elif event.type == "tool_result" and event.data["result"].startswith("Error"):
                    emit({"type": "tool_error", "text": event.data["result"][:300]})

            if session.trace:
                on_event = session.trace.chain(on_event)
            # Pictures come from the folder the server was started in, nowhere else.
            return create_agent(session.doc, llm=session.llm_factory(), on_event=on_event, image_root=os.getcwd())

        @staticmethod
        def _trace_start(agent, instruction: str, mode: str) -> None:
            if session.trace:
                session.trace.start(
                    instruction, source="ui", mode=mode, model=getattr(agent.llm, "model", None),
                    document=session.name, track_changes=session.track_changes,
                )

        @staticmethod
        def _trace_end(agent, result=None, **extra) -> None:
            if session.trace:
                session.trace.end(result, getattr(agent.llm, "usage", None), **extra)

        # ---------------------------------------------------------- routing

        def _local(self) -> bool:
            # Refuse requests that arrive under another hostname (DNS rebinding).
            return self.headers.get("Host", "").rsplit(":", 1)[0] in ALLOWED_HOSTS

        def do_GET(self) -> None:
            if not self._local():
                return self._json({"error": "forbidden"}, 403)
            if self.path == "/":
                self._send(200, page.replace("{{MODEL}}", model_label).encode(), "text/html; charset=utf-8")
            elif self.path == "/state":
                self._json(session.state())
            elif self.path == "/download":
                with session.lock:
                    data = session.doc.package.to_bytes()
                self._send(
                    200, data,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    {"Content-Disposition": f"attachment; filename*=UTF-8''{quote(session.name)}"},
                )
            else:
                self._json({"error": "not found"}, 404)

        def do_POST(self) -> None:
            # The custom header cannot be sent cross-site without a CORS
            # preflight, which this server never grants.
            if not self._local() or self.headers.get("X-Docx-Harness") != "1":
                return self._json({"error": "forbidden"}, 403)
            length = int(self.headers.get("Content-Length") or 0)
            if length > MAX_UPLOAD_BYTES:
                return self._json({"error": "file too large"}, 413)
            body = self.rfile.read(length)

            if self.path == "/chat":
                payload = json.loads(body or b"{}")
                message = str(payload.get("message", "")).strip()
                if not message:
                    return self._json({"error": "empty message"}, 400)
                self._stream(lambda emit: self._chat(emit, message, bool(payload.get("plan"))))
            elif self.path == "/approve":
                self._stream(self._approve)
            elif self.path == "/discard":
                with session.lock:
                    session.pending = None
                self._json(session.state())
            elif self.path == "/track":
                # Like the switch in Word: applies to edits made from now on.
                with session.lock:
                    session.track_changes = bool(json.loads(body or b"{}").get("on"))
                    session.doc.track_changes = session.track_changes
                self._json(session.state())
            elif self.path == "/revisions":
                # Accept or reject every tracked change, like the buttons in Word.
                action = json.loads(body or b"{}").get("action")
                if action not in ("accept", "reject"):
                    return self._json({"error": "action must be 'accept' or 'reject'"}, 400)
                with session.lock:
                    if action == "accept":
                        session.doc.accept_revisions()
                    else:
                        session.doc.reject_revisions()
                self._json(session.state())
            elif self.path == "/reset":
                with session.lock:
                    session.reset()
                self._json(session.state())
            elif self.path == "/upload":
                name = unquote(self.headers.get("X-Filename", "document.docx"))
                try:
                    with session.lock:
                        session.reset(body, name)
                except (PackageError, ValueError) as exc:
                    return self._json({"error": str(exc)}, 400)
                self._json(session.state())
            else:
                self._json({"error": "not found"}, 404)

        # ------------------------------------------------------------- work

        async def _chat(self, emit, message: str, plan: bool) -> None:
            agent = self._agent(emit)
            if session.pending is not None:
                # A plan is waiting: treat the message as feedback on it.
                self._trace_start(agent, message, "revise_plan")
                session.pending = await agent.aplan(
                    session.pending.instruction, revise=session.pending, feedback=message
                )
                emit({"type": "plan", "text": session.pending.text})
                self._usage(emit, agent, session.pending.steps)
                self._trace_end(agent, status="planned", steps=session.pending.steps, plan=session.pending.text)
            elif plan:
                self._trace_start(agent, message, "plan")
                session.pending = await agent.aplan(message, history=session.history)
                emit({"type": "plan", "text": session.pending.text})
                self._usage(emit, agent, session.pending.steps)
                self._trace_end(agent, status="planned", steps=session.pending.steps, plan=session.pending.text)
            else:
                self._trace_start(agent, message, "run")
                result = await agent.arun(message, history=session.history)
                session.history = result.messages
                self._reply(emit, result)
                self._usage(emit, agent, result.steps)
                self._trace_end(agent, result)

        async def _approve(self, emit) -> None:
            if session.pending is None:
                emit({"type": "error", "text": "There is no plan to approve."})
                return
            agent = self._agent(emit)
            self._trace_start(agent, session.pending.instruction, "approve")
            result = await agent.aexecute(session.pending)
            session.history = result.messages
            session.pending = None
            self._reply(emit, result)
            # Only this request's model calls: the client is new for each request.
            self._usage(emit, agent, result.steps - result.plan.steps if result.plan else result.steps)
            self._trace_end(agent, result)

        @staticmethod
        def _reply(emit, result) -> None:
            emit({"type": "reply", "text": result.final_text or "(no reply)"})
            if result.status == "max_steps":
                emit({"type": "error", "text": f"Stopped after {result.steps} steps; the document may be incomplete."})

        @staticmethod
        def _usage(emit, agent, steps: int) -> None:
            text = format_usage(getattr(agent.llm, "usage", None), steps)
            if text:
                emit({"type": "usage", "text": text})

    return Handler


def make_server(
    port: int = 8765,
    model: str | None = None,
    llm_factory: Callable[[], LLMClient] | None = None,
    trace_dir: str | os.PathLike | None = None,
) -> ThreadingHTTPServer:
    """Build the server (not yet serving). ``llm_factory`` lets tests and
    custom setups supply their own client. With ``trace_dir`` each
    conversation is written to a JSONL file there (see ``docx_harness.trace``)."""
    if llm_factory is None:
        client = OpenRouterClient(model=model)  # fails early if the key is missing
        label = client.model
        llm_factory = lambda: OpenRouterClient(model=model)  # noqa: E731
    else:
        label = model or "custom client"
    session = Session(llm_factory, trace_dir)
    server = ThreadingHTTPServer(("127.0.0.1", port), make_handler(session, label))
    server.session = session  # type: ignore[attr-defined]
    return server


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="docx-harness-ui", description="Local chat UI for trying docx-harness.")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--model", help=f"OpenRouter model id (default: {DEFAULT_MODEL} or $DOCX_HARNESS_MODEL).")
    parser.add_argument("--no-browser", action="store_true", help="Do not open a browser tab.")
    parser.add_argument("--trace-dir", metavar="DIR", help="Write each conversation (model replies, tool calls and results, usage) to a JSONL file in DIR.")
    args = parser.parse_args(argv)

    try:
        from dotenv import load_dotenv

        load_dotenv()
    except ImportError:
        pass

    try:
        server = make_server(args.port, args.model, trace_dir=args.trace_dir)
    except (RuntimeError, OSError) as exc:
        print(f"error: {exc}")
        return 1
    url = f"http://127.0.0.1:{server.server_address[1]}"
    print(f"docx-harness UI running at {url}  (Ctrl+C to stop)")
    if args.trace_dir:
        print(f"Writing traces to {Path(args.trace_dir).resolve()}")
    if not args.no_browser:
        threading.Timer(0.3, webbrowser.open, [url]).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print()
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
