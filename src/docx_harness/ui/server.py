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
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib import resources
from typing import Callable
from urllib.parse import quote, unquote

from ..agent import AgentEvent, LLMClient, OpenRouterClient
from ..agent.llm import DEFAULT_MODEL
from ..api import create_agent
from ..document import PackageError, WordDocument, WordPackage

MAX_UPLOAD_BYTES = 50 * 1024 * 1024
ALLOWED_HOSTS = {"127.0.0.1", "localhost"}


class Session:
    """The one document and conversation the UI is working on."""

    def __init__(self, llm_factory: Callable[[], LLMClient]):
        self.llm_factory = llm_factory
        self.lock = threading.Lock()
        self.reset()

    def reset(self, data: bytes | None = None, name: str | None = None) -> None:
        self.doc = WordDocument(WordPackage.from_bytes(data)) if data else WordDocument.open()
        self.name = name or "document.docx"
        self.history: list[dict] | None = None
        self.pending = None  # a Plan awaiting approval

    def state(self) -> dict:
        return {
            "name": self.name,
            "blocks": self.doc.read(count=1_000_000, include_formatting=True)["blocks"],
            "pending": self.pending.text if self.pending else None,
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
                emit({"type": "state", **session.state()})

        def _agent(self, emit):
            def on_event(event: AgentEvent) -> None:
                if event.type == "tool_call":
                    emit({"type": "tool", "name": event.data["name"], "arguments": event.data.get("arguments") or ""})
                elif event.type == "tool_result" and event.data["result"].startswith("Error"):
                    emit({"type": "tool_error", "text": event.data["result"][:300]})

            return create_agent(session.doc, llm=session.llm_factory(), on_event=on_event)

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
                session.pending = await agent.aplan(
                    session.pending.instruction, revise=session.pending, feedback=message
                )
                emit({"type": "plan", "text": session.pending.text})
            elif plan:
                session.pending = await agent.aplan(message, history=session.history)
                emit({"type": "plan", "text": session.pending.text})
            else:
                result = await agent.arun(message, history=session.history)
                session.history = result.messages
                self._reply(emit, result)

        async def _approve(self, emit) -> None:
            if session.pending is None:
                emit({"type": "error", "text": "There is no plan to approve."})
                return
            result = await self._agent(emit).aexecute(session.pending)
            session.history = result.messages
            session.pending = None
            self._reply(emit, result)

        @staticmethod
        def _reply(emit, result) -> None:
            emit({"type": "reply", "text": result.final_text or "(no reply)"})
            if result.status == "max_steps":
                emit({"type": "error", "text": f"Stopped after {result.steps} steps; the document may be incomplete."})

    return Handler


def make_server(
    port: int = 8765,
    model: str | None = None,
    llm_factory: Callable[[], LLMClient] | None = None,
) -> ThreadingHTTPServer:
    """Build the server (not yet serving). ``llm_factory`` lets tests and
    custom setups supply their own client."""
    if llm_factory is None:
        client = OpenRouterClient(model=model)  # fails early if the key is missing
        label = client.model
        llm_factory = lambda: OpenRouterClient(model=model)  # noqa: E731
    else:
        label = model or "custom client"
    session = Session(llm_factory)
    server = ThreadingHTTPServer(("127.0.0.1", port), make_handler(session, label))
    server.session = session  # type: ignore[attr-defined]
    return server


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="docx-harness-ui", description="Local chat UI for trying docx-harness.")
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--model", help=f"OpenRouter model id (default: {DEFAULT_MODEL} or $DOCX_HARNESS_MODEL).")
    parser.add_argument("--no-browser", action="store_true", help="Do not open a browser tab.")
    args = parser.parse_args(argv)

    try:
        from dotenv import load_dotenv

        load_dotenv()
    except ImportError:
        pass

    try:
        server = make_server(args.port, args.model)
    except (RuntimeError, OSError) as exc:
        print(f"error: {exc}")
        return 1
    url = f"http://127.0.0.1:{server.server_address[1]}"
    print(f"docx-harness UI running at {url}  (Ctrl+C to stop)")
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
