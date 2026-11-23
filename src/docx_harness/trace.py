"""Write what happens in a run to a JSONL file, one JSON object per line.

    trace = JsonlTrace("traces/run.jsonl")
    trace.start("Fix the typos", model="openai/gpt-6-luna")
    agent = create_agent(doc, llm=llm, on_event=trace.chain(my_callback))
    result = agent.run("Fix the typos")
    trace.end(result, llm.usage)

Records share a ``run`` id, so one file can hold many runs (the demo UI writes
one file per session). Each has ``ts`` (UTC, ISO 8601), ``type`` and
``elapsed_ms`` since the previous record of the same run: on an ``assistant``
record that is roughly the model call, on a ``tool_result`` the tool call
(tools that run concurrently overlap). Types are ``run_start``, the agent's
events (``assistant``, ``tool_call``, ``tool_result``, ``plan``) and
``run_end``. Every line is flushed when written, so a run that crashes still
leaves a readable file. Load a folder of them with pandas
(``pd.read_json(path, lines=True)``) or DuckDB (``read_json_auto('traces/*.jsonl')``).
"""

from __future__ import annotations

import json
import os
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from .agent import AgentEvent, AgentResult


class JsonlTrace:
    def __init__(self, path: str | os.PathLike):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.run = ""
        self._last = time.monotonic()

    def write(self, type_: str, **data: Any) -> None:
        now = time.monotonic()
        record = {
            "ts": datetime.now(timezone.utc).isoformat(timespec="milliseconds"),
            "run": self.run,
            "type": type_,
            "elapsed_ms": round((now - self._last) * 1000),
            **data,
        }
        self._last = now
        with self.path.open("a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False, default=str) + "\n")

    def start(self, instruction: str, **meta: Any) -> str:
        """Begin a new run; returns its id. ``run`` stays set until :meth:`end`."""
        self.run = uuid.uuid4().hex[:8]
        self._last = time.monotonic()
        self.write("run_start", instruction=instruction, **meta)
        return self.run

    def end(self, result: AgentResult | None = None, usage: dict | None = None, **extra: Any) -> None:
        data: dict[str, Any] = dict(extra)
        if result is not None:
            data.update(status=result.status, steps=result.steps, final_text=result.final_text)
        if usage:
            data["usage"] = dict(usage)
        self.write("run_end", **data)
        self.run = ""

    def __call__(self, event: AgentEvent) -> None:
        self.write(event.type, **event.data)

    def chain(self, on_event: Callable[[AgentEvent], Any] | None = None) -> Callable[[AgentEvent], Any]:
        """An ``on_event`` callback that records the event, then calls ``on_event``."""

        def callback(event: AgentEvent):
            self(event)
            return on_event(event) if on_event else None

        return callback
