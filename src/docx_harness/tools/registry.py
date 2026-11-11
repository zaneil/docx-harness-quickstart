"""ToolRegistry: holds tools, exports their schemas and executes calls.

The registry never raises on a bad call. Failures come back as ``"Error: ..."``
strings so the model can read what went wrong and try again.
"""

from __future__ import annotations

import asyncio
import inspect
import json
from typing import Any, Callable, Iterable, Iterator

from .base import Tool, tool as make_tool


class ToolRegistry:
    def __init__(self, tools: Iterable[Tool | Callable] = ()):
        self._tools: dict[str, Tool] = {}
        for item in tools:
            self.register(item)

    def register(self, item: Tool | Callable) -> Tool:
        """Add a Tool, or a plain function (wrapped with ``@tool``). A tool with
        the same name is replaced, which is how you override a built-in."""
        registered = item if isinstance(item, Tool) else make_tool(item)
        self._tools[registered.name] = registered
        return registered

    def remove(self, name: str) -> None:
        self._tools.pop(name, None)

    def get(self, name: str) -> Tool | None:
        return self._tools.get(name)

    def __iter__(self) -> Iterator[Tool]:
        return iter(self._tools.values())

    def __len__(self) -> int:
        return len(self._tools)

    def __contains__(self, name: str) -> bool:
        return name in self._tools

    @property
    def names(self) -> list[str]:
        return list(self._tools)

    def to_openai_schema(self, read_only: bool = False) -> list[dict]:
        """Tool definitions in OpenAI ``tools`` format. With ``read_only=True``
        only tools that do not modify state are included (plan mode)."""
        return [t.to_openai() for t in self if t.read_only or not read_only]

    async def aexecute(self, name: str, arguments: str | dict | None = None, *, read_only: bool = False) -> str:
        """Run a tool call and return its result as text."""
        target = self._tools.get(name)
        if target is None:
            return f"Error: unknown tool {name!r}. Available tools: {', '.join(self._tools)}."
        if read_only and not target.read_only:
            return (
                f"Error: {name} modifies state and is not available in plan mode. "
                "Only inspect now; describe this change in your plan instead."
            )
        try:
            kwargs = _parse_arguments(arguments)
        except ValueError as exc:
            return f"Error: arguments for {name} are not a valid JSON object ({exc})."
        try:
            if target.blocking:
                result = await asyncio.to_thread(target.handler, **kwargs)
            else:
                result = target.handler(**kwargs)
                if inspect.isawaitable(result):
                    result = await result
        except Exception as exc:  # noqa: BLE001 - reported back to the model
            return f"Error: {exc}" if str(exc) else f"Error: {type(exc).__name__}"
        return _to_text(result)

    def execute(self, name: str, arguments: str | dict | None = None, *, read_only: bool = False) -> str:
        """Synchronous :meth:`aexecute` for use outside an event loop."""
        return asyncio.run(self.aexecute(name, arguments, read_only=read_only))


def _parse_arguments(arguments: str | dict | None) -> dict:
    if arguments is None or arguments == "":
        return {}
    if isinstance(arguments, dict):
        return arguments
    parsed = json.loads(arguments)
    if not isinstance(parsed, dict):
        raise ValueError("expected an object")
    return parsed


def _to_text(result: Any) -> str:
    if result is None:
        return "OK"
    if isinstance(result, str):
        return result
    return json.dumps(result, ensure_ascii=False, default=str)
