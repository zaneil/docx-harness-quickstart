"""Tool: a function an LLM can call, plus the JSON schema that describes it.

The ``@tool`` decorator builds the schema from the function's type hints and
Google-style docstring, so adding a tool is one decorated function:

    @tool(read_only=True)
    def word_count(block_id: str) -> int:
        '''Count the words in a block.

        Args:
            block_id: Id of the block to count.
        '''
"""

from __future__ import annotations

import inspect
import re
import types
import typing
from dataclasses import dataclass
from typing import Any, Callable


@dataclass
class Tool:
    name: str
    description: str
    parameters: dict
    handler: Callable[..., Any]
    #: True when the tool only inspects state. Read-only tools stay available
    #: in plan mode and may run concurrently.
    read_only: bool = False
    #: True for sync handlers that block on I/O; they run in a worker thread.
    blocking: bool = False

    def to_openai(self) -> dict:
        """The tool in OpenAI chat-completions ``tools`` format."""
        return {
            "type": "function",
            "function": {
                "name": self.name,
                "description": self.description,
                "parameters": self.parameters,
            },
        }


_SIMPLE = {str: "string", int: "integer", float: "number", bool: "boolean"}


def _schema_for(annotation: Any) -> dict:
    if annotation is inspect.Parameter.empty or annotation is Any:
        return {}
    if annotation in _SIMPLE:
        return {"type": _SIMPLE[annotation]}
    origin = typing.get_origin(annotation)
    args = typing.get_args(annotation)
    if origin is typing.Union or origin is types.UnionType:
        options = [a for a in args if a is not type(None)]
        if len(options) == 1:
            return _schema_for(options[0])
        return {"anyOf": [_schema_for(a) for a in options]}
    if origin is typing.Literal:
        kinds = {_SIMPLE.get(type(a), "string") for a in args}
        return {"type": kinds.pop() if len(kinds) == 1 else "string", "enum": list(args)}
    if annotation is list or origin in (list, tuple, typing.Sequence):
        schema: dict = {"type": "array"}
        if args:
            schema["items"] = _schema_for(args[0])
        return schema
    if annotation is dict or origin is dict:
        return {"type": "object"}
    return {}


_SECTION = re.compile(r"^(Args|Arguments|Parameters|Returns|Raises|Example|Examples|Note|Notes):\s*$")
_PARAM = re.compile(r"^(\*{0,2}\w+)\s*(?:\([^)]*\))?:\s*(.*)$")


def _parse_docstring(doc: str | None) -> tuple[str, dict[str, str]]:
    """Split a Google-style docstring into (description, {param: description})."""
    if not doc:
        return "", {}
    description: list[str] = []
    params: dict[str, str] = {}
    section = None
    current = None
    param_indent = None
    for line in inspect.cleandoc(doc).splitlines():
        stripped = line.strip()
        header = _SECTION.match(stripped)
        if header and not line.startswith(" "):
            section = header.group(1)
            current = None
            continue
        if section is None:
            description.append(stripped)
        elif section in ("Args", "Arguments", "Parameters") and stripped:
            indent = len(line) - len(line.lstrip())
            match = _PARAM.match(stripped)
            if match and (param_indent is None or indent <= param_indent):
                param_indent = indent
                current = match.group(1).lstrip("*")
                params[current] = match.group(2).strip()
            elif current:
                params[current] += " " + stripped
    return " ".join(" ".join(description).split()), params


def build_schema(fn: Callable) -> dict:
    """JSON schema for a function's parameters."""
    try:
        hints = typing.get_type_hints(fn)
    except Exception:
        hints = {}
    _, docs = _parse_docstring(fn.__doc__)
    properties: dict[str, dict] = {}
    required: list[str] = []
    for name, param in inspect.signature(fn).parameters.items():
        if name in ("self", "cls") or param.kind in (param.VAR_POSITIONAL, param.VAR_KEYWORD):
            continue
        schema = _schema_for(hints.get(name, param.annotation))
        if name in docs:
            schema["description"] = docs[name]
        properties[name] = schema
        if param.default is inspect.Parameter.empty:
            required.append(name)
    out: dict = {"type": "object", "properties": properties}
    if required:
        out["required"] = required
    return out


def tool(
    fn: Callable | None = None,
    *,
    name: str | None = None,
    description: str | None = None,
    parameters: dict | None = None,
    read_only: bool = False,
    blocking: bool = False,
):
    """Turn a function (sync or async) into a :class:`Tool`.

    Usable bare (``@tool``) or with options (``@tool(read_only=True)``). Pass
    ``parameters`` to supply a hand-written JSON schema instead of the derived one.
    """

    def wrap(func: Callable) -> Tool:
        doc_description, _ = _parse_docstring(func.__doc__)
        return Tool(
            name=name or func.__name__,
            description=description or doc_description or func.__name__,
            parameters=parameters or build_schema(func),
            handler=func,
            read_only=read_only,
            blocking=blocking,
        )

    return wrap(fn) if fn is not None else wrap
