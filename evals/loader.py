"""Read task files into :class:`Task` objects, refusing anything misspelt.

A typo in a check name must be an error, not a check that silently never runs,
so every key, input step, check and solution tool is validated on load.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any

import yaml

from docx_harness import WordDocument, default_tools

from . import builder, checks

TASKS_DIR = Path(__file__).parent / "tasks"

_KEYS = {"id", "category", "instruction", "input", "checks", "judge", "solution",
         "track_changes", "author", "plan", "max_steps"}


class TaskError(ValueError):
    """A task file that cannot be used; the message names the file."""


@dataclass
class Task:
    id: str
    category: str
    instruction: str
    #: Steps that build the starting document, ``{"file": path}``, or None for a blank one.
    input: list[dict] | dict | None = None
    #: ``(check name, argument)`` pairs, in file order.
    checks: list[tuple[str, Any]] = field(default_factory=list)
    #: Rubric for the judge model, or None for no judging.
    judge: str | None = None
    #: ``(tool name, arguments)`` pairs: what a perfect model would call.
    solution: list[tuple[str, dict]] = field(default_factory=list)
    track_changes: bool = False
    author: str | None = None
    plan: bool = False
    max_steps: int = 40
    path: Path | None = None


@lru_cache(maxsize=1)
def tool_names() -> frozenset[str]:
    return frozenset(t.name for t in default_tools(WordDocument.open()))


def _pairs(items: Any, what: str, known: set[str] | frozenset[str]) -> list[tuple[str, Any]]:
    """``[{name: argument}, ...]`` as pairs, each name checked against ``known``."""
    if not isinstance(items, list):
        raise TaskError(f"{what} must be a list of single-key entries like '- name: value'")
    out = []
    for item in items:
        if not isinstance(item, dict) or len(item) != 1:
            raise TaskError(f"each entry of {what} must have exactly one key, got {item!r}")
        (name, argument), = item.items()
        if name not in known:
            raise TaskError(f"unknown name {name!r} in {what}. Known: {', '.join(sorted(known))}")
        out.append((name, argument))
    return out


def parse_task(data: Any, *, default_id: str = "task", default_category: str = "misc",
               path: Path | None = None) -> Task:
    """Build a Task from the parsed YAML of one file."""
    try:
        if not isinstance(data, dict):
            raise TaskError("the file must be a mapping with at least 'instruction' and 'checks'")
        unknown = set(data) - _KEYS
        if unknown:
            raise TaskError(f"unknown key(s) {', '.join(sorted(unknown))}. Known: {', '.join(sorted(_KEYS))}")
        instruction = data.get("instruction")
        if not isinstance(instruction, str) or not instruction.strip():
            raise TaskError("'instruction' is required")
        task_checks = _pairs(data.get("checks"), "checks", set(checks.CHECKS))
        if not task_checks:
            raise TaskError("'checks' needs at least one check")
        solution = _pairs(data.get("solution") or [], "solution", tool_names())
        for name, arguments in solution:
            if arguments is not None and not isinstance(arguments, dict):
                raise TaskError(f"solution step {name!r} needs a mapping of arguments, got {arguments!r}")
        try:
            builder.validate(data.get("input"))
        except ValueError as exc:
            raise TaskError(str(exc)) from None
        judge = data.get("judge")
        return Task(
            id=str(data.get("id") or default_id),
            category=str(data.get("category") or default_category),
            instruction=instruction.strip(),
            input=data.get("input"),
            checks=task_checks,
            judge=str(judge).strip() if judge else None,
            solution=[(name, arguments or {}) for name, arguments in solution],
            track_changes=bool(data.get("track_changes", False)),
            author=data.get("author"),
            plan=bool(data.get("plan", False)),
            max_steps=int(data.get("max_steps", 40)),
            path=path,
        )
    except TaskError as exc:
        raise TaskError(f"{path or default_id}: {exc}") from None


def load_task(path: str | Path) -> Task:
    path = Path(path)
    try:
        data = yaml.safe_load(path.read_text(encoding="utf-8"))
    except yaml.YAMLError as exc:
        raise TaskError(f"{path}: not valid YAML ({exc})") from None
    return parse_task(data, default_id=path.stem, default_category=path.parent.name, path=path)


def load_tasks(root: str | Path = TASKS_DIR) -> list[Task]:
    """Every task under ``root``, sorted by category then id. Ids must be unique."""
    tasks = [load_task(p) for p in sorted(Path(root).rglob("*.yaml"))]
    seen: dict[str, Path | None] = {}
    for task in tasks:
        if task.id in seen:
            raise TaskError(f"task id {task.id!r} is used by both {seen[task.id]} and {task.path}")
        seen[task.id] = task.path
    return sorted(tasks, key=lambda t: (t.category, t.id))


def select(tasks: list[Task], keyword: str | None = None, category: str | None = None) -> list[Task]:
    return [
        t for t in tasks
        if (not keyword or keyword.lower() in t.id.lower()) and (not category or t.category == category)
    ]
