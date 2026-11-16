"""Run tasks against a model and collect what happened.

``run_suite`` is the live path. ``solve`` replays a task's reference
``solution`` with no model at all, which is how the task files themselves are
tested: the checks must fail on the untouched input and pass after the solution.
"""

from __future__ import annotations

import asyncio
import json
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Callable

from docx_harness import AgentEvent, LLMClient, ToolRegistry, WordDocument, create_agent, default_tools

from . import builder
from .checks import CheckResult, Context, run_checks
from .judge import judge as run_judge
from .loader import Task


@dataclass
class TaskResult:
    id: str
    category: str
    #: Every check passed and the run did not error.
    passed: bool = False
    checks: list[dict] = field(default_factory=list)
    #: "completed", "max_steps", or "error" when the run raised or timed out.
    status: str = "error"
    error: str = ""
    steps: int = 0
    tool_calls: int = 0
    #: Tool results that came back as "Error: ...".
    tool_errors: int = 0
    prompt_tokens: int = 0
    completion_tokens: int = 0
    #: In US dollars when the provider reports it, else None.
    cost: float | None = None
    seconds: float = 0.0
    judge_score: int | None = None
    judge_reason: str = ""
    final_text: str = ""
    #: The reviewer's feedback for each round (None: no problems found).
    reviews: list = field(default_factory=list)
    #: Review calls that failed (the work was kept unreviewed).
    review_errors: list = field(default_factory=list)

    @property
    def failed_checks(self) -> list[dict]:
        return [c for c in self.checks if not c["passed"]]


def prepare(task: Task, folder: Path) -> Path:
    """Write the task's starting document to ``folder`` and return its path."""
    folder.mkdir(parents=True, exist_ok=True)
    return builder.build(task.input).save(folder / "input.docx")


def _open(task: Task, path: Path) -> WordDocument:
    return WordDocument.open(path, track_changes=task.track_changes, author=task.author)


def grade(task: Task, out_path: Path, in_path: Path, status: str = "completed") -> list[CheckResult]:
    """The task's checks, after two that apply to every task: the agent
    finished, and the file opens in an independent reader."""
    results = [CheckResult("finished", None, status == "completed",
                           "" if status == "completed" else f"the run ended with status {status!r}")]
    try:
        import docx  # python-docx, a dev dependency
    except ImportError:
        docx = None
    if docx is not None:
        try:
            docx.Document(str(out_path))
            results.append(CheckResult("opens_with_python_docx", None, True))
        except Exception as exc:  # noqa: BLE001 - any failure to open is the finding
            results.append(CheckResult("opens_with_python_docx", None, False, f"python-docx could not open it: {exc}"))
    return results + run_checks(task.checks, Context(out_path, in_path))


async def solve(task: Task, folder: Path) -> list[CheckResult]:
    """Apply the task's reference solution through the tools, as a model
    would call them, and grade the result. Raises ValueError when a solution
    step is refused by its tool."""
    in_path = prepare(task, folder)
    doc = _open(task, in_path)
    registry = ToolRegistry(default_tools(doc))
    for name, arguments in task.solution:
        result = await registry.aexecute(name, arguments)
        if result.startswith("Error:"):
            raise ValueError(f"{task.id}: solution step {name} {arguments} failed: {result}")
    out_path = doc.save(folder / "output.docx")
    return grade(task, out_path, in_path)


async def verify(task: Task, folder: Path) -> list[str]:
    """Problems with the task file itself (an empty list when it is sound):
    its checks must not all pass before any work is done, and must all pass
    after its solution."""
    problems = []
    in_path = prepare(task, folder)
    untouched = _open(task, in_path).save(folder / "untouched.docx")
    if all(c.passed for c in grade(task, untouched, in_path)):
        problems.append("every check passes on the untouched input, so the task cannot fail")
    if not task.solution:
        return problems
    try:
        results = await solve(task, folder)
    except ValueError as exc:
        return problems + [str(exc)]
    problems += [f"after the solution, check {c.name} fails: {c.message}" for c in results if not c.passed]
    return problems


async def run_task(task: Task, llm: LLMClient, folder: Path, judge_llm: LLMClient | None = None,
                   timeout: float | None = 600, *, review: bool = True,
                   review_llm: LLMClient | None = None) -> TaskResult:
    """Run one task. Never raises: a failure is recorded on the result.
    With ``review`` the agent's work is checked by a reviewer (``review_llm``,
    else the agent's own client) and fixed; its tokens count toward the task."""
    result = TaskResult(task.id, task.category)
    started = time.monotonic()

    def on_event(event: AgentEvent) -> None:
        if event.type == "tool_call":
            result.tool_calls += 1
        elif event.type == "tool_result" and str(event.data.get("result", "")).startswith("Error:"):
            result.tool_errors += 1
        elif event.type == "review" and event.data.get("error"):
            result.review_errors.append(event.data["error"])

    try:
        in_path = prepare(task, folder)
        doc = _open(task, in_path)
        agent = create_agent(doc, llm=llm, max_steps=task.max_steps, on_event=on_event,
                             review=review, review_llm=review_llm)
        run = await asyncio.wait_for(agent.arun(task.instruction, plan_mode=task.plan), timeout)
        result.status, result.steps, result.final_text = run.status, run.steps, run.final_text
        result.reviews = run.reviews
        out_path = doc.save(folder / "output.docx")
        (folder / "trace.json").write_text(json.dumps(run.messages, indent=1, ensure_ascii=False), encoding="utf-8")
        result.checks = [asdict(c) for c in grade(task, out_path, in_path, run.status)]
        result.passed = all(c["passed"] for c in result.checks)
        if task.judge and judge_llm is not None:
            result.judge_score, result.judge_reason = await run_judge(
                judge_llm, task.instruction, task.judge, WordDocument.open(in_path), WordDocument.open(out_path))
    except Exception as exc:  # noqa: BLE001 - one broken task must not stop the suite
        result.status = "error"
        result.passed = False
        result.error = "timed out" if isinstance(exc, asyncio.TimeoutError) else f"{type(exc).__name__}: {exc}"
    clients = [llm] + ([review_llm] if review and review_llm is not None and review_llm is not llm else [])
    for client in clients:
        usage = getattr(client, "usage", None) or {}
        result.prompt_tokens += int(usage.get("prompt_tokens", 0))
        result.completion_tokens += int(usage.get("completion_tokens", 0))
        if usage.get("cost") is not None:
            result.cost = (result.cost or 0.0) + usage["cost"]
    result.seconds = round(time.monotonic() - started, 1)
    return result


async def run_suite(
    tasks: list[Task],
    make_llm: Callable[[], LLMClient],
    out_dir: Path,
    *,
    judge_llm: LLMClient | None = None,
    concurrency: int = 4,
    timeout: float | None = 600,
    on_result: Callable[[TaskResult], None] | None = None,
    review: bool = True,
    make_review_llm: Callable[[], LLMClient] | None = None,
) -> list[TaskResult]:
    """Run ``tasks`` concurrently, each with its own client from ``make_llm``
    (and ``make_review_llm``, when the reviewer uses another model) so token
    counts are per task. Results come back in task order; each task's
    documents and message trace are written under ``out_dir/<id>/``."""
    semaphore = asyncio.Semaphore(max(concurrency, 1))

    async def one(task: Task) -> TaskResult:
        async with semaphore:
            try:
                llm = make_llm()
                review_llm = make_review_llm() if review and make_review_llm is not None else None
            except Exception as exc:  # noqa: BLE001 - e.g. no API key
                result = TaskResult(task.id, task.category, error=f"{type(exc).__name__}: {exc}")
            else:
                result = await run_task(task, llm, Path(out_dir) / task.id, judge_llm, timeout,
                                        review=review, review_llm=review_llm)
            if on_result is not None:
                on_result(result)
            return result

    return list(await asyncio.gather(*(one(task) for task in tasks)))
