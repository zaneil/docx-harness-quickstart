"""Command line for the evals.

    uv run python -m evals list
    uv run python -m evals show fix-typos
    uv run python -m evals check
    uv run python -m evals run --model openai/gpt-4.1-mini
    uv run python -m evals report evals/results/<run> --markdown
"""

from __future__ import annotations

import argparse
import asyncio
import os
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

from docx_harness import OpenRouterClient
from docx_harness.agent.llm import DEFAULT_MODEL

from . import report
from .judge import DEFAULT_JUDGE_MODEL
from .loader import TaskError, load_tasks, select
from .runner import prepare, run_suite, verify

RESULTS_DIR = Path(__file__).parent / "results"


def _tasks(args) -> list:
    tasks = select(load_tasks(), getattr(args, "keyword", None), getattr(args, "category", None))
    if not tasks:
        raise TaskError("no task matches; see 'python -m evals list'")
    return tasks


def cmd_list(args) -> int:
    tasks = _tasks(args)
    print(report._table(
        ["category", "task", "checks", "judge", "instruction"],
        [[t.category, t.id, len(t.checks), "yes" if t.judge else "", t.instruction.split("\n")[0][:70]] for t in tasks],
    ))
    print(f"\n{len(tasks)} task(s)")
    return 0


def cmd_show(args) -> int:
    """Print a task's starting document with block ids, for writing a solution."""
    from docx_harness import WordDocument

    for task in _tasks(args):
        with tempfile.TemporaryDirectory() as folder:
            doc = WordDocument.open(prepare(task, Path(folder)))
        print(f"{task.id} ({task.category}): {task.instruction}\n")
        for entry in doc.outline():
            print(f"  {entry['id']} [{entry.get('style') or entry['type']}] {entry['preview']}")
        if not len(doc):
            print("  (blank document)")
        for story in doc.stories():
            print(f"  {story['part']} ({story['kind']}): {' / '.join(story['paragraphs'])}")
        for r in doc.revisions():
            print(f"  revision [{r['id']}] {r['type']} by {r['author']}: {r['text']!r}")
        print()
    return 0


def cmd_check(args) -> int:
    """Offline: every task's checks fail before its solution and pass after it."""
    failed = 0
    for task in _tasks(args):
        with tempfile.TemporaryDirectory() as folder:
            problems = asyncio.run(verify(task, Path(folder)))
        if not task.solution:
            problems.append("no solution, so the checks are not proven passable")
        failed += bool(problems)
        print(f"{'ok  ' if not problems else 'FAIL'}  {task.id}")
        for problem in problems:
            print(f"      {problem}")
    print(f"\n{failed} task(s) with problems" if failed else "\nAll tasks are sound.")
    return 1 if failed else 0


def cmd_run(args) -> int:
    tasks = _tasks(args)
    judge_model = None if args.no_judge else (
        args.judge_model or os.environ.get("DOCX_HARNESS_JUDGE_MODEL") or DEFAULT_JUDGE_MODEL)
    judge_llm = OpenRouterClient(model=judge_model) if judge_model and any(t.judge for t in tasks) else None
    models = args.model or [os.environ.get("DOCX_HARNESS_MODEL") or DEFAULT_MODEL]
    worst = 0
    for model in models:
        started = datetime.now(timezone.utc)
        out_dir = Path(args.out or RESULTS_DIR) / f"{started:%Y%m%d-%H%M%S}_{re.sub(r'[^A-Za-z0-9.]+', '-', model)}"
        print(f"Running {len(tasks)} task(s) on {model}; output in {out_dir}\n", file=sys.stderr)
        results = asyncio.run(run_suite(
            tasks, lambda: OpenRouterClient(model=model), out_dir, judge_llm=judge_llm,
            concurrency=args.jobs, timeout=args.timeout,
            on_result=lambda r: print(report.task_line(r), file=sys.stderr),
        ))
        report.write_results(out_dir / "results.json", model, results, judge_model if judge_llm else None,
                             started.isoformat(timespec="seconds"))
        print("\n" + report.console(model, results) + "\n")
        worst = max(worst, 0 if all(r.passed for r in results) else 1)
    return worst


def cmd_report(args) -> int:
    runs = [report.load_results(path) for path in args.results]
    if args.markdown:
        print(report.markdown(runs))
        return 0
    for run in runs:
        s = run["summary"]
        print(f"{run['model']}: {s['passed']}/{s['tasks']} passed, {s['mean_steps']} steps per task, "
              f"{s['prompt_tokens'] + s['completion_tokens']} tokens, cost {report._money(s['cost'])}, "
              f"judge {s['mean_judge_score']}")
        for task in run["tasks"]:
            if not task["passed"]:
                reasons = "; ".join(f"{c['name']}: {c['message']}" for c in task["checks"] if not c["passed"])
                print(f"  FAIL {task['id']}: {task['error'] or reasons}")
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m evals", description="Evals for docx-harness.")
    commands = parser.add_subparsers(dest="command", required=True)

    def add(name: str, handler, help_: str, filters: bool = True) -> argparse.ArgumentParser:
        sub = commands.add_parser(name, help=help_)
        sub.set_defaults(handler=handler)
        if filters:
            sub.add_argument("-k", "--keyword", help="Only tasks whose id contains this.")
            sub.add_argument("--category", help="Only tasks in this category (the folder name).")
        return sub

    add("list", cmd_list, "List the tasks.")
    add("show", cmd_show, "Print a task's starting document with block ids.").add_argument(
        "keyword_positional", nargs="?", metavar="task", help="Task id (or part of one).")
    add("check", cmd_check, "Offline: prove each task's checks with its reference solution.").add_argument(
        "keyword_positional", nargs="?", metavar="task", help="Task id (or part of one).")
    run = add("run", cmd_run, "Run the tasks against a model (needs OPENROUTER_API_KEY).")
    run.add_argument("--model", action="append", help="OpenRouter model id. Repeat to run several models.")
    run.add_argument("-j", "--jobs", type=int, default=4, help="Tasks to run at once (default: 4).")
    run.add_argument("--no-judge", action="store_true", help="Skip the judge model.")
    run.add_argument("--judge-model", help=f"Judge model (default: $DOCX_HARNESS_JUDGE_MODEL or {DEFAULT_JUDGE_MODEL}).")
    run.add_argument("--timeout", type=float, default=600, help="Seconds allowed per task (default: 600).")
    run.add_argument("--out", help="Folder for results (default: evals/results).")
    rep = add("report", cmd_report, "Summarise results.json files.", filters=False)
    rep.add_argument("results", nargs="+", help="results.json files or run folders.")
    rep.add_argument("--markdown", action="store_true", help="Print a markdown table, one row per model.")

    args = parser.parse_args(argv)
    if getattr(args, "keyword_positional", None):
        args.keyword = args.keyword_positional

    try:
        from dotenv import load_dotenv

        load_dotenv()
    except ImportError:
        pass

    try:
        return args.handler(args)
    except (TaskError, RuntimeError, FileNotFoundError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
