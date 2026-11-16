"""Summaries of a run: a console table, ``results.json`` and a markdown table
for the README."""

from __future__ import annotations

import json
from dataclasses import asdict
from pathlib import Path

from .runner import TaskResult


def summarize(results: list[TaskResult | dict]) -> dict:
    """Totals for one model's run."""
    rows = [asdict(r) if isinstance(r, TaskResult) else r for r in results]
    count = len(rows)
    passed = sum(1 for r in rows if r["passed"])
    scores = [r["judge_score"] for r in rows if r.get("judge_score") is not None]
    costs = [r["cost"] for r in rows if r.get("cost") is not None]
    categories: dict[str, list[int]] = {}
    for r in rows:
        tally = categories.setdefault(r["category"], [0, 0])
        tally[0] += bool(r["passed"])
        tally[1] += 1
    return {
        "tasks": count,
        "passed": passed,
        "pass_rate": round(passed / count, 3) if count else 0.0,
        "categories": {name: {"passed": p, "tasks": n} for name, (p, n) in sorted(categories.items())},
        "mean_steps": round(sum(r["steps"] for r in rows) / count, 1) if count else 0.0,
        "tool_errors": sum(r["tool_errors"] for r in rows),
        "prompt_tokens": sum(r["prompt_tokens"] for r in rows),
        "completion_tokens": sum(r["completion_tokens"] for r in rows),
        "cost": round(sum(costs), 4) if costs else None,
        "mean_judge_score": round(sum(scores) / len(scores), 2) if scores else None,
        #: Tasks where the reviewer found something to fix.
        "reviewed_fixes": sum(1 for r in rows if any(r.get("reviews") or [])),
        "seconds": round(sum(r["seconds"] for r in rows), 1),
    }


def label(model: str, review: bool, review_model: str | None = None) -> str:
    """How a run is named in tables and folder names."""
    if not review:
        return f"{model} (no review)"
    return f"{model} (review by {review_model})" if review_model and review_model != model else model


def label_of(run: dict) -> str:
    """The label of a loaded ``results.json``. Runs from before the review
    step existed had none."""
    return label(run["model"], run.get("review", False), run.get("review_model"))


def write_results(path: Path, model: str, results: list[TaskResult], judge_model: str | None, started: str,
                  *, review: bool = True, review_model: str | None = None) -> dict:
    data = {
        "model": model,
        "review": review,
        "review_model": review_model,
        "judge_model": judge_model,
        "started": started,
        "summary": summarize(results),
        "tasks": [asdict(r) for r in results],
    }
    Path(path).write_text(json.dumps(data, indent=1, ensure_ascii=False), encoding="utf-8")
    return data


def _table(header: list[str], rows: list[list[str]]) -> str:
    widths = [max(len(str(row[i])) for row in [header, *rows]) for i in range(len(header))]
    line = lambda cells: "  ".join(str(c).ljust(w) for c, w in zip(cells, widths)).rstrip()  # noqa: E731
    return "\n".join([line(header), line(["-" * w for w in widths]), *(line(r) for r in rows)])


def _money(cost: float | None) -> str:
    return "n/a" if cost is None else f"${cost:.4f}" if cost < 1 else f"${cost:.2f}"


def _tokens(count: int) -> str:
    return f"{count / 1000:.1f}k" if count >= 1000 else str(count)


def task_line(result: TaskResult) -> str:
    """One line for a finished task, followed by what failed."""
    mark = "PASS" if result.passed else "FAIL"
    judge = f"  judge {result.judge_score}/5" if result.judge_score is not None else ""
    fixes = sum(1 for feedback in result.reviews if feedback)
    review = f"  {fixes} review fix round{'s' * (fixes != 1)}" if fixes else ""
    lines = [f"{mark}  {result.id}  ({result.steps} steps, {result.seconds}s){judge}{review}"]
    if result.error:
        lines.append(f"      error: {result.error}")
    lines += [f"      review failed: {error[:200]}" for error in result.review_errors]
    lines += [f"      {c['name']}: {c['message']}" for c in result.failed_checks]
    return "\n".join(lines)


def console(model: str, results: list[TaskResult]) -> str:
    """A per-task table and the totals."""
    rows = [
        [r.category, r.id, "pass" if r.passed else "FAIL",
         f"{sum(c['passed'] for c in r.checks)}/{len(r.checks)}", r.steps,
         sum(1 for feedback in r.reviews if feedback), r.tool_errors,
         _tokens(r.prompt_tokens + r.completion_tokens), _money(r.cost),
         "" if r.judge_score is None else f"{r.judge_score}/5"]
        for r in results
    ]
    s = summarize(results)
    judge = "" if s["mean_judge_score"] is None else f", judge {s['mean_judge_score']}/5"
    total = (
        f"{model}: {s['passed']}/{s['tasks']} passed ({s['pass_rate']:.0%}), "
        f"{s['mean_steps']} steps per task, {s['reviewed_fixes']} task(s) fixed after review, {_tokens(s['prompt_tokens'] + s['completion_tokens'])} tokens, "
        f"{_money(s['cost'])}{judge}"
    )
    header = ["category", "task", "result", "checks", "steps", "fixes", "tool errors", "tokens", "cost", "judge"]
    return _table(header, rows) + "\n\n" + total


def markdown(runs: list[dict]) -> str:
    """One row per model, from the data of one or more ``results.json`` files."""
    categories = sorted({name for run in runs for name in run["summary"]["categories"]})
    header = ["Model", "Passed", *[c.capitalize() for c in categories], "Steps per task", "Tokens per task",
              "Cost per task", "Judge (1-5)"]
    lines = ["| " + " | ".join(header) + " |", "| " + " | ".join("---" for _ in header) + " |"]
    for run in runs:
        s = run["summary"]
        count = max(s["tasks"], 1)
        per_category = [
            f"{s['categories'][c]['passed']}/{s['categories'][c]['tasks']}" if c in s["categories"] else ""
            for c in categories
        ]
        cells = [
            f"`{run['model']}`" + label_of(run)[len(run["model"]):],
            f"{s['passed']}/{s['tasks']} ({s['pass_rate']:.0%})",
            *per_category,
            str(s["mean_steps"]),
            _tokens(round((s["prompt_tokens"] + s["completion_tokens"]) / count)),
            _money(None if s["cost"] is None else s["cost"] / count),
            "" if s["mean_judge_score"] is None else str(s["mean_judge_score"]),
        ]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def load_results(path: str | Path) -> dict:
    path = Path(path)
    if path.is_dir():
        path = path / "results.json"
    return json.loads(path.read_text(encoding="utf-8"))
