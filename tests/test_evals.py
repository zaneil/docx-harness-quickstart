"""The eval suite itself, tested offline: no model and no network.

Every task file is proven with its reference solution, every check has a
passing and a failing case, and the runner is driven by a scripted model.
"""

from __future__ import annotations

import json
from types import SimpleNamespace

import pytest

from docx_harness import OpenAICompatibleClient

from conftest import ScriptedLLM, assistant, call
from evals import builder, checks, judge, report
from evals.__main__ import main
from evals.loader import TaskError, load_tasks, parse_task, select
from evals.runner import run_suite, verify

TASKS = load_tasks()


# ------------------------------------------------------------- task files

def test_the_suite_has_tasks_in_every_category():
    assert len(TASKS) >= 20
    assert {t.category for t in TASKS} == {"draft", "edit", "restructure", "tables", "redline"}
    assert len({t.id for t in TASKS}) == len(TASKS)


@pytest.mark.parametrize("task", TASKS, ids=lambda t: t.id)
async def test_task_fails_untouched_and_passes_with_its_solution(task, tmp_path):
    assert task.solution, "every bundled task needs a reference solution"
    assert await verify(task, tmp_path) == []


def test_select_filters_by_keyword_and_category():
    assert [t.id for t in select(TASKS, keyword="fix-typos")] == ["fix-typos"]
    assert {t.category for t in select(TASKS, category="tables")} == {"tables"}


# ----------------------------------------------------------------- loader

GOOD = {"instruction": "Do it.", "checks": [{"text_contains": "x"}]}


def test_parse_task_defaults():
    task = parse_task(GOOD, default_id="mine", default_category="edit")
    assert (task.id, task.category, task.max_steps, task.track_changes, task.judge) == ("mine", "edit", 40, False, None)
    assert task.checks == [("text_contains", "x")]


@pytest.mark.parametrize("change, expected", [
    ({"cheks": []}, "unknown key(s) cheks"),
    ({"instruction": " "}, "'instruction' is required"),
    ({"checks": []}, "at least one check"),
    ({"checks": [{"text_contians": "x"}]}, "unknown name 'text_contians' in checks. Known:"),
    ({"checks": [{"text_contains": "x", "regex": "y"}]}, "exactly one key"),
    ({"solution": [{"replace_txt": {}}]}, "unknown name 'replace_txt' in solution"),
    ({"solution": [{"replace_text": "x"}]}, "needs a mapping of arguments"),
    ({"input": [{"paragraf": "x"}]}, "must have exactly one of"),
    ({"input": [{"heading": "x", "style": "Quote"}]}, "does not take style"),
    ({"input": {"file": "fixtures/missing.docx"}}, "was not found"),
    ({"input": [{"tracked": {"author": "A", "steps": [{"swap": 1}]}}]}, "tracked step"),
])
def test_parse_task_rejects_mistakes(change, expected):
    with pytest.raises(TaskError) as error:
        parse_task({**GOOD, **change}, default_id="broken")
    assert expected in str(error.value)
    assert "broken" in str(error.value)  # says which task


def test_duplicate_ids_are_refused(tmp_path):
    for folder in ("a", "b"):
        (tmp_path / folder).mkdir()
        (tmp_path / folder / "same.yaml").write_text("instruction: x\nchecks:\n  - text_contains: x\n")
    with pytest.raises(TaskError, match="used by both"):
        load_tasks(tmp_path)


# ----------------------------------------------------------------- checks

DOC = [
    {"header": "Acme internal"},
    {"footer": "Page {PAGE} of {NUMPAGES}"},
    {"title": "Launch plan"},
    {"heading": "Goals"},
    {"paragraph": "Ship the “beta” by March 3 and review the handbook."},
    {"format": {"text": "beta", "bold": True, "color": "FF0000"}},
    {"link": {"text": "the handbook", "url": "https://example.com/handbook"}},
    {"heading": "Details", "level": 2},
    {"bullets": ["Alpha", "Bravo"]},
    {"numbered": ["First"]},
    {"table": [["Name", "Qty"], ["Bolts", "4"]]},
    {"paragraph": "Old clause."},
    {"tracked": {"author": "Alice", "steps": [{"replace": ["March 3", "March 9"]}, {"delete": "Old clause."}]}},
]

# check -> (argument that passes on DOC, argument that fails on DOC)
CASES = {
    "text_contains": (["Ship the \"beta\"", "March 9"], "March 3"),
    "text_contains_any": (["nope", "bolts"], ["nope", "nada"]),
    "text_absent": (["Old clause", "ship the"], {"text": "ship the", "case": False}),
    "regex": ({"pattern": r"^Bolts$", "min": 1, "max": 1}, r"^\d+\.\s"),
    "word_count": ({"min": 10, "max": 40}, {"max": 5}),
    "section_words": ({"heading": "Goals", "min": 10, "max": 10}, {"heading": "Details", "max": 3}),
    "unchanged": (["Goals", "Bolts"], ["Ship the"]),
    "unchanged_except": (["Old clause", "March"], ["Old clause"]),
    "order": (["Launch plan", "Goals", "Alpha", "Bolts"], ["Bolts", "Goals"]),
    "heading": ({"text": "Details", "level": 2}, {"text": "Details", "level": 1}),
    "heading_count": (2, {"level": 1, "min": 2}),
    "style_of": ({"text": "Launch plan", "style": ["Subtitle", "Title"]}, {"text": "Goals", "style": "Title"}),
    "block_count": (8, {"max": 3}),
    "list_items": ({"kind": "bullet", "items": ["Alpha", "Bravo"], "min": 2, "max": 2},
                   {"kind": "number", "items": ["Alpha"]}),
    "paragraph_format": ({"text": "Goals"}, {"text": "Goals", "alignment": "center"}),
    "table_count": (1, {"min": 2}),
    "table": ({"with_text": "Bolts", "rows": 2, "cols": 2, "cells": [{"at": [1, 1], "text": "4"}],
               "has_row": [["Name", "Qty"]], "contains": ["bolts"], "absent": ["nuts"]},
              {"has_row": [["Bolts", "5"]]}),
    "formatting": ({"text": "beta", "bold": True, "color": "ff0000", "only": True}, {"text": "Ship", "bold": True}),
    "link": ({"text": "the handbook", "url": "https://example.com/handbook/"},
             {"text": "handbook", "url": "https://example.com/handbook"}),
    "header_contains": ("Acme internal", "Zenith"),
    "footer_contains": (["{PAGE}", "{NUMPAGES}"], "Confidential"),
    "revisions": ({"author": "Alice", "type": "delete", "min": 2}, 0),
    "reject_all_restores_input": ({"author": "Alice"}, None),
}


@pytest.fixture(scope="module")
def ctx(tmp_path_factory):
    folder = tmp_path_factory.mktemp("checks")
    path = builder.build(DOC).save(folder / "doc.docx")
    # the "input" is the same document before Alice's revisions
    before = builder.build(DOC[:-1]).save(folder / "before.docx")
    return checks.Context(path, before)


def test_every_check_has_a_case():
    assert set(CASES) == set(checks.CHECKS)


@pytest.mark.parametrize("name", sorted(CASES))
def test_check_passes_and_fails(name, ctx):
    passing, failing = CASES[name]
    assert checks.CHECKS[name](ctx, passing) is None
    if failing is not None:
        message = checks.CHECKS[name](ctx, failing)
        assert message and isinstance(message, str)


def test_more_table_and_format_failures(ctx):
    table = checks.CHECKS["table"]
    assert "no table containing" in table(ctx, {"with_text": "Washers"})
    assert "does not exist" in table(ctx, {"index": 3})
    assert "table rows" in table(ctx, {"rows": {"min": 5}})
    assert "expected '7'" in table(ctx, {"cells": [{"at": [1, 1], "text": "7"}]})
    assert "cells without shading" in table(ctx, {"shaded_row": 0})
    assert "still in the table" in table(ctx, {"absent": ["Bolts"]})
    assert "not in a body paragraph" in checks.CHECKS["formatting"](ctx, {"text": "zebra", "bold": True})
    assert "no heading 'Budget'" in checks.CHECKS["section_words"](ctx, {"heading": "Budget", "max": 5})
    # 'only' catches formatting that spilled onto the rest of the paragraph
    assert "other than" in checks.CHECKS["formatting"](ctx, {"text": "Ship", "bold": False, "only": True})


def test_shading_and_untracked_edits_are_seen(tmp_path):
    doc = builder.build([{"paragraph": "One"}, {"table": [["A", "B"], ["1", "2"]]}])
    before = doc.save(tmp_path / "before.docx")
    doc.format_cells("b2[0,*]", shading="D9D9D9")
    doc.track_changes = True
    doc.replace_text("One", "Two")
    doc.track_changes = False
    doc.set_cell_text("b2", 1, 0, "9")  # not recorded as a revision
    context = checks.Context(doc.save(tmp_path / "after.docx"), before)
    assert checks.CHECKS["table"](context, {"shaded_row": 0, "shaded": [[0, 1]]}) is None
    assert "differs from the input" in checks.CHECKS["reject_all_restores_input"](context, True)


def test_a_crashing_check_is_a_failed_check(ctx):
    results = checks.run_checks([("style_of", {"text": "Goals"}), ("text_contains", "Goals")], ctx)
    assert [r.passed for r in results] == [False, True]
    assert "check could not run" in results[0].message


def test_norm_forgives_typography():
    assert checks.norm("It’s  “done” — now go") == "It's \"done\" - now go"


# ------------------------------------------------------------------ judge

def test_parse_reply():
    assert judge.parse_reply('Sure: {"score": 4, "reason": "Good."}') == (4, "Good.")
    assert judge.parse_reply("no json here")[0] is None
    assert judge.parse_reply('{"score": 9}')[0] is None
    assert judge.parse_reply(None)[0] is None


def test_describe_shows_structure_and_revisions(ctx):
    text = judge.describe(ctx.out)
    assert "[Title] Launch plan" in text
    assert "Bolts | 4" in text
    assert "[header] Acme internal" in text
    assert "link: the handbook -> https://example.com/handbook" in text
    assert "Tracked changes pending" in text and "by Alice" in text


# ----------------------------------------------------------------- runner

TYPOS = next(t for t in TASKS if t.id == "fix-typos")
MEMO = next(t for t in TASKS if t.id == "memo")


def solved(task, final="Done.", review='{"ok": true}'):
    """A scripted model that makes the task's reference solution calls, then
    answers the reviewer with ``review`` (None: no reviewer turn)."""
    turns = [assistant(*(call(name, **arguments) for name, arguments in task.solution)), assistant(content=final)]
    return ScriptedLLM(turns + ([assistant(content=review)] if review is not None else []))


async def test_run_suite_grades_judges_and_survives_errors(tmp_path):
    lazy = parse_task({**GOOD, "instruction": "Fix it.", "input": [{"paragraph": "Teh end."}],
                       "checks": [{"text_contains": "The end."}]}, default_id="lazy")
    broken = parse_task(GOOD, default_id="broken")
    models = {
        "fix-typos": solved(TYPOS),
        "memo": solved(MEMO),
        "lazy": ScriptedLLM([assistant(call("update_paragraph", block_id="b9", text="x")),
                             assistant(content="All fixed!")]),
        "broken": ScriptedLLM([]),  # raises on the first call
    }
    order = iter(["fix-typos", "memo", "lazy", "broken"])
    judge_llm = ScriptedLLM([assistant(content='{"score": 5, "reason": "Clear memo."}')])
    seen = []
    results = await run_suite([TYPOS, MEMO, lazy, broken], lambda: models[next(order)], tmp_path,
                              judge_llm=judge_llm, concurrency=1, on_result=seen.append)

    typos, memo, lazy_result, broken_result = results
    assert seen == results
    assert typos.passed and typos.status == "completed" and typos.steps == 2 and typos.tool_calls == 5
    assert typos.judge_score is None  # no rubric on that task
    assert typos.reviews == [None] and memo.reviews == [None]
    assert lazy_result.reviews == []  # its only edit failed, so there was nothing to review
    assert broken_result.review_errors == []
    assert {c["name"] for c in typos.checks} >= {"finished", "opens_with_python_docx", "text_contains"}
    assert memo.passed and (memo.judge_score, memo.judge_reason) == (5, "Clear memo.")
    # the judge saw the request, the rubric and the document, with no tools
    request = judge_llm.requests[0]
    assert request["tools"] is None
    assert "RUBRIC:" in request["messages"][1]["content"] and "200 Harbor Street" in request["messages"][1]["content"]

    assert not lazy_result.passed and lazy_result.tool_errors == 1
    assert [c["name"] for c in lazy_result.failed_checks] == ["text_contains"]
    assert not broken_result.passed and broken_result.status == "error"
    assert "ran out of turns" in broken_result.error

    for name in ("input.docx", "output.docx", "trace.json"):
        assert (tmp_path / "fix-typos" / name).exists()
    assert json.loads((tmp_path / "fix-typos" / "trace.json").read_text())[0]["role"] == "system"

    data = report.write_results(tmp_path / "results.json", "fake/model", results, "fake/judge", "2026-01-01T00:00:00")
    summary = data["summary"]
    assert (summary["tasks"], summary["passed"], summary["pass_rate"]) == (4, 2, 0.5)
    assert summary["categories"]["edit"] == {"passed": 1, "tasks": 1}
    assert summary["mean_judge_score"] == 5 and summary["cost"] is None and summary["tool_errors"] == 1
    assert "fake/model: 2/4 passed (50%)" in report.console("fake/model", results)
    assert "FAIL  lazy" in report.task_line(lazy_result) and "text_contains:" in report.task_line(lazy_result)
    table = report.markdown([report.load_results(tmp_path)])
    assert "| `fake/model` | 2/4 (50%) |" in table and "Draft" in table


async def test_run_suite_records_max_steps_tokens_and_a_bad_client(tmp_path):
    class Counting(ScriptedLLM):
        usage = {"prompt_tokens": 120, "completion_tokens": 30, "cost": 0.002}

    looping = parse_task({**GOOD, "max_steps": 1}, default_id="looping")
    results = await run_suite([looping], lambda: Counting([assistant(call("get_outline"))]), tmp_path, review=False)
    assert results[0].status == "max_steps" and not results[0].passed
    assert results[0].failed_checks[0]["name"] == "finished"
    assert (results[0].prompt_tokens, results[0].completion_tokens, results[0].cost) == (120, 30, 0.002)
    assert report.summarize(results)["cost"] == 0.002

    def no_key():
        raise RuntimeError("No OpenRouter API key.")

    results = await run_suite([looping], no_key, tmp_path)
    assert results[0].status == "error" and "No OpenRouter API key" in results[0].error


async def test_tracked_task_runs_with_tracking_on(tmp_path):
    tracked = next(t for t in TASKS if t.id == "tracked-typos")
    results = await run_suite([tracked], lambda: solved(tracked), tmp_path)
    assert results[0].passed, results[0].failed_checks


# -------------------------------------------------------------------- CLI

def test_cli_list_show_check_and_report(tmp_path, capsys):
    assert main(["list", "--category", "redline"]) == 0
    out = capsys.readouterr().out
    assert "tracked-typos" in out and "fix-typos" not in out

    assert main(["show", "accept-reject"]) == 0
    out = capsys.readouterr().out
    assert "b2 [Normal] Invoices are payable within 45 days." in out and "by Bob" in out

    assert main(["check", "fix-typos"]) == 0
    assert "ok    fix-typos" in capsys.readouterr().out

    assert main(["list", "-k", "no-such-task"]) == 2
    assert "no task matches" in capsys.readouterr().err

    (tmp_path / "results.json").write_text(json.dumps({
        "model": "m", "summary": report.summarize([]), "tasks": []}))
    assert main(["report", str(tmp_path), "--markdown"]) == 0
    assert "| `m` (no review) | 0/0 (0%) |" in capsys.readouterr().out  # written before the review step existed
    (tmp_path / "results.json").write_text(json.dumps({
        "model": "m", "review": True, "summary": report.summarize([]), "tasks": []}))
    assert main(["report", str(tmp_path), "--markdown"]) == 0
    assert "| `m` | 0/0 (0%) |" in capsys.readouterr().out


# ------------------------------------------------------------------- cost

async def test_client_sums_cost_when_the_provider_reports_it():
    def response(cost):
        usage = SimpleNamespace(prompt_tokens=10, completion_tokens=5)
        if cost is not None:
            usage.cost = cost
        message = SimpleNamespace(content="hi", tool_calls=None)
        return SimpleNamespace(usage=usage, choices=[SimpleNamespace(message=message)])

    replies = [response(0.001), response(0.002), response(None)]

    async def create(**request):
        return replies.pop(0)

    client = OpenAICompatibleClient("some/model", api_key="k")
    client._get_client = lambda: SimpleNamespace(chat=SimpleNamespace(completions=SimpleNamespace(create=create)))
    await client.acomplete([{"role": "user", "content": "x"}])
    assert client.usage["cost"] == pytest.approx(0.001)
    await client.acomplete([{"role": "user", "content": "x"}])
    await client.acomplete([{"role": "user", "content": "x"}])
    assert client.usage == {"prompt_tokens": 30, "completion_tokens": 15, "requests": 3, "cost": pytest.approx(0.003)}


async def test_review_feedback_is_fixed_and_counted(tmp_path):
    lazy = parse_task({**GOOD, "instruction": "Fix it.", "input": [{"paragraph": "Teh end."}],
                       "checks": [{"text_contains": "The end."}]}, default_id="lazy")

    class Counting(ScriptedLLM):
        def __init__(self, turns, tokens):
            super().__init__(turns)
            self.usage = {"prompt_tokens": tokens, "completion_tokens": 1, "cost": 0.001}

    agent_llm = Counting([assistant(call("replace_text", find="end", replace="end")), assistant(content="Done."),
                          assistant(call("replace_text", find="Teh", replace="The")), assistant(content="Fixed.")], 100)
    reviewer_llm = Counting([assistant(content='{"ok": false, "issues": ["b1 still says Teh."]}'),
                             assistant(content='{"ok": true}')], 50)
    results = await run_suite([lazy], lambda: agent_llm, tmp_path, make_review_llm=lambda: reviewer_llm)
    result = results[0]
    assert result.passed and result.reviews == ["1. b1 still says Teh.", None]
    assert (result.prompt_tokens, result.completion_tokens) == (150, 2) and result.cost == 0.002
    assert "1 review fix round" in report.task_line(result)
    assert report.summarize(results)["reviewed_fixes"] == 1
    assert "1 task(s) fixed after review" in report.console("m", results)

    report.write_results(tmp_path / "results.json", "m", results, None, "2026-01-01T00:00:00", review_model="judge/x")
    assert "| `m` (review by judge/x) |" in report.markdown([report.load_results(tmp_path)])
    report.write_results(tmp_path / "results.json", "m", results, None, "2026-01-01T00:00:00", review=False)
    assert "| `m` (no review) |" in report.markdown([report.load_results(tmp_path)])


async def test_a_failed_review_is_recorded(tmp_path):
    result = (await run_suite([TYPOS], lambda: solved(TYPOS, review=None), tmp_path))[0]
    assert result.passed and result.reviews == []
    assert result.review_errors and "ran out of turns" in result.review_errors[0]
    assert "review failed: AssertionError" in report.task_line(result)
