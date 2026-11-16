"""The review step: the harness loop around a Reviewer, and the Word reviewer."""

from __future__ import annotations

from docx_harness import Agent, DocumentReviewer, ToolRegistry, WordDocument, create_agent, default_tools
from docx_harness.agent.prompts import REVIEW_FEEDBACK_PROMPT
from docx_harness.review import describe, parse_review

from conftest import ScriptedLLM, assistant, call


def paragraphs(*texts):
    return call("insert_paragraphs", paragraphs=[{"text": t} for t in texts])


class FakeReviewer:
    """Returns the given feedback in order, then None."""

    def __init__(self, *feedback, fail: Exception | None = None):
        self.feedback = list(feedback)
        self.fail = fail
        self.began: list[str] = []
        self.calls: list[tuple[str, str]] = []

    def begin(self, instruction):
        self.began.append(instruction)

    async def review(self, instruction, summary):
        self.calls.append((instruction, summary))
        if self.fail:
            raise self.fail
        return self.feedback.pop(0) if self.feedback else None


def agent(doc, llm, reviewer, **options):
    events = []
    return Agent(llm, ToolRegistry(default_tools(doc)), reviewer=reviewer, on_event=events.append, **options), events


# ----------------------------------------------------------------- harness

async def test_feedback_goes_back_to_the_model_until_the_review_passes(doc):
    llm = ScriptedLLM([
        assistant(paragraphs("Helo")), assistant(content="Wrote it."),
        assistant(call("replace_text", find="Helo", replace="Hello")), assistant(content="Fixed the typo."),
    ])
    reviewer = FakeReviewer("1. b1 misspells Hello.")
    bot, events = agent(doc, llm, reviewer)
    result = await bot.arun("say hello")

    assert doc.text() == "Hello"
    assert result.status == "completed" and result.final_text == "Fixed the typo."
    assert result.steps == 4
    assert result.reviews == ["1. b1 misspells Hello.", None]
    assert reviewer.began == ["say hello"]
    assert reviewer.calls == [("say hello", "Wrote it."), ("say hello", "Fixed the typo.")]
    assert [e.data for e in events if e.type == "review"] == [
        {"round": 1, "feedback": "1. b1 misspells Hello."}, {"round": 2, "feedback": None}]
    fix_request = llm.requests[2]["messages"][-1]
    assert fix_request == {"role": "user", "content": REVIEW_FEEDBACK_PROMPT.format(feedback="1. b1 misspells Hello.")}


async def test_review_rounds_are_capped(doc):
    llm = ScriptedLLM([
        assistant(paragraphs("a")), assistant(content="1"),
        assistant(paragraphs("b")), assistant(content="2"),
        assistant(paragraphs("c")), assistant(content="3"),
    ])
    reviewer = FakeReviewer("more", "still more", "never asked")
    bot, _ = agent(doc, llm, reviewer, max_reviews=2)
    result = await bot.arun("write")
    assert result.reviews == ["more", "still more"]
    assert result.final_text == "3" and llm.turns == []
    assert len(reviewer.calls) == 2  # the last fix is not reviewed again


async def test_no_review_when_nothing_changed(doc):
    llm = ScriptedLLM([assistant(call("get_outline")), assistant(content="It is empty.")])
    reviewer = FakeReviewer("should not be asked")
    bot, events = agent(doc, llm, reviewer)
    result = await bot.arun("what is in it?")
    assert result.reviews == [] and reviewer.calls == []
    assert not [e for e in events if e.type == "review"]


async def test_fix_round_without_changes_is_not_reviewed_again(doc):
    llm = ScriptedLLM([assistant(paragraphs("Fine")), assistant(content="Done."),
                       assistant(content="That item is wrong; the text is as asked.")])
    reviewer = FakeReviewer("1. Change it.", "never asked")
    bot, _ = agent(doc, llm, reviewer)
    result = await bot.arun("write Fine")
    assert result.reviews == ["1. Change it."]
    assert result.final_text == "That item is wrong; the text is as asked."
    assert len(reviewer.calls) == 1


async def test_failing_reviewer_keeps_the_work(doc):
    llm = ScriptedLLM([assistant(paragraphs("Kept")), assistant(content="Done.")])
    bot, events = agent(doc, llm, FakeReviewer(fail=RuntimeError("no network")))
    result = await bot.arun("write")
    assert result.status == "completed" and doc.text() == "Kept" and result.reviews == []
    assert [e.data for e in events if e.type == "review"] == [
        {"round": 1, "feedback": None, "error": "RuntimeError: no network"}]


async def test_reviews_share_the_step_budget(doc):
    llm = ScriptedLLM([assistant(paragraphs("a")), assistant(content="1"),
                       assistant(paragraphs("b")), assistant(paragraphs("c"))])
    bot, _ = agent(doc, llm, FakeReviewer("again"), max_steps=3)
    result = await bot.arun("write")
    assert result.status == "max_steps" and result.steps == 3
    assert result.reviews == ["again"]


async def test_no_review_round_when_the_agent_ran_out_of_steps(doc):
    llm = ScriptedLLM([assistant(paragraphs("a")), assistant(paragraphs("b"))])
    reviewer = FakeReviewer("never asked")
    bot, _ = agent(doc, llm, reviewer, max_steps=2)
    result = await bot.arun("write")
    assert result.status == "max_steps" and reviewer.calls == []


async def test_plan_mode_reviews_after_execution(doc):
    llm = ScriptedLLM([assistant(content="1. Add a line."),
                       assistant(paragraphs("Line")), assistant(content="Added.")])
    reviewer = FakeReviewer()
    bot, _ = agent(doc, llm, reviewer)
    result = await bot.arun("add a line", plan_mode=True)
    assert result.steps == 3 and result.plan is not None
    assert result.reviews == [None]
    assert reviewer.began == ["add a line"]


# ------------------------------------------------------------ Word reviewer

def test_describe_shows_ids_styles_word_counts_lists_and_tables():
    doc = WordDocument.open()
    doc.insert_paragraph("Report", "Title")
    doc.insert_paragraph("One two three four", None)
    doc.insert_paragraph("Item", "List Bullet")
    doc.insert_table([["A", "B"], ["1", "2"]])
    doc.set_header("Acme")
    text = describe(doc, ids=True, word_counts=True)
    assert "b2 [Normal, 4 words] One two three four" in text
    assert "b3 [List Bullet, bullet level 0, 1 word] Item" in text
    assert "b4 [Table]\n  A | B\n  1 | 2" in text
    assert "[header] Acme" in text
    assert describe(doc).splitlines()[1] == "[Normal] One two three four"


def test_parse_review():
    assert parse_review('{"ok": true}') is None
    assert parse_review('Sure: {"ok": false, "issues": ["b2 is too short.", " ", "b3 has a typed bullet."]}') == (
        "1. b2 is too short.\n2. b3 has a typed bullet.")
    assert parse_review('{"ok": false, "issues": "One problem."}') == "1. One problem."
    assert parse_review('{"ok": false, "issues": []}') is None
    assert parse_review("looks good to me") is None
    assert parse_review("{not json}") is None
    assert parse_review(None) is None


async def test_document_reviewer_sends_before_now_and_summary():
    doc = WordDocument.open()
    doc.insert_paragraph("Old text here")
    llm = ScriptedLLM([assistant(content='{"ok": false, "issues": ["b1 still says Old."]}')])
    reviewer = DocumentReviewer(doc, llm)
    reviewer.begin("make it new")
    doc.replace_text("Old", "New")
    assert await reviewer.review("make it new", "Changed it.") == "1. b1 still says Old."
    request = llm.requests[0]
    assert request["tools"] is None
    prompt = request["messages"][1]["content"]
    assert "REQUEST:\nmake it new" in prompt
    assert "DOCUMENT BEFORE:\nb1 [Normal, 3 words] Old text here" in prompt
    assert "DOCUMENT NOW:\nb1 [Normal, 3 words] New text here" in prompt
    assert prompt.endswith("ASSISTANT'S SUMMARY:\nChanged it.")


async def test_document_reviewer_on_a_new_document():
    doc = WordDocument.open()
    llm = ScriptedLLM([assistant(content='{"ok": true}')])
    reviewer = DocumentReviewer(doc, llm)
    reviewer.begin("write")
    doc.insert_paragraph("Hi")
    assert await reviewer.review("write", "") is None
    assert "DOCUMENT BEFORE:\n(new, empty document)" in llm.requests[0]["messages"][1]["content"]


async def test_create_agent_reviews_by_default_with_the_agent_client(doc):
    llm = ScriptedLLM([assistant(paragraphs("Hi")), assistant(content="Done."), assistant(content='{"ok": true}')])
    result = await create_agent(doc, llm=llm).arun("greet")
    assert result.reviews == [None] and llm.turns == []
    assert "You review the work" in llm.requests[2]["messages"][0]["content"]


async def test_create_agent_review_options(doc):
    main = ScriptedLLM([assistant(paragraphs("Hi")), assistant(content="Done.")])
    other = ScriptedLLM([assistant(content='{"ok": true}')])
    result = await create_agent(doc, llm=main, review_llm=other).arun("greet")
    assert result.reviews == [None] and other.turns == [] and len(main.requests) == 2

    plain = ScriptedLLM([assistant(paragraphs("Bye")), assistant(content="Done.")])
    bot = create_agent(doc, llm=plain, review=False)
    assert bot.reviewer is None
    assert (await bot.arun("bye")).reviews == []


def test_cli_review_flags_and_output(monkeypatch, capsys):
    from docx_harness import cli
    from docx_harness.agent import AgentEvent, AgentResult

    seen = {}

    def fake_draft(instruction, input_path, output_path, **options):
        seen.update(options)
        for data in ({"round": 1, "feedback": "1. b2 is too short."}, {"round": 2, "feedback": None}):
            options["on_event"](AgentEvent("review", data))
        return AgentResult("Done.", 3)

    monkeypatch.setattr(cli, "draft", fake_draft)
    assert cli.main(["write", "-o", "x.docx"]) == 0
    err = capsys.readouterr().err
    assert "review found problems; fixing:\n    1. b2 is too short." in err and "review: no problems found" in err
    assert seen["review"] is True and seen["review_llm"] is None

    assert cli.main(["write", "--no-review"]) == 0
    assert seen["review"] is False


def test_describe_for_review_shows_formatting_links_shading_and_revisions():
    doc = WordDocument.open()
    doc.insert_paragraph("A WARNING here", "Heading 1")
    doc.format_text("b1", "WARNING", bold=True, color="FF0000")
    doc.insert_paragraph("See the handbook now")
    doc.add_hyperlink("b2", "https://x.com", match="handbook")
    doc.insert_paragraph("")
    table = doc.insert_table([["City", "Staff"], ["Lisbon", "42"]])
    doc.format_cells(table + "[0,*]", shading="D9D9D9")
    doc.track_changes = True
    doc.edits.author = "Bob"
    doc.replace_text("See", "Read")
    text = describe(doc, ids=True, word_counts=True, formatting=True)
    assert "b1 [heading 1, 3 words] A «WARNING»{bold, color FF0000} here" in text
    assert "b2 [Normal, 4 words] «Read»{inserted} the «handbook»{link https://x.com} now" in text
    assert "b3 [Normal, 0 words] (empty)" in text
    assert "b4 [Table, 2 rows x 2 columns]\n  row 0: City | Staff  (shaded D9D9D9)\n  row 1: Lisbon | 42" in text
    assert "  deletion by Bob in b2: 'See'" in text and "  insertion by Bob in b2: 'Read'" in text
    assert text.endswith("Original text, with every pending change rejected:\n"
                         "  [heading 1] A WARNING here\n  [Normal] See the handbook now (link: handbook -> https://x.com)\n"
                         "  [Normal] \n  [Table]\n    City | Staff\n    Lisbon | 42")
    assert describe(doc).count("Original text") == 0
