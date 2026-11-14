"""Tracked changes: edits recorded as Word revisions (w:ins / w:del / *PrChange)."""

from __future__ import annotations

import pytest

from docx_harness import ToolRegistry, WordDocument, WordPackage, create_agent, default_tools
from docx_harness.agent.prompts import TRACK_CHANGES_PROMPT
from docx_harness.document import styling
from docx_harness.document.ooxml import DEL, INS, P, PPR_ORDER, PROPERTY_CHANGES, RPR_ORDER, iter_runs, qn, run_text

from conftest import ScriptedLLM, assistant, call, make_docx, reopen_with_python_docx

DATE = "2026-01-02T03:04:05Z"
MIXED = (
    "<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>big wor</w:t></w:r>"
    "<w:r><w:rPr><w:i/></w:rPr><w:t>ld</w:t></w:r><w:r><w:t>!</w:t></w:r></w:p>"
)
OTHER = 'w:author="Alice" w:date="2025-01-01T00:00:00Z"'


def tracked(body_xml: str | None = None, author: str = "Agent") -> WordDocument:
    package = WordPackage.from_bytes(make_docx(body_xml)) if body_xml is not None else WordPackage.open()
    doc = WordDocument(package, track_changes=True, author=author)
    doc.edits.date = DATE
    return doc


def xml(doc: WordDocument) -> str:
    return doc.read_part_xml("word/document.xml")


def body(doc: WordDocument):
    return doc._body


def check(doc: WordDocument, tmp_path):
    """Revision ids are unique, nothing nests illegally, and the file reopens."""
    ids = [e.get(qn("w:id")) for e in body(doc).iter(INS, DEL, qn("w:rPrChange"), qn("w:pPrChange"))]
    assert len(ids) == len(set(ids))
    for wrapper in body(doc).iter(INS):
        assert not any(a.tag == INS for a in wrapper.iterancestors())
    for wrapper in body(doc).iter(DEL):
        assert not any(a.tag == DEL for a in wrapper.iterancestors())
        assert not list(wrapper.iter(qn("w:t")))  # deleted runs hold w:delText only
    return reopen_with_python_docx(doc, tmp_path)


def test_replace_is_recorded_as_delete_then_insert(tmp_path):
    doc = tracked(MIXED)
    assert doc.replace_text("big world", "small planet") == 1
    assert doc.text() == "Hello small planet!"
    out = xml(doc)
    assert out.count("<w:del ") == 2 and out.count("<w:ins ") == 1  # two runs deleted, one inserted
    assert "<w:delText" in out and 'w:author="Agent"' in out and DATE in out
    assert out.index("<w:del ") < out.index("<w:ins ")
    inserted = next(body(doc).iter(INS))[0]
    assert inserted.find(qn("w:rPr")).find(qn("w:b")) is not None  # formatting of the replaced text
    revisions = doc.revisions()
    assert [r["type"] for r in revisions] == ["delete", "delete", "insert"]
    assert "".join(r["text"] for r in revisions if r["type"] == "delete") == "big world"
    assert revisions[-1] == {"id": revisions[-1]["id"], "type": "insert", "author": "Agent", "date": DATE,
                             "block": "b1", "text": "small planet"}
    check(doc, tmp_path)


def test_untracked_default_leaves_no_revisions():
    doc = WordDocument(WordPackage.from_bytes(make_docx(MIXED)))
    doc.replace_text("big world", "small planet")
    doc.delete_block(doc.insert_paragraph("x"))
    assert doc.revisions() == [] and "w:del" not in xml(doc)


def test_inserted_blocks_are_marked(tmp_path):
    doc = tracked()
    p = doc.insert_paragraph("New paragraph", "Heading 1")
    table = doc.insert_table([["A", "B"], ["1", "2"]])
    assert doc.insert_table_row(table, ["3", "4"]) == 2
    paragraph = doc._get(p)
    assert paragraph.find(qn("w:pPr")).find(qn("w:rPr"))[0].tag == INS
    assert [child.tag for child in paragraph][1:] == [INS]
    rows = doc._get(table).findall(qn("w:tr"))
    assert len(rows) == 3 and all(tr.find(qn("w:trPr")).find(INS) is not None for tr in rows)
    assert doc.read(table)["blocks"][0]["rows"] == [["A", "B"], ["1", "2"], ["3", "4"]]
    reopened = check(doc, tmp_path)
    assert len(reopened.tables[0].rows) == 3  # python-docx does not read text inside w:ins


def test_deleted_blocks_stay_in_the_file_but_leave_the_document(tmp_path):
    doc = tracked("<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>"
                  "<w:p><w:r><w:t>three</w:t></w:r></w:p>")
    doc.delete_block("b2")
    assert [e["id"] for e in doc.outline()] == ["b1", "b3"]
    assert doc.text().split() == ["one", "three"] and len(doc) == 2
    assert "two</w:delText>" in xml(doc)
    with pytest.raises(ValueError):
        doc.set_paragraph_text("b2", "back")
    assert doc.insert_paragraph("after one", after="b1") == "b4"
    assert [e["id"] for e in doc.outline()] == ["b1", "b4", "b3"]
    assert doc.search("two") == []
    check(doc, tmp_path)
    # Reloading the saved file hides the deleted paragraph too.
    doc.save(tmp_path / "again.docx")
    assert WordDocument.open(tmp_path / "again.docx").text().split() == ["one", "after", "one", "three"]


def test_deleting_your_own_insertion_removes_it():
    doc = tracked("<w:p><w:r><w:t>keep</w:t></w:r></w:p>")
    before = xml(doc)
    new = doc.insert_paragraph("temporary")
    doc.replace_text("temporary", "changed", ref=new)
    assert doc.text() == "keep\nchanged" and "<w:del " not in xml(doc)
    doc.delete_block(new)
    table = doc.insert_table([["a"], ["b"]])
    doc.delete_table_row(table, 1)
    doc.delete_block(table)
    assert xml(doc) == before and doc.revisions() == []


def test_table_rows_track_and_indexes_skip_deleted_rows(tmp_path):
    cell = "<w:tc><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:tc>"
    rows = "".join(f"<w:tr>{cell.format(f'r{i}a')}{cell.format(f'r{i}b')}</w:tr>" for i in range(3))
    doc = tracked(f"<w:tbl>{rows}</w:tbl>")
    doc.delete_table_row("b1", 1)
    assert doc.read("b1")["blocks"][0]["rows"] == [["r0a", "r0b"], ["r2a", "r2b"]]
    doc.set_cell_text("b1", 1, 0, "changed")  # row 1 is now the old row 2
    assert doc.insert_table_row("b1", ["x", "y"], index=1) == 1
    assert doc.read("b1")["blocks"][0]["rows"] == [["r0a", "r0b"], ["x", "y"], ["changed", "r2b"]]
    out = xml(doc)
    assert "r1a</w:delText>" in out and "r2a</w:delText>" in out
    new_row = doc._live_rows(doc._get("b1"))[1]
    assert not list(new_row.iter(DEL)) and new_row.find(qn("w:trPr")).find(INS) is not None
    assert doc.text().split("\n") == ["r0a", "r0b", "x", "y", "changed", "r2b"]  # no trace of deleted rows
    doc.set_cell_text("b1", 0, 0, "two\nlines")
    assert doc.read("b1")["blocks"][0]["rows"][0][0] == "two\nlines"
    doc.delete_block("b1")
    assert doc.text() == "" and len(doc) == 0 and len(doc._body.findall(qn("w:tbl"))) == 1
    check(doc, tmp_path)


def test_other_authors_revisions_are_kept_and_never_nested(tmp_path):
    doc = tracked(
        "<w:p><w:r><w:t>Start </w:t></w:r>"
        f'<w:ins w:id="7" {OTHER}><w:r><w:t>alpha beta gamma</w:t></w:r></w:ins>'
        f'<w:del w:id="8" {OTHER}><w:r><w:delText>gone</w:delText></w:r></w:del>'
        "<w:r><w:t> end</w:t></w:r></w:p>"
    )
    doc.replace_text("beta", "BETA")
    assert doc.text() == "Start alpha BETA gamma end"
    out = xml(doc)
    assert "gone</w:delText>" in out and out.count('w:author="Alice"') == 3  # her insertion was split in two
    mine = [r for r in doc.revisions() if r["author"] == "Agent"]
    assert [(r["type"], r["text"]) for r in mine] == [("delete", "beta"), ("insert", "BETA")]
    assert all(int(r["id"]) > 8 for r in mine)
    deleted = next(d for d in body(doc).iter(DEL) if d.get(qn("w:author")) == "Agent")
    assert deleted.getparent().tag == INS  # deleting inside her insertion is a del within her ins
    check(doc, tmp_path)


def test_update_paragraph_revises_only_the_words_that_changed(tmp_path):
    doc = tracked("<w:p><w:r><w:t>The quick brown fox jumps over the lazy dog.</w:t></w:r></w:p>")
    doc.set_paragraph_text("b1", "The quick red fox leaps over the lazy dog today.")
    assert doc.text() == "The quick red fox leaps over the lazy dog today."
    changes = [(r["type"], r["text"]) for r in doc.revisions()]
    assert changes == [("delete", "brown"), ("insert", "red"), ("delete", "jumps"), ("insert", "leaps"),
                       ("insert", " today")]
    check(doc, tmp_path)

    doc = tracked("<w:p><w:r><w:t>Completely different words here.</w:t></w:r></w:p><w:p/>")
    doc.set_paragraph_text("b1", "Nothing in common at all, really")
    doc.set_paragraph_text("b2", "filled in")
    assert doc.text() == "Nothing in common at all, really\nfilled in"
    assert [(r["type"], r["text"]) for r in doc.revisions()] == [
        ("delete", "Completely different words here."), ("insert", "Nothing in common at all, really"),
        ("insert", "filled in"),
    ]


def test_formatting_changes_keep_the_previous_properties(tmp_path):
    doc = tracked(MIXED)
    doc.format_text("b1", match="big", italic=True, color="FF0000")
    doc.format_text("b1", match="big", size_pt=14)  # second change to the same run: one record
    doc.set_paragraph_format("b1", alignment="center")
    doc.set_paragraph_style("b1", "Heading 1")
    run = next(r for r in doc._get("b1").iter(qn("w:r")) if "".join(r.itertext()) == "big")
    rpr = run.find(qn("w:rPr"))
    assert [c.tag for c in rpr] == sorted((c.tag for c in rpr), key=[qn(t) for t in RPR_ORDER].index)
    assert rpr[-1].tag == qn("w:rPrChange")
    assert [c.tag for c in rpr[-1][0]] == [qn("w:b")]  # what it was before
    ppr = doc._get("b1").find(qn("w:pPr"))
    assert [c.tag for c in ppr] == sorted((c.tag for c in ppr), key=[qn(t) for t in PPR_ORDER].index)
    assert ppr[-1].tag == qn("w:pPrChange") and len(ppr[-1][0]) == 0  # no properties before
    assert [r["type"] for r in doc.revisions()] == ["format", "format"]
    # Text inserted in this session needs no formatting history.
    new = doc.insert_paragraph("fresh")
    doc.format_text(new, bold=True)
    doc.set_paragraph_format(new, alignment="right")
    assert len([r for r in doc.revisions() if r["type"] == "format"]) == 2
    # A replacement does not inherit the history of the run it copies.
    doc.replace_text("big", "large")
    assert xml(doc).count("<w:rPrChange") == 1
    check(doc, tmp_path)


def test_page_break_empty_and_multiline_text(tmp_path):
    doc = tracked()
    doc.insert_paragraph("line one\nline two\tend")
    doc.insert_page_break()
    doc.insert_paragraph("")
    assert doc.text().split("\n")[:2] == ["line one", "line two\tend"]
    doc.replace_text("one\nline", "1 / line")
    assert doc.text().startswith("line 1 / line two")
    check(doc, tmp_path)


def test_tools_and_agent_use_tracking():
    doc = tracked(MIXED)
    registry = ToolRegistry(default_tools(doc))
    assert registry.execute("list_revisions", {}) == "There are no tracked changes in the document."
    registry.execute("replace_text", {"find": "Hello", "replace": "Goodbye"})
    registry.execute("delete_blocks", {"block_ids": [doc.insert_paragraph("x")]})
    listing = registry.execute("list_revisions", {})
    first, second = (r["id"] for r in doc.revisions())
    assert listing.splitlines() == [f"[{first}] delete by Agent in b1: 'Hello'",
                                    f"[{second}] insert by Agent in b1: 'Goodbye'"]
    assert registry.get("list_revisions").read_only

    llm = ScriptedLLM([assistant(call("replace_text", find="Goodbye", replace="Hi")), assistant(content="done")])
    create_agent(doc, llm=llm).run("greet differently")
    assert TRACK_CHANGES_PROMPT in llm.requests[0]["messages"][0]["content"]
    assert doc.text() == "Hi big world!"

    plain = ScriptedLLM([assistant(content="done")])
    create_agent(WordDocument.open(), llm=plain).run("hi")
    assert TRACK_CHANGES_PROMPT not in plain.requests[0]["messages"][0]["content"]


# ------------------------------------------------- the switch in settings.xml

SETTINGS = "word/settings.xml"


def test_tracking_switches_on_track_revisions_in_settings(tmp_path):
    doc = tracked(MIXED)
    assert doc.track_changes is True
    children = [child.tag for child in doc.package.xml(SETTINGS)]
    assert children[0] == qn("w:trackRevisions")  # before w:defaultTabStop, as the schema requires
    assert children.count(qn("w:trackRevisions")) == 1
    reopen_with_python_docx(doc, tmp_path)

    doc.track_changes = False
    assert doc.edits.track_changes is False
    assert "trackRevisions" not in doc.read_part_xml(SETTINGS)
    doc.replace_text("Hello", "Goodbye")
    assert not doc.revisions()


def test_untracked_documents_keep_their_settings_untouched():
    data = make_docx(MIXED)
    original = WordPackage.from_bytes(data).part_bytes(SETTINGS)
    doc = WordDocument(WordPackage.from_bytes(data))
    doc.replace_text("Hello", "Goodbye")
    assert doc.track_changes is False
    assert doc.package.part_bytes(SETTINGS) == original


def test_track_revisions_goes_after_the_settings_that_precede_it():
    settings = (
        f'<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:zoom w:percent="100"/><w:trackRevisions w:val="false"/><w:defaultTabStop w:val="720"/></w:settings>'
    ).encode()
    doc = WordDocument(WordPackage.from_bytes(make_docx(MIXED, {SETTINGS: settings})), track_changes=True)
    flag = doc.package.xml(SETTINGS).find(qn("w:trackRevisions"))
    assert flag.get(qn("w:val")) is None  # an explicit "false" is turned on
    doc.track_changes = False
    doc.track_changes = True
    assert [c.tag for c in doc.package.xml(SETTINGS)] == [qn("w:zoom"), qn("w:trackRevisions"), qn("w:defaultTabStop")]


def test_tracking_works_without_a_settings_part(tmp_path):
    from conftest import template_parts, zip_parts

    parts = template_parts()
    del parts[SETTINGS]
    doc = WordDocument(WordPackage.from_bytes(zip_parts(parts)), track_changes=True)
    doc.insert_paragraph("Tracked anyway")
    assert [r["type"] for r in doc.revisions()] == ["insert", "insert"]
    doc.save(tmp_path / "out.docx")


# ------------------------------------------------------- accepting, rejecting

THREE = "".join(f"<w:p><w:r><w:t>{word}</w:t></w:r></w:p>" for word in ("one", "two", "three"))
CELL = "<w:tc><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:tc>"
TABLE = "<w:tbl>" + "".join(f"<w:tr>{CELL.format(f'r{i}a')}{CELL.format(f'r{i}b')}</w:tr>" for i in range(3)) + "</w:tbl>"


def view(doc: WordDocument) -> list:
    """What the document says and how it is formatted, whatever the run boundaries."""
    out = []
    for element in doc._block_elements():
        info = doc._describe(element)
        if info["type"] != "paragraph":
            out.append((info["type"], info.get("rows") or info.get("text")))
            continue
        pieces: list[list] = []
        for run in iter_runs(element):
            fmt, text = styling.run_format(run), run_text(run)
            if not text:
                continue
            if pieces and pieces[-1][1] == fmt:
                pieces[-1][0] += text
            else:
                pieces.append([text, fmt])
        ppr = element.find(qn("w:pPr"))
        layout = sorted(c.tag for c in ppr if c.tag not in (qn("w:rPr"), qn("w:pStyle"), qn("w:pPrChange"))) if ppr is not None else []
        out.append(("paragraph", info["style"], layout, pieces))
    return out


def markup(doc: WordDocument) -> list:
    return list(body(doc).iter(INS, DEL, *PROPERTY_CHANGES))


def _replace(doc):
    doc.replace_text("big world", "small planet")
    doc.replace_text("Hello ", "")


def _insert(doc):
    doc.insert_paragraph("New heading", "Heading 1", after="b1")
    table = doc.insert_table([["A", "B"], ["1", "2"]], after="start")
    doc.insert_table_row(table, ["3", "4"])
    doc.insert_paragraph("last\nwith a break")
    doc.insert_page_break()
    doc.insert_paragraph("")


def _delete(doc):
    doc.delete_block("b2")
    doc.delete_block("b4")
    doc.delete_block("b5")


def _rows(doc):
    doc.delete_table_row("b1", 1)
    doc.set_cell_text("b1", 1, 0, "changed")
    doc.insert_table_row("b1", ["x", "y"], index=1)
    doc.set_cell_text("b1", 0, 0, "two\nlines")


def _rewrite(doc):
    doc.set_paragraph_text("b1", "Hello big wide world, again!")
    doc.set_paragraph_text("b2", "Nothing in common at all")


def _format(doc):
    doc.format_text("b1", match="big", italic=True, color="FF0000")
    doc.format_text("b2", bold=True)
    doc.set_paragraph_format("b1", alignment="center")
    doc.set_paragraph_style("b2", "Heading 1")
    doc.replace_text("big", "large")


SCENARIOS = {
    "replace": (MIXED, _replace),
    "insert": (THREE, _insert),
    "delete": (THREE + TABLE + "<w:p/>", _delete),
    "rows": (TABLE, _rows),
    "rewrite": (MIXED + THREE, _rewrite),
    "format": (MIXED + THREE, _format),
}


@pytest.mark.parametrize("name", SCENARIOS)
def test_reject_all_restores_the_original_and_accept_all_keeps_the_new(name, tmp_path):
    source, edit = SCENARIOS[name]
    original = view(WordDocument(WordPackage.from_bytes(make_docx(source))))

    doc = tracked(source)
    edit(doc)
    pending = len(doc.revisions())
    assert pending and doc.reject_revisions() == pending
    assert view(doc) == original and not markup(doc) and doc.revisions() == []
    reopened = reopen_with_python_docx(doc, tmp_path)
    assert [p.text for p in reopened.paragraphs] == [paragraph_text_of(b) for b in view(doc) if b[0] == "paragraph"]

    doc = tracked(source)
    edit(doc)
    live, text = view(doc), doc.text()
    assert doc.accept_revisions() == pending
    assert view(doc) == live and doc.text() == text and not markup(doc)
    assert live != original
    reopened = reopen_with_python_docx(doc, tmp_path)
    assert [p.text for p in reopened.paragraphs] == [paragraph_text_of(b) for b in view(doc) if b[0] == "paragraph"]
    # Every block that is left can be addressed again.
    assert [e["id"] for e in doc.outline()] == [doc._id_of(e) for e in doc._block_elements()]
    for entry in doc.outline():
        doc._get(entry["id"])


def paragraph_text_of(block) -> str:
    text = "".join(text for text, _ in block[3])
    return text if text.strip("\n") else ""  # python-docx reads a page break as no text


def test_accept_and_reject_select_by_id_author_and_block():
    source = (
        "<w:p><w:r><w:t>Start </w:t></w:r>"
        f'<w:ins w:id="7" {OTHER}><w:r><w:t>alpha </w:t></w:r></w:ins>'
        f'<w:del w:id="8" {OTHER}><w:r><w:delText>gone </w:delText></w:r></w:del>'
        "<w:r><w:t>end</w:t></w:r></w:p>" + THREE
    )
    doc = tracked(source)
    doc.replace_text("end", "finish")
    doc.replace_text("two", "2")
    assert doc.text().split("\n") == ["Start alpha finish", "one", "2", "three"]

    assert doc.reject_revisions(author="Alice") == 2
    assert doc.text().split("\n")[0] == "Start gone finish"
    assert {r["author"] for r in doc.revisions()} == {"Agent"}

    assert doc.accept_revisions(block="b1") == 2
    assert doc.text().split("\n")[0] == "Start gone finish" and [r["block"] for r in doc.revisions()] == ["b3", "b3"]

    deletion = next(r["id"] for r in doc.revisions() if r["type"] == "delete")
    assert doc.reject_revisions(ids=[deletion]) == 1
    assert doc.text().split("\n")[2] == "two2"
    with pytest.raises(ValueError, match="No pending revision with id 999"):
        doc.accept_revisions(ids=["999"])
    assert doc.reject_revisions() == 1 and doc.text().split("\n")[2] == "two"
    assert doc.accept_revisions() == 0


def test_rejecting_a_deletion_brings_the_block_back_with_an_id():
    doc = tracked(THREE + TABLE)
    doc.delete_block("b2")
    doc.delete_table_row("b4", 0)
    doc.delete_block("b4")
    assert [e["id"] for e in doc.outline()] == ["b1", "b3"]
    doc.reject_revisions()
    outline = doc.outline()
    assert [e["preview"] for e in outline[:3]] == ["one", "two", "three"] and outline[3]["type"] == "table"
    assert [e["id"] for e in outline] == ["b1", "b5", "b3", "b6"]
    doc.set_paragraph_text("b5", "two again")
    assert len(doc.read("b6")["blocks"][0]["rows"]) == 3


def test_a_paragraph_mark_revision_joins_or_splits_paragraphs():
    split = (  # Alice pressed Enter in the middle of "one two"
        f'<w:p><w:pPr><w:rPr><w:ins w:id="3" {OTHER}/></w:rPr></w:pPr><w:r><w:t>one </w:t></w:r></w:p>'
        "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>two</w:t></w:r></w:p>"
    )
    doc = tracked(split)
    assert doc.reject_revisions() == 1
    assert doc.text() == "one two" and len(doc) == 1
    assert doc._block_elements()[0].find(qn("w:pPr")).find(qn("w:jc")) is not None
    doc = tracked(split)
    doc.accept_revisions()
    assert doc.text() == "one \ntwo" and not markup(doc)

    joined = split.replace("w:ins", "w:del")  # Alice deleted the paragraph break instead
    doc = tracked(joined)
    doc.accept_revisions()
    assert doc.text() == "one two"
    doc = tracked(joined)
    doc.reject_revisions()
    assert doc.text() == "one \ntwo" and not markup(doc)


def test_revision_tools():
    doc = tracked(MIXED)
    registry = ToolRegistry(default_tools(doc))
    registry.execute("replace_text", {"find": "Hello", "replace": "Goodbye"})
    registry.execute("format_text", {"blocks": ["b1"], "match": "!", "bold": True})
    assert registry.execute("accept_revisions", {}).startswith("Error: Say which revisions")
    assert len(doc.revisions()) == 3
    insertion = next(r["id"] for r in doc.revisions() if r["type"] == "insert")
    assert registry.execute("reject_revisions", {"ids": [insertion]}).startswith("Rejected 1 revision(s); 2 still pending")
    assert registry.execute("reject_revisions", {"ids": [insertion]}).startswith("Error: No pending revision with id")
    assert registry.execute("accept_revisions", {"author": "Nobody"}) == "No tracked changes matched; nothing changed."
    assert registry.execute("accept_revisions", {"all_revisions": True}).startswith("Accepted 2 revision(s); 0 still pending")
    assert doc.text() == " big world!" and not markup(doc)


# ------------------------------------------------- headers, footers and notes

from conftest import W, template_parts  # noqa: E402

HEADER, FOOTER, NOTES = "word/header1.xml", "word/footer1.xml", "word/footnotes.xml"


def story_parts() -> dict[str, bytes]:
    parts = template_parts()
    kinds = {"header1": "header", "footer1": "footer", "footnotes": "footnotes"}
    rels = "".join(
        f'<Relationship Id="rId9{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}"'
        f' Target="{name}.xml"/>' for i, (name, kind) in enumerate(kinds.items())
    )
    types = "".join(
        f'<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        f'wordprocessingml.{kind}+xml"/>' for name, kind in kinds.items()
    )
    declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    return {
        "word/_rels/document.xml.rels": parts["word/_rels/document.xml.rels"].replace(
            b"</Relationships>", rels.encode() + b"</Relationships>"),
        "[Content_Types].xml": parts["[Content_Types].xml"].replace(b"</Types>", types.encode() + b"</Types>"),
        HEADER: f"{declaration}<w:hdr {W}><w:p><w:r><w:t>Draft 2024 report</w:t></w:r></w:p></w:hdr>".encode(),
        FOOTER: (f"{declaration}<w:ftr {W}><w:p><w:r><w:t>Page </w:t></w:r>"
                 f'<w:ins w:id="40" {OTHER}><w:r><w:t>of many</w:t></w:r></w:ins></w:p></w:ftr>').encode(),
        NOTES: (f'{declaration}<w:footnotes {W}><w:footnote w:type="separator" w:id="-1"><w:p><w:r>'
                f"<w:t>2024 separator</w:t></w:r></w:p></w:footnote>"
                f'<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t> See the 2024 survey.</w:t></w:r></w:p>'
                f"</w:footnote></w:footnotes>").encode(),
    }


def with_stories(track: bool = True) -> WordDocument:
    package = WordPackage.from_bytes(make_docx("<w:p><w:r><w:t>Results for 2024.</w:t></w:r></w:p>", story_parts()))
    doc = WordDocument(package, track_changes=track, author="Agent")
    doc.edits.date = DATE
    return doc


def test_stories_are_read_searched_and_replaced_as_tracked_changes(tmp_path):
    doc = with_stories()
    assert doc.stories() == [
        {"part": HEADER, "kind": "header", "paragraphs": ["Draft 2024 report"]},
        {"part": FOOTER, "kind": "footer", "paragraphs": ["Page of many"]},
        {"part": NOTES, "kind": "footnotes", "paragraphs": [" See the 2024 survey."]},  # no separator
    ]
    assert [hit["ref"] for hit in doc.search("2024")] == ["b1", HEADER, NOTES]
    footer_before = doc.package.part_bytes(FOOTER)

    assert doc.replace_text("2024", "2025") == 3  # body, header and footnote; the separator is left alone
    assert doc.text() == "Results for 2025."
    assert [s["paragraphs"] for s in doc.stories()] == [["Draft 2025 report"], ["Page of many"], [" See the 2025 survey."]]
    assert "2024 separator" in doc.read_part_xml(NOTES) and "<w:footnoteRef/>" in doc.read_part_xml(NOTES)
    assert "2024</w:delText>" in doc.read_part_xml(HEADER) and "<w:ins " in doc.read_part_xml(HEADER)
    assert doc.package.part_bytes(FOOTER) == footer_before  # untouched parts are saved as they were

    revisions = doc.revisions()
    ids = [r["id"] for r in revisions]
    assert len(ids) == len(set(ids)) == 7 and all(int(i) > 40 for i in ids if i != "40")
    assert [(r["block"], r.get("part")) for r in revisions] == [
        ("b1", None), ("b1", None), (None, HEADER), (None, HEADER), (None, FOOTER), (None, NOTES), (None, NOTES)]
    reopen_with_python_docx(doc, tmp_path)

    assert doc.replace_text("report", "paper", ref=HEADER) == 1 and doc.replace_text("report", "x", ref=FOOTER) == 0
    assert doc.stories()[0]["paragraphs"] == ["Draft 2025 paper"]


def test_accept_and_reject_reach_headers_and_notes(tmp_path):
    doc = with_stories()
    doc.replace_text("2024", "2025")
    assert doc.reject_revisions(author="Agent") == 6
    assert doc.text() == "Results for 2024." and doc.stories()[0]["paragraphs"] == ["Draft 2024 report"]
    assert doc.stories()[2]["paragraphs"] == [" See the 2024 survey."]
    assert [r["part"] for r in doc.revisions()] == [FOOTER]  # Alice's insertion is still pending
    with pytest.raises(ValueError):
        doc.accept_revisions(block="b1", ids=["40"])  # that revision is not in b1
    assert doc.accept_revisions() == 1
    assert doc.revisions() == [] and "w:ins" not in doc.read_part_xml(FOOTER)
    assert doc.stories()[1]["paragraphs"] == ["Page of many"]
    reopen_with_python_docx(doc, tmp_path)


def test_story_tools():
    doc = with_stories()
    registry = ToolRegistry(default_tools(doc))
    listing = registry.execute("read_headers_footnotes", {})
    assert listing.splitlines() == [f"{HEADER} (header):", "  Draft 2024 report", f"{FOOTER} (footer):", "  Page of many",
                                    f"{NOTES} (footnotes):", "   See the 2024 survey."]
    assert registry.execute("replace_text", {"find": "Draft", "replace": "Final", "block": HEADER}) == "Replaced 1 occurrence(s)."
    assert f"in {HEADER}: 'Final'" in registry.execute("list_revisions", {})
    plain = ToolRegistry(default_tools(WordDocument.open()))
    assert plain.execute("read_headers_footnotes", {}).startswith("The document has no headers, footers, footnotes or endnotes.")


def test_untracked_story_edits_and_untouched_parts():
    doc = with_stories(track=False)
    before = {name: doc.package.part_bytes(name) for name in (HEADER, FOOTER, NOTES)}
    doc.replace_text("Results", "Findings")  # body only
    assert {name: doc.package.part_bytes(name) for name in before} == before
    doc.replace_text("Draft", "Final")
    assert doc.stories()[0]["paragraphs"] == ["Final 2024 report"] and doc.revisions()[0]["author"] == "Alice"


# ---------------------------------------------------------------------- moves

def order(doc: WordDocument) -> list[tuple[str, str]]:
    return [(e["id"], e["preview"]) for e in doc.outline()]


def test_untracked_move_keeps_ids_and_leaves_no_markup():
    doc = WordDocument(WordPackage.from_bytes(make_docx(THREE)))
    doc.move_block("b3", "start")
    doc.move_block("b1", "end")
    assert order(doc) == [("b3", "three"), ("b2", "two"), ("b1", "one")]
    assert doc._body[-1].tag == qn("w:sectPr") and "move" not in xml(doc)
    before = xml(doc)
    doc.move_block("b2", "b3")  # already there
    doc.move_block("b1", "end")
    doc.move_block("b3", "b3")
    assert xml(doc) == before
    with pytest.raises(ValueError):
        doc.move_block("b9", "start")


def test_tracked_move_keeps_both_halves_under_one_name(tmp_path):
    doc = tracked(THREE)
    doc.move_block("b3", "start")
    assert order(doc) == [("b3", "three"), ("b1", "one"), ("b2", "two")] and doc.text() == "three\none\ntwo"
    paragraphs = body(doc).findall(P)
    assert len(paragraphs) == 4  # the original is still there, hidden, as the "moved from" half
    to, source = paragraphs[0], paragraphs[3]
    assert [c.tag for c in to] == [qn(t) for t in ("w:pPr", "w:moveToRangeStart", "w:moveTo", "w:moveToRangeEnd")]
    assert [c.tag for c in source] == [qn(t) for t in ("w:pPr", "w:moveFromRangeStart", "w:moveFrom", "w:moveFromRangeEnd")]
    assert to[1].get(qn("w:name")) == source[1].get(qn("w:name"))
    assert to[1].get(qn("w:id")) == to[3].get(qn("w:id")) != source[1].get(qn("w:id"))
    assert "three</w:t>" in xml(doc) and "delText" not in xml(doc)  # moved text is not deleted text
    ids = [e.get(qn("w:id")) for e in body(doc).iter() if e.get(qn("w:id"))]
    assert len(set(ids)) == len(ids) - 2  # only each range end repeats its start
    moves = doc.revisions()
    assert [(r["type"], r["move"], r["block"], r["text"]) for r in moves] == [
        ("move", "to", "b3", ""), ("move", "to", "b3", "three"), ("move", "from", None, ""), ("move", "from", None, "three")]

    doc.replace_text("three", "3")  # the moved paragraph is edited like any other
    assert doc.text() == "3\none\ntwo"
    doc.move_block("b3", "b1")  # moving it again relocates the same half
    assert order(doc) == [("b1", "one"), ("b3", "3"), ("b2", "two")] and len(body(doc).findall(P)) == 4
    reopen_with_python_docx(doc, tmp_path)
    doc.save(tmp_path / "again.docx")
    assert WordDocument.open(tmp_path / "again.docx").text() == "one\n3\ntwo"


def test_moves_are_accepted_and_rejected_as_a_pair(tmp_path):
    original = view(WordDocument(WordPackage.from_bytes(make_docx(THREE + TABLE))))

    def moved() -> WordDocument:
        doc = tracked(THREE + TABLE)
        doc.move_block("b1", "end")
        doc.move_block("b4", "start")  # the table: a deletion plus an insertion
        return doc

    doc = moved()
    assert [e["id"] for e in doc.outline()] == ["b4", "b2", "b3", "b1"] and doc.outline()[0]["type"] == "table"
    assert "moveFrom" not in "".join(c.tag for c in body(doc).findall(qn("w:tbl")) for c in c.iter())
    live = view(doc)
    doc.accept_revisions()
    assert view(doc) == live and not markup(doc) and "move" not in xml(doc)
    assert [e["id"] for e in doc.outline()] == ["b4", "b2", "b3", "b1"]
    reopen_with_python_docx(doc, tmp_path)

    doc = moved()
    doc.reject_revisions()
    assert view(doc) == original and not markup(doc) and "move" not in xml(doc)
    reopen_with_python_docx(doc, tmp_path)

    # Picking one half of a move takes the other half with it.
    doc = moved()
    one = next(r["id"] for r in doc.revisions() if r["type"] == "move" and r["text"] == "one")
    assert doc.reject_revisions(ids=[one]) == 4
    assert [r["type"] for r in doc.revisions() if r["type"] == "move"] == []
    assert doc.text().split("\n")[-3:] == ["one", "two", "three"] and "move" not in xml(doc)


def test_moving_your_own_insertion_and_the_move_tool():
    doc = tracked(THREE)
    new = doc.insert_paragraph("fresh")
    doc.move_block(new, "start")
    assert order(doc)[0] == (new, "fresh") and "move" not in xml(doc)

    registry = ToolRegistry(default_tools(doc))
    assert registry.execute("move_blocks", {"block_ids": ["b2", "b3"], "after": "start"}) == "Moved 2 block(s) after start."
    assert [text for _, text in order(doc)] == ["two", "three", "fresh", "one"]
    assert registry.execute("move_blocks", {"block_ids": ["b2"], "after": "b2"}).startswith("Error: after='b2' is one of")
    listing = registry.execute("list_revisions", {})
    assert "move by Agent in b2: 'two'" in listing and "move by Agent in its old place: 'two'" in listing
    doc.reject_revisions(author="Agent")
    assert doc.text() == "one\ntwo\nthree"
