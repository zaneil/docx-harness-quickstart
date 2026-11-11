from __future__ import annotations

import pytest

from word_agent import DocumentError, StyleError, WordDocument
from word_agent.document.ooxml import qn

from conftest import doc_from_body, reopen_with_python_docx

MIXED = (
    "<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>big wor</w:t></w:r>"
    "<w:r><w:rPr><w:i/></w:rPr><w:t>ld</w:t></w:r><w:r><w:t>!</w:t></w:r></w:p>"
)


def test_insert_read_and_stable_ids(doc, tmp_path):
    title = doc.insert_paragraph("My Title", "Title")
    body = doc.insert_paragraph("Body text")
    heading = doc.insert_paragraph("Intro", "heading 1", after=title)
    first = doc.insert_paragraph("Very first", after="start")

    assert [b["id"] for b in doc.outline()] == [first, title, heading, body]
    doc.delete_block(first)
    # ids of the remaining blocks did not shift
    assert [b["id"] for b in doc.outline()] == [title, heading, body]
    assert doc.read(heading, 1)["blocks"][0] == {"id": heading, "type": "paragraph", "style": "heading 1", "text": "Intro"}
    assert doc.read(title, 2)["next"] == body

    with pytest.raises(DocumentError, match="No block with id"):
        doc.delete_block(first)

    reopened = reopen_with_python_docx(doc, tmp_path)
    assert [(p.text, p.style.name) for p in reopened.paragraphs] == [
        ("My Title", "Title"), ("Intro", "Heading 1"), ("Body text", "Normal"),
    ]


def test_content_stays_before_section_properties(doc):
    doc.insert_paragraph("a")
    doc.insert_table([["x"]])
    assert doc._body[-1].tag == qn("w:sectPr")


def test_unknown_style_lists_available(doc):
    with pytest.raises(StyleError, match="Available:.*Title"):
        doc.insert_paragraph("x", "Fancy Style")


def test_newlines_and_tabs_round_trip(doc):
    block = doc.insert_paragraph("line one\nline two\tend")
    assert doc.read(block, 1)["blocks"][0]["text"] == "line one\nline two\tend"


def test_replace_across_runs_keeps_formatting(tmp_path):
    doc = doc_from_body(MIXED)
    assert doc.text() == "Hello big world!"
    assert doc.replace_text("big world", "small planet") == 1
    assert doc.text() == "Hello small planet!"

    runs = doc.read(include_formatting=True)["blocks"][0]["runs"]
    assert runs == [{"text": "Hello "}, {"text": "small planet", "bold": True}, {"text": "!"}]

    reopened = reopen_with_python_docx(doc, tmp_path)
    assert reopened.paragraphs[0].text == "Hello small planet!"


def test_replace_multiple_case_insensitive_and_scoped():
    doc = doc_from_body("<w:p><w:r><w:t>cat Cat CAT</w:t></w:r></w:p><w:p><w:r><w:t>cat</w:t></w:r></w:p>")
    first, second = [b["id"] for b in doc.outline()]
    assert doc.replace_text("cat", "dog", ref=first, case_sensitive=False) == 3
    assert doc.text() == "dog dog dog\ncat"
    # a replacement containing the search text must not loop forever
    assert doc.replace_text("cat", "catcat", ref=second) == 1
    assert doc.replace_text("zebra", "x") == 0
    assert doc.replace_text("dog ", "") == 2


def test_tracked_deletions_are_not_read_and_are_preserved(tmp_path):
    body = (
        "<w:p><w:r><w:t>Keep </w:t></w:r>"
        '<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>removed </w:delText></w:r></w:del>'
        '<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>added </w:t></w:r></w:ins>'
        "<w:r><w:t>text</w:t></w:r></w:p>"
    )
    doc = doc_from_body(body)
    assert doc.text() == "Keep added text"
    doc.replace_text("text", "words")
    xml = doc.read_part_xml("word/document.xml")
    assert "removed " in xml and "<w:del " in xml and "<w:ins " in xml
    assert doc.text() == "Keep added words"


def test_set_paragraph_text_keeps_style_and_run_format():
    doc = doc_from_body('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>Old</w:t></w:r></w:p>')
    block = doc.outline()[0]["id"]
    doc.set_paragraph_text(block, "New heading")
    info = doc.read(block, 1, include_formatting=True)["blocks"][0]
    assert info["style"] == "heading 1"
    assert info["runs"] == [{"text": "New heading", "italic": True}]


def test_tables(doc, tmp_path):
    table = doc.insert_table([["Name", "Qty"], ["Apples", "3"]])
    assert doc.read(table, 1)["blocks"][0]["rows"] == [["Name", "Qty"], ["Apples", "3"]]

    doc.set_cell_text(table, 1, 1, "5")
    assert doc.insert_table_row(table, ["Pears", "7"]) == 2
    assert doc.insert_table_row(table, ["Kiwis"], index=1) == 1
    assert doc.read(table, 1)["blocks"][0]["rows"] == [["Name", "Qty"], ["Kiwis", ""], ["Apples", "5"], ["Pears", "7"]]

    doc.delete_table_row(table, 1)
    assert doc.search("pears") == [{"ref": f"{table}[2,0]", "matches": 1, "snippet": "Pears"}]
    assert doc.replace_text("Pears", "Plums", ref=f"{table}[2,0]") == 1

    with pytest.raises(DocumentError, match="out of range"):
        doc.set_cell_text(table, 9, 0, "x")
    with pytest.raises(DocumentError, match="not a table"):
        doc.set_cell_text(doc.insert_paragraph("p"), 0, 0, "x")

    reopened = reopen_with_python_docx(doc, tmp_path)
    cells = [[c.text for c in row.cells] for row in reopened.tables[0].rows]
    assert cells == [["Name", "Qty"], ["Apples", "5"], ["Plums", "7"]]
    assert reopened.tables[0].rows[0].cells[0].paragraphs[0].runs[0].bold is True
    assert not reopened.tables[0].rows[2].cells[0].paragraphs[0].runs[0].bold


def test_format_text_whole_and_substring(doc, tmp_path):
    block = doc.insert_paragraph("The quick brown fox")
    assert doc.format_text(block, match="quick", bold=True, color="#ff0000") == 1
    assert doc.format_text(block, size_pt=14, font="Arial") == 3

    reopened = reopen_with_python_docx(doc, tmp_path)
    runs = reopened.paragraphs[0].runs
    assert [r.text for r in runs] == ["The ", "quick", " brown fox"]
    assert [bool(r.bold) for r in runs] == [False, True, False]
    assert str(runs[1].font.color.rgb) == "FF0000"
    assert all(r.font.size.pt == 14 and r.font.name == "Arial" for r in runs)

    doc.format_text(block, match="quick", bold=False)
    assert "bold" not in doc.read(block, 1, include_formatting=True)["blocks"][0]["runs"][1]

    with pytest.raises(DocumentError, match="not found"):
        doc.format_text(block, match="zebra", bold=True)
    with pytest.raises(DocumentError, match="at least one"):
        doc.format_text(block)
    with pytest.raises(StyleError, match="Invalid color"):
        doc.format_text(block, color="reddish")


def test_run_properties_are_written_in_schema_order(doc):
    block = doc.insert_paragraph("x")
    doc.format_text(block, highlight="yellow")
    doc.format_text(block, underline=True)
    doc.format_text(block, size_pt=12)
    doc.format_text(block, bold=True)
    doc.format_text(block, font="Arial")
    rpr = doc._get(block).find(".//" + qn("w:rPr"))
    assert [child.tag for child in rpr] == [qn(t) for t in ("w:rFonts", "w:b", "w:sz", "w:szCs", "w:highlight", "w:u")]


def test_paragraph_format_and_style(doc, tmp_path):
    block = doc.insert_paragraph("Centered")
    doc.set_paragraph_format(block, alignment="center", space_after_pt=12, line_spacing=1.5, first_line_indent_in=-0.25)
    doc.set_paragraph_style(block, "Quote")

    ppr = doc._get(block).find(qn("w:pPr"))
    assert [child.tag for child in ppr] == [qn(t) for t in ("w:pStyle", "w:spacing", "w:ind", "w:jc")]

    from docx.enum.text import WD_ALIGN_PARAGRAPH

    paragraph = reopen_with_python_docx(doc, tmp_path).paragraphs[0]
    assert paragraph.style.name == "Quote"
    assert paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER
    assert paragraph.paragraph_format.space_after.pt == 12
    assert paragraph.paragraph_format.line_spacing == 1.5
    assert round(paragraph.paragraph_format.first_line_indent.inches, 2) == -0.25


def test_cell_reference_formatting(doc):
    table = doc.insert_table([["a", "b"], ["c", "d"]], header=False)
    assert doc.format_text(f"{table}[1,1]", italic=True) == 1
    assert doc.set_paragraph_format(table, alignment="right") == 4
    with pytest.raises(DocumentError, match="Invalid block reference"):
        doc.format_text("table one", bold=True)


def test_raw_part_write_resets_ids(doc):
    doc.insert_paragraph("before")
    xml = doc.read_part_xml("word/document.xml").replace("before", "after")
    doc.write_part_xml("word/document.xml", xml)
    assert doc.text() == "after"
    assert len(doc.outline()) == 1


def test_track_changes_hook_is_reserved():
    with pytest.raises(NotImplementedError):
        WordDocument.open(track_changes=True)


def test_page_break_and_lists(doc, tmp_path):
    doc.insert_paragraph("one", "List Bullet")
    doc.insert_paragraph("two", "List Number")
    doc.insert_page_break()
    reopened = reopen_with_python_docx(doc, tmp_path)
    assert [p.style.name for p in reopened.paragraphs[:2]] == ["List Bullet", "List Number"]
    assert 'w:type="page"' in doc.read_part_xml("word/document.xml")
