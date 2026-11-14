"""Tables on their grid: merged cells, nested tables, columns, widths, cell formatting."""

from __future__ import annotations

import json

import pytest

from docx_harness import ToolRegistry, WordDocument, WordPackage, create_agent, default_tools
from docx_harness.document import DocumentError, tables
from docx_harness.document.ooxml import TBLPR_ORDER, TCPR_ORDER, qn

from conftest import ScriptedLLM, assistant, call, doc_from_body, make_docx, reopen_with_python_docx


def cell(text: str, props: str = "", width: int = 3000) -> str:
    paragraphs = "".join(f"<w:p><w:r><w:t>{line}</w:t></w:r></w:p>" for line in text.split("\n"))
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{props}</w:tcPr>{paragraphs}</w:tc>'


GRID = '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
TBLPR = '<w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>'
# A | A | B        A spans two columns
# C | D | E        C runs down two rows
# C | F | G
MERGED = (
    f"<w:tbl>{TBLPR}{GRID}"
    + "<w:tr>" + cell("A", '<w:gridSpan w:val="2"/>', 6000) + cell("B") + "</w:tr>"
    + "<w:tr>" + cell("C", '<w:vMerge w:val="restart"/>') + cell("D") + cell("E") + "</w:tr>"
    + "<w:tr>" + cell("", "<w:vMerge/>") + cell("F") + cell("G") + "</w:tr>"
    + "</w:tbl>"
)
PLAIN = (
    f"<w:tbl>{TBLPR}{GRID}"
    + "".join(f"<w:tr>{cell(f'r{r}a')}{cell(f'r{r}b')}{cell(f'r{r}c')}</w:tr>" for r in range(3))
    + "</w:tbl>"
)


def rows(doc: WordDocument, table: str = "b1") -> list[list[str]]:
    return doc.read(table, 1)["blocks"][0]["rows"]


def tracked(body_xml: str) -> WordDocument:
    doc = WordDocument(WordPackage.from_bytes(make_docx(body_xml)), track_changes=True, author="Agent")
    doc.edits.date = "2026-01-02T03:04:05Z"
    return doc


def children(element, tag: str) -> list[str]:
    """Local names of the children of an element's property child."""
    return [c.tag.split("}")[1] for c in element.find(qn(tag))]


def in_order(names: list[str], order: list[str]) -> bool:
    positions = [order.index(f"w:{n}") for n in names]
    return positions == sorted(positions)


# ------------------------------------------------------------ grid addressing

def test_merged_cells_are_read_and_addressed_on_the_grid(tmp_path):
    doc = doc_from_body(MERGED)
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["A", "", "B"], ["C", "D", "E"], ["", "F", "G"]]
    assert block["merged"] == [{"cell": "b1[0,0]", "rows": 1, "cols": 2}, {"cell": "b1[1,0]", "rows": 2, "cols": 1}]
    assert doc.outline()[0]["preview"].startswith("3x3 table: A |")

    # Before: [0,1] was B and [2,1] was G, one column off.
    doc.set_cell_text("b1", 0, 2, "B2")
    doc.set_cell_text("b1", 2, 1, "F2")
    doc.set_cell_text("b1", 0, 1, "A2")  # a covered position edits the merged cell
    doc.set_cell_text("b1", 2, 0, "C2")
    assert rows(doc) == [["A2", "", "B2"], ["C2", "D", "E"], ["", "F2", "G"]]
    assert doc.search("C2") == [{"ref": "b1[1,0]", "matches": 1, "snippet": "C2"}]
    assert doc.replace_text("G", "G2", ref="b1[2,2]") == 1
    assert doc.format_text("b1[*,0]", bold=True) == 2  # A2 and C2, each once
    assert doc.format_text("b1[1,*]", italic=True) == 3
    with pytest.raises(DocumentError, match="Column 3 is out of range"):
        doc.set_cell_text("b1", 1, 3, "x")
    with pytest.raises(DocumentError, match="Row 5 is out of range"):
        doc.format_text("b1[5,*]", bold=True)

    table = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert table.cell(0, 0)._tc is table.cell(0, 1)._tc and table.cell(1, 0)._tc is table.cell(2, 0)._tc
    assert [table.cell(0, 0).text, table.cell(0, 2).text, table.cell(2, 0).text, table.cell(2, 1).text] == ["A2", "B2", "C2", "F2"]


def test_rows_in_content_controls_and_indented_rows():
    doc = doc_from_body(
        f"<w:tbl>{TBLPR}{GRID}"
        f"<w:tr>{cell('a')}{cell('b')}{cell('c')}</w:tr>"
        f"<w:sdt><w:sdtContent><w:tr>{cell('d')}<w:sdt><w:sdtContent>{cell('e')}</w:sdtContent></w:sdt>{cell('f')}</w:tr></w:sdtContent></w:sdt>"
        f'<w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>{cell("g")}{cell("h")}</w:tr>'
        "</w:tbl>"
    )
    assert rows(doc) == [["a", "b", "c"], ["d", "e", "f"], ["", "g", "h"]]
    doc.set_cell_text("b1", 1, 1, "E")
    doc.set_cell_text("b1", 2, 1, "G")
    assert rows(doc) == [["a", "b", "c"], ["d", "E", "f"], ["", "G", "h"]]
    with pytest.raises(DocumentError, match="out of range"):
        doc.set_cell_text("b1", 2, 0, "x")  # the row starts one column in
    assert doc.insert_table_row("b1", ["x", "y", "z"], index=2) == 2
    assert rows(doc)[2] == ["x", "y", "z"]


# ------------------------------------------------------------------ cell text

def test_cell_text_keeps_paragraphs_and_nested_tables():
    one_column = '<w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>'
    inner = f"<w:tbl>{TBLPR}{one_column}<w:tr>{cell('inner')}</w:tr></w:tbl>"
    doc = doc_from_body(
        f"<w:tbl>{TBLPR}{one_column}<w:tr><w:tc><w:tcPr/>"
        '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Title</w:t></w:r></w:p>'
        f"{inner}"
        '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Note</w:t></w:r></w:p>'
        "</w:tc></w:tr></w:tbl>"
    )
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["Title\nNote"]]  # the nested table's text is not mixed in
    assert block["nested"] == [{"cell": "b1[0,0]", "tables": ["b2"]}]

    def alignments():
        return [p.find(qn("w:pPr")).find(qn("w:jc")).get(qn("w:val")) for p in tables.cell_paragraphs(doc._cell(doc._get("b1"), 0, 0))]

    doc.set_cell_text("b1", 0, 0, "Heading\nRemark")  # line for line: both paragraphs keep their formatting
    tc = doc._cell(doc._get("b1"), 0, 0)
    assert rows(doc) == [["Heading\nRemark"]] and alignments() == ["center", "right"]
    assert tc.find(qn("w:p")).find(qn("w:r")).find(qn("w:rPr")).find(qn("w:b")) is not None

    doc.set_cell_text("b1", 0, 0, "Heading\nRemark\nMore")  # an extra line copies the last paragraph
    assert rows(doc) == [["Heading\nRemark\nMore"]] and alignments() == ["center", "right", "right"]

    doc.set_cell_text("b1", 0, 0, "Only")  # fewer lines: the cell still ends with a paragraph after the table
    assert rows(doc) == [["Only\n"]] and tc[-1].tag == qn("w:p")
    assert rows(doc, "b2") == [["inner"]]

    # The nested table is a table like any other, under its own id.
    doc.set_cell_text("b2", 0, 0, "changed")
    doc.insert_table_row("b2", ["second"])
    assert rows(doc, "b2") == [["changed"], ["second"]]
    assert doc.search("second") == [{"ref": "b2[1,0]", "matches": 1, "snippet": "second"}]
    assert [e["id"] for e in doc.outline()] == ["b1"]
    for use in (lambda: doc.move_block("b2", "start"), lambda: doc.insert_paragraph("x", after="b2")):
        with pytest.raises(DocumentError, match="nested inside a cell"):
            use()
    doc.delete_block("b2")
    assert "nested" not in doc.read("b1", 1)["blocks"][0]
    with pytest.raises(DocumentError, match="No block with id 'b2'"):
        doc.set_cell_text("b2", 0, 0, "x")


def test_new_tables_take_multi_paragraph_cells_and_widths(doc, tmp_path):
    table = doc.insert_table([["Item", "Notes"], ["Pen", "blue\nfine tip"]], column_widths_in=[1, 3])
    block = doc.read(table, 1, include_formatting=True)["blocks"][0]
    assert block["rows"][1] == ["Pen", "blue\nfine tip"]
    assert block["format"] == {"column_widths_in": [1.0, 3.0], "header_rows": 1}
    with pytest.raises(DocumentError, match="one width in inches per column"):
        doc.insert_table([["a", "b"]], column_widths_in=[1])
    reopened = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert [p.text for p in reopened.cell(1, 1).paragraphs] == ["blue", "fine tip"]
    assert in_order(children(doc._get(table), "w:tblPr"), TBLPR_ORDER)


# ----------------------------------------------------------------------- rows

def test_rows_inside_and_at_the_top_of_a_vertical_merge(tmp_path):
    doc = doc_from_body(MERGED)
    assert doc.insert_table_row("b1", ["new1", "new2"], index=2) == 2  # between the two rows C covers
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["A", "", "B"], ["C", "D", "E"], ["", "new1", "new2"], ["", "F", "G"]]
    assert {"cell": "b1[1,0]", "rows": 3, "cols": 1} in block["merged"]
    with pytest.raises(DocumentError, match="the new row has 2 cells"):
        doc.insert_table_row("b1", ["1", "2", "3"], index=2)

    doc.insert_table_row("b1", ["x", "y", "z"])  # below the merge: three cells of its own
    assert rows(doc)[-1] == ["x", "y", "z"]

    doc.delete_table_row("b1", 1)  # the row where C starts: C moves down
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["A", "", "B"], ["C", "new1", "new2"], ["", "F", "G"], ["x", "y", "z"]]
    assert {"cell": "b1[1,0]", "rows": 2, "cols": 1} in block["merged"]
    doc.delete_table_row("b1", 2)
    assert doc.read("b1", 1)["blocks"][0]["merged"] == [{"cell": "b1[0,0]", "rows": 1, "cols": 2}]
    table = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert [c.text for c in table.rows[1].cells] == ["C", "new1", "new2"]


def test_header_rows_stay_at_the_top(doc):
    table = doc.insert_table([["H1", "H2"], ["a", "b"]])
    doc.insert_table_row(table, ["c", "d"])
    doc.insert_table_row(table, ["top1", "top2"], index=0)
    header = [tables.is_header_row(tr) for tr in doc._live_rows(doc._get(table))]
    assert header == [True, True, False, False]


# -------------------------------------------------------------------- columns

def test_insert_and_delete_columns(tmp_path):
    doc = doc_from_body(PLAIN)
    assert doc.insert_table_column("b1", cells=["x0", "x1", "x2"]) == 3
    assert doc.insert_table_column("b1", index=0, cells=["n0", "n1"]) == 0
    assert rows(doc) == [["n0", "r0a", "r0b", "r0c", "x0"], ["n1", "r1a", "r1b", "r1c", "x1"], ["", "r2a", "r2b", "r2c", "x2"]]
    tbl = doc._get("b1")
    widths = tables.grid_widths(tbl)
    assert len(widths) == 5 and sum(widths) == 9000  # the table keeps its width
    assert all(tables.cell_width(tc) == (widths[i], "dxa") for tr in tables.live_rows(tbl) for i, tc in enumerate(tables.row_cells(tr)))

    doc.delete_table_column("b1", 1)
    assert rows(doc)[0] == ["n0", "r0b", "r0c", "x0"]
    assert len(tables.grid_widths(tbl)) == 4 and sum(tables.grid_widths(tbl)) == 9000
    doc.insert_table_column("b1", index=2, width_in=1)  # an explicit width is taken as given
    assert tables.grid_widths(tbl)[2] == 1440 and rows(doc)[1] == ["n1", "r1b", "", "r1c", "x1"]
    with pytest.raises(DocumentError, match="table has 3 rows"):
        doc.insert_table_column("b1", cells=["1", "2", "3", "4"])
    with pytest.raises(DocumentError, match="Column 9 is out of range"):
        doc.delete_table_column("b1", 9)
    table = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert len(table.columns) == 5 and [c.text for c in table.rows[0].cells] == ["n0", "r0b", "", "r0c", "x0"]


def test_columns_through_merged_cells(tmp_path):
    doc = doc_from_body(MERGED)
    doc.insert_table_column("b1", index=1, cells=["ignored", "n1", "n2"])  # inside A's span: A grows
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["A", "", "", "B"], ["C", "n1", "D", "E"], ["", "n2", "F", "G"]]
    assert block["merged"][0] == {"cell": "b1[0,0]", "rows": 1, "cols": 3}

    doc.delete_table_column("b1", 0)  # takes C with it and narrows A
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["A", "", "B"], ["n1", "D", "E"], ["n2", "F", "G"]]
    assert block["merged"] == [{"cell": "b1[0,0]", "rows": 1, "cols": 2}]
    doc.delete_table_column("b1", 0)
    assert doc.read("b1", 1)["blocks"][0] == {"id": "b1", "type": "table", "rows": [["A", "B"], ["D", "E"], ["F", "G"]]}
    doc.delete_table_column("b1", 0)
    with pytest.raises(DocumentError, match="only column"):
        doc.delete_table_column("b1", 0)
    table = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert len(table.columns) == 1 and [r.cells[0].text for r in table.rows] == ["B", "E", "G"]


# -------------------------------------------------------------------- merging

def test_merge_and_unmerge(tmp_path):
    doc = doc_from_body(PLAIN)
    doc.merge_cells("b1", 0, 0, 0, 2)  # a heading across the table
    doc.merge_cells("b1", 2, 2, 1, 1)  # corners in any order: a 2x2 block
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["r0a\nr0b\nr0c", "", ""], ["r1a", "r1b\nr1c\nr2b\nr2c", ""], ["r2a", "", ""]]
    assert block["merged"] == [{"cell": "b1[0,0]", "rows": 1, "cols": 3}, {"cell": "b1[1,1]", "rows": 2, "cols": 2}]
    tbl = doc._get("b1")
    assert tables.cell_width(doc._cell(tbl, 0, 0)) == (9000, "dxa") and tables.cell_width(doc._cell(tbl, 2, 2)) == (6000, "dxa")
    assert all(in_order(children(tc, "w:tcPr"), TCPR_ORDER) for tc in tbl.iter(qn("w:tc")))

    with pytest.raises(DocumentError, match=r"cuts through the merged cell at \[1,1\]"):
        doc.merge_cells("b1", 1, 0, 1, 1)
    with pytest.raises(DocumentError, match="one cell already"):
        doc.merge_cells("b1", 1, 1, 2, 2)
    with pytest.raises(DocumentError, match="not merged"):
        doc.unmerge_cells("b1", 1, 0)

    table = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert table.cell(0, 0)._tc is table.cell(0, 2)._tc and table.cell(1, 1)._tc is table.cell(2, 2)._tc
    assert table.cell(2, 0).text == "r2a"

    doc.unmerge_cells("b1", 2, 2)
    doc.unmerge_cells("b1", 0, 1)
    block = doc.read("b1", 1)["blocks"][0]
    assert "merged" not in block
    assert block["rows"] == [["r0a\nr0b\nr0c", "", ""], ["r1a", "r1b\nr1c\nr2b\nr2c", ""], ["r2a", "", ""]]
    assert all(tables.cell_width(tc) == (3000, "dxa") for tc in tbl.iter(qn("w:tc")))
    assert len(reopen_with_python_docx(doc, tmp_path).tables[0].rows[2].cells) == 3


def test_merging_an_empty_corner_takes_the_other_text():
    doc = doc_from_body(f"<w:tbl>{TBLPR}{GRID}<w:tr>{cell('')}{cell('kept')}{cell('c')}</w:tr></w:tbl>")
    doc.merge_cells("b1", 0, 0, 0, 1)
    assert rows(doc) == [["kept", "", "c"]]


# ----------------------------------------------------------------- formatting

def test_cell_and_table_formatting(tmp_path):
    doc = doc_from_body(MERGED)
    assert doc.format_cells("b1[0,*]", shading="D9E2F3", vertical_align="middle") == 2
    assert doc.format_cells("b1[*,0]", shading="yellow") == 2  # A and C
    assert doc.format_cells("b1", vertical_align="bottom") == 7
    assert doc.format_cells("b1[0,2]", shading="none") == 1
    doc.set_table_format("b1", column_widths_in=[1, 2, 2.5], header_rows=1, alignment="center")
    block = doc.read("b1", 1, include_formatting=True)["blocks"][0]
    assert block["format"] == {
        "column_widths_in": [1.0, 2.0, 2.5], "header_rows": 1, "alignment": "center",
        "cells": [
            {"cell": "b1[0,0]", "shading": "FFFF00", "vertical_align": "bottom"},
            {"cell": "b1[0,2]", "vertical_align": "bottom"},
            {"cell": "b1[1,0]", "shading": "FFFF00", "vertical_align": "bottom"},
            *({"cell": f"b1[{r},{c}]", "vertical_align": "bottom"} for r, c in ((1, 1), (1, 2), (2, 1), (2, 2))),
        ],
    }
    tbl = doc._get("b1")
    assert tables.cell_width(doc._cell(tbl, 0, 0)) == (4320, "dxa")  # the merged cell: 1in + 2in
    assert tables.table_width(tbl) == (7920, "dxa")
    assert in_order(children(tbl, "w:tblPr"), TBLPR_ORDER)
    assert all(in_order(children(tc, "w:tcPr"), TCPR_ORDER) for tc in tbl.iter(qn("w:tc")))
    # The continuation cell of a vertical merge is shaded with it.
    assert tables.cell_format(tables.build_grid(tbl).at(2, 0).own)["shading"] == "FFFF00"

    doc.set_table_format("b1", header_rows=0)
    assert "header_rows" not in doc.read("b1", 1, include_formatting=True)["blocks"][0]["format"]
    for bad, message in (
        (dict(column_widths_in=[1, 2]), "table has 3 columns"),
        (dict(header_rows=9), "between 0 and 3"),
        (dict(alignment="justify"), "Invalid alignment"),
        ({}, "at least one table property"),
    ):
        with pytest.raises(DocumentError, match=message):
            doc.set_table_format("b1", **bad)
    with pytest.raises(DocumentError, match="Invalid vertical_align"):
        doc.format_cells("b1", vertical_align="sideways")

    reopened = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert reopened.alignment is not None and reopened.columns[2].width.inches == 2.5


def borders(doc: WordDocument, row: int, col: int, table: str = "b1") -> dict:
    """side -> style of the cell's own borders."""
    found = tables.cell_format(doc._cell(doc._get(table), row, col)).get("borders", {})
    return {side: edge["style"] for side, edge in found.items()}


def test_cell_borders_padding_and_text_direction(tmp_path):
    doc = doc_from_body(PLAIN)
    # A box around the middle row: its outline only, shared edges set on both sides.
    assert doc.format_cells("b1[1,*]", border="double", border_sides="outer", border_color="C00000",
                            border_width_pt=1.5) == 3
    assert borders(doc, 1, 0) == {"top": "double", "left": "double", "bottom": "double"}
    assert borders(doc, 1, 1) == {"top": "double", "bottom": "double"}
    assert borders(doc, 1, 2) == {"top": "double", "bottom": "double", "right": "double"}
    assert borders(doc, 0, 1) == {"bottom": "double"} and borders(doc, 2, 1) == {"top": "double"}
    edge = tables.cell_format(doc._cell(doc._get("b1"), 1, 0))["borders"]["top"]
    assert edge == {"style": "double", "width_pt": 1.5, "color": "C00000"}

    # The lines between the cells of the first column, then one removed.
    doc.format_cells("b1[*,0]", border_sides=["inner"])
    assert borders(doc, 0, 0) == {"bottom": "single"} and borders(doc, 2, 0) == {"top": "single"}
    doc.format_cells("b1[2,2]", border="none", border_sides="top")
    assert borders(doc, 2, 2) == {"top": "none"} and borders(doc, 1, 2)["bottom"] == "none"

    assert doc.format_cells("b1[0,*]", margins_in=0.1, text_direction="vertical") == 3
    look = tables.cell_format(doc._cell(doc._get("b1"), 0, 2))
    assert look["margins_in"] == 0.1 and look["text_direction"] == "vertical"
    doc.format_cells("b1[0,2]", text_direction="horizontal")
    assert "text_direction" not in tables.cell_format(doc._cell(doc._get("b1"), 0, 2))

    doc.set_table_format("b1", borders="horizontal", cell_margins_in=0.05)
    layout = doc.read("b1", 1, include_formatting=True)["blocks"][0]["format"]
    assert layout["borders"] == "horizontal" and layout["cell_margins_in"] == 0.05
    corner = next(c for c in layout["cells"] if c["cell"] == "b1[0,2]")
    assert corner["margins_in"] == 0.1 and "text_direction" not in corner and set(corner["borders"]) == {"bottom"}

    tbl = doc._get("b1")
    assert in_order(children(tbl, "w:tblPr"), TBLPR_ORDER)
    assert all(in_order(children(tc, "w:tcPr"), TCPR_ORDER) for tc in tbl.iter(qn("w:tc")))
    for tc in tbl.iter(qn("w:tc")):
        edges = tc.find(qn("w:tcPr")).find(qn("w:tcBorders"))
        if edges is not None:
            names = [e.tag.split("}")[1] for e in edges]
            assert names == [n for n in ("top", "left", "bottom", "right") if n in names]

    for bad, message in (
        (dict(border="wavy"), "Invalid border "),
        (dict(border_sides=["diagonal"]), "Invalid border_sides"),
        (dict(border_width_pt=20), "between 0.25 and 6"),
        (dict(margins_in=3), "between 0 and 1"),
        (dict(text_direction="upside down"), "Invalid text_direction"),
        ({}, "at least one cell property"),
    ):
        with pytest.raises(DocumentError, match=message):
            doc.format_cells("b1", **bad)
    with pytest.raises(DocumentError, match="Invalid borders"):
        doc.set_table_format("b1", borders="zigzag")
    assert len(reopen_with_python_docx(doc, tmp_path).tables[0].rows) == 3


def test_border_of_a_merged_cell_runs_along_all_its_rows():
    doc = doc_from_body(MERGED)
    doc.format_cells("b1[1,0]", border="thick", border_sides=["left", "bottom", "top"])
    grid = tables.build_grid(doc._get("b1"))
    top, bottom = grid.at(1, 0).own, grid.at(2, 0).own
    assert set(tables.cell_format(top)["borders"]) == {"top", "left"}
    assert set(tables.cell_format(bottom)["borders"]) == {"left", "bottom"}
    # A, above it, is wider than C, so its bottom edge is left alone.
    assert borders(doc, 0, 0) == {}


def test_tracked_borders_are_format_revisions_and_reject():
    doc = tracked(PLAIN)
    doc.format_cells("b1[0,0]", border="single", margins_in=0.1)
    doc.set_table_format("b1", borders="none")
    assert {r["type"] for r in doc.revisions()} == {"format"}
    assert len(doc.revisions()) == 4  # the cell, its two neighbours, the table
    doc.reject_revisions()
    assert "tcBorders" not in doc.read_part_xml("word/document.xml")
    layout = doc.read("b1", 1, include_formatting=True)["blocks"][0]["format"]
    assert "cells" not in layout and "borders" not in layout


LEGACY = (
    f"<w:tbl>{TBLPR}{GRID}"
    + "<w:tr>" + cell("Wide", '<w:hMerge w:val="restart"/>') + cell("", "<w:hMerge/>") + cell("B") + "</w:tr>"
    + "<w:tr>" + cell("C") + cell("D") + cell("E") + "</w:tr>"
    + "</w:tbl>"
)


def test_legacy_horizontal_merge_reads_as_one_cell(tmp_path):
    doc = doc_from_body(LEGACY)
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["Wide", "", "B"], ["C", "D", "E"]]
    assert block["merged"] == [{"cell": "b1[0,0]", "rows": 1, "cols": 2}]
    doc.set_cell_text("b1", 0, 1, "Wider")  # any position inside it
    assert rows(doc)[0] == ["Wider", "", "B"]
    assert "hMerge" in doc.read_part_xml("word/document.xml")  # reading and text edits leave the markup alone

    doc.insert_table_column("b1", 1, ["x", "y"])
    xml = doc.read_part_xml("word/document.xml")
    assert "hMerge" not in xml and xml.count("<w:gridSpan") == 1
    assert rows(doc) == [["Wider", "", "", "B"], ["C", "y", "D", "E"]]
    tbl = doc._get("b1")
    assert tables.span(doc._cell(tbl, 0, 0)) == 3
    reopened = reopen_with_python_docx(doc, tmp_path).tables[0]
    assert [c.text for c in reopened.rows[1].cells] == ["C", "y", "D", "E"]


def test_legacy_merge_can_be_split():
    doc = doc_from_body(LEGACY)
    doc.unmerge_cells("b1", 0, 0)
    assert "merged" not in doc.read("b1", 1)["blocks"][0]
    assert rows(doc)[0] == ["Wide", "", "B"]
    assert "hMerge" not in doc.read_part_xml("word/document.xml")
    assert tables.cell_width(doc._cell(doc._get("b1"), 0, 0)) == (3000, "dxa")


# ------------------------------------------------------------ tracked changes

def test_tracked_columns_are_cell_revisions_that_accept_and_reject(tmp_path):
    def edited() -> WordDocument:
        doc = tracked(PLAIN)
        doc.insert_table_column("b1", index=1, cells=["n0", "n1", "n2"])
        doc.delete_table_column("b1", 3)  # r*c, read as it will be once accepted
        return doc

    doc = edited()
    before = [["r0a", "r0b", "r0c"], ["r1a", "r1b", "r1c"], ["r2a", "r2b", "r2c"]]
    after = [["r0a", "n0", "r0b"], ["r1a", "n1", "r1b"], ["r2a", "n2", "r2b"]]
    assert rows(doc) == after
    tbl = doc._get("b1")
    assert len(tbl.findall(f".//{qn('w:tc')}")) == 12  # nothing has left the file
    assert len(tbl.findall(f".//{qn('w:cellIns')}")) == 3 and len(tbl.findall(f".//{qn('w:cellDel')}")) == 3
    assert all(in_order(children(tc, "w:tcPr"), TCPR_ORDER) for tc in tbl.iter(qn("w:tc")))
    cell_revisions = [r for r in doc.revisions() if r.get("target") == "table cell"]
    assert [(r["type"], r["author"], r["block"]) for r in cell_revisions] == [
        (kind, "Agent", "b1") for _ in range(3) for kind in ("insert", "delete")
    ]
    assert "r0c" not in doc.text() and doc.search("r1c") == []
    doc.set_cell_text("b1", 1, 2, "R1B")  # addressing skips the deleted cells
    assert rows(doc)[1] == ["r1a", "n1", "R1B"]
    assert len(reopen_with_python_docx(doc, tmp_path).tables[0].rows) == 3

    doc = edited()
    doc.accept_revisions()
    assert rows(doc) == after and not doc.revisions()
    assert len(tables.grid_cols(doc._get("b1"))) == 3 and len(list(doc._get("b1").iter(qn("w:tc")))) == 9

    doc = edited()
    doc.reject_revisions()
    assert rows(doc) == before and not doc.revisions()
    assert len(tables.grid_cols(doc._get("b1"))) == 3 and len(list(doc._get("b1").iter(qn("w:tc")))) == 9
    assert [c.text for c in reopen_with_python_docx(doc, tmp_path).tables[0].rows[2].cells] == ["r2a", "r2b", "r2c"]

    # Taking back your own pending column leaves no trace.
    doc = tracked(PLAIN)
    doc.insert_table_column("b1", cells=["x", "y", "z"])
    doc.delete_table_column("b1", 3)
    assert rows(doc) == before and not [r for r in doc.revisions() if r.get("target") == "table cell"]


def test_tracked_table_formatting_and_merge_refusal():
    doc = tracked(MERGED)
    original = doc.read("b1", 1, include_formatting=True)["blocks"][0]
    doc.format_cells("b1[1,1]", shading="FF0000")
    doc.set_table_format("b1", column_widths_in=[1, 1, 1], header_rows=1, alignment="right")
    kinds = {r.get("target") for r in doc.revisions()}
    assert kinds == {"table format"} and all(r["type"] == "format" for r in doc.revisions())
    assert doc.read("b1", 1, include_formatting=True)["blocks"][0]["format"]["alignment"] == "right"
    doc.reject_revisions()
    restored = doc.read("b1", 1, include_formatting=True)["blocks"][0]
    restored["format"].pop("column_widths_in"), original["format"].pop("column_widths_in")  # the grid is not tracked
    assert restored == original and not doc.revisions()
    tbl = doc._get("b1")
    assert tables.cell_width(doc._cell(tbl, 0, 0)) == (6000, "dxa") and tables.table_width(tbl) == (9000, "dxa")

    for attempt in (lambda: doc.merge_cells("b1", 1, 1, 1, 2), lambda: doc.unmerge_cells("b1", 0, 0)):
        with pytest.raises(DocumentError, match="cannot be recorded as a tracked change"):
            attempt()
    assert not doc.revisions()


def test_tracked_row_delete_moves_a_merge_down_and_rejects_cleanly():
    doc = tracked(MERGED)
    doc.delete_table_row("b1", 1)
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["A", "", "B"], ["C", "F", "G"]] and block["merged"] == [{"cell": "b1[0,0]", "rows": 1, "cols": 2}]
    doc.reject_revisions()
    block = doc.read("b1", 1)["blocks"][0]
    assert block["rows"] == [["A", "", "B"], ["C", "D", "E"], ["", "F", "G"]]
    assert {"cell": "b1[1,0]", "rows": 2, "cols": 1} in block["merged"]


# ---------------------------------------------------------------------- tools

async def test_table_tools_end_to_end(doc):
    registry = ToolRegistry(default_tools(doc))
    await registry.aexecute("insert_table", {"rows": [["Budget", "", ""], ["Item", "Q1", "Q2"], ["Travel", "5", "7"]]})
    assert await registry.aexecute("insert_table_column", {"table_id": "b1", "cells": ["", "Total", "12"]}) == (
        "Inserted column 3 in b1; columns to its right moved up by one."
    )
    assert await registry.aexecute("merge_table_cells", {"table_id": "b1", "row": 0, "col": 0, "end_row": 0, "end_col": 3}) == (
        "Merged into b1[0,0] (1 row(s) by 4 column(s))."
    )
    assert await registry.aexecute("format_table_cells", {"cells": ["b1[0,*]", "b1[*,3]"], "shading": "D9E2F3"}) == "Formatted 4 cell(s)."
    assert "Formatted 3 paragraph(s)" in await registry.aexecute("set_paragraph_format", {"blocks": ["b1[*,3]"], "alignment": "right"})
    assert await registry.aexecute("set_table_format", {"table_id": "b1", "column_widths_in": [2, 1, 1, 1], "header_rows": 2}) == (
        "Set 4 column width(s), 2 header row(s) on b1."
    )
    assert await registry.aexecute("delete_table_column", {"table_id": "b1", "col": 2}) == (
        "Deleted column 2 of b1; columns to its right moved down by one."
    )
    block = json.loads(await registry.aexecute("read_blocks", {"include_formatting": True}))["blocks"][0]
    assert block["rows"] == [["Budget", "", ""], ["Item", "Q1", "Total"], ["Travel", "5", "12"]]
    assert block["merged"] == [{"cell": "b1[0,0]", "rows": 1, "cols": 3}]
    assert block["format"]["header_rows"] == 2 and len(block["format"]["column_widths_in"]) == 3
    assert await registry.aexecute("unmerge_table_cells", {"table_id": "b1", "row": 0, "col": 2}) == (
        "Split the merged cell at b1[0,2] into single cells."
    )

    # Errors tell the model what to do next.
    result = await registry.aexecute("merge_table_cells", {"table_id": "b1", "row": 0, "col": 0, "end_row": 9, "end_col": 0})
    assert result.startswith("Error:") and "out of range" in result
    result = await registry.aexecute("set_table_format", {"table_id": "b1", "column_widths_in": [1]})
    assert result.startswith("Error:") and "3 columns" in result


async def test_agent_run_edits_a_merged_table_under_track_changes():
    doc = tracked(MERGED)
    llm = ScriptedLLM([
        assistant(call("read_blocks")),
        assistant(
            call("insert_table_column", table_id="b1", cells=["Total", "1", "2"]),
            call("merge_table_cells", table_id="b1", row=1, col=1, end_row=1, end_col=2),
        ),
        assistant(content="Added a Total column. The merge could not be tracked."),
    ])
    result = await create_agent(doc, llm=llm).arun("Add a Total column and merge D with E")
    tool_results = [m["content"] for m in result.messages if m["role"] == "tool"]
    assert '"merged"' in tool_results[0] and tool_results[1].startswith("Inserted column 3")
    assert tool_results[2].startswith("Error:") and "track changes switched off" in tool_results[2]
    assert rows(doc) == [["A", "", "B", "Total"], ["C", "D", "E", "1"], ["", "F", "G", "2"]]
