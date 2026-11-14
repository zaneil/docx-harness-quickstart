from __future__ import annotations

import asyncio
import json
from typing import Literal, Optional

from docx_harness import ToolRegistry, default_tools, tool


def test_schema_from_signature_and_docstring():
    @tool(read_only=True)
    def lookup(term: str, limit: int = 5, mode: Literal["fast", "deep"] = "fast", tags: Optional[list[str]] = None) -> str:
        """Look something up.

        Second line of the description.

        Args:
            term: What to look for.
            limit: Maximum results,
                continued on a second line.
            tags: Optional tags.

        Returns:
            Text.
        """
        return term

    assert lookup.name == "lookup"
    assert lookup.read_only is True
    assert lookup.description == "Look something up. Second line of the description."
    assert lookup.parameters == {
        "type": "object",
        "properties": {
            "term": {"type": "string", "description": "What to look for."},
            "limit": {"type": "integer", "description": "Maximum results, continued on a second line."},
            "mode": {"type": "string", "enum": ["fast", "deep"]},
            "tags": {"type": "array", "items": {"type": "string"}, "description": "Optional tags."},
        },
        "required": ["term"],
    }
    assert lookup.to_openai()["function"]["name"] == "lookup"


def test_default_tools_have_valid_schemas(doc):
    tools = default_tools(doc)
    names = [t.name for t in tools]
    assert "write_part_xml" not in names
    assert "write_part_xml" in [t.name for t in default_tools(doc, include_raw=True)]
    assert len(names) == len(set(names))
    for t in tools:
        assert t.description and len(t.description) < 400
        assert t.parameters["type"] == "object"
        for prop, schema in t.parameters["properties"].items():
            assert "description" in schema, f"{t.name}.{prop} has no description"
        json.dumps(t.to_openai())

    assert {"accept_revisions", "reject_revisions"} <= {t.name for t in tools if not t.read_only}
    read_only = {t.name for t in tools if t.read_only}
    assert read_only == {
        "get_outline", "read_blocks", "search_text", "list_styles", "list_revisions", "list_parts", "read_part_xml",
        "read_headers_footnotes",
    }


async def test_builtin_tools_end_to_end(doc):
    registry = ToolRegistry(default_tools(doc))
    assert "empty" in await registry.aexecute("get_outline")

    result = await registry.aexecute("insert_paragraphs", json.dumps({
        "paragraphs": [
            {"text": "Report", "style": "Title"},
            {"text": "Summary", "style": "Heading 1"},
            {"text": "Revenue grew 10 percent."},
        ]
    }))
    assert result == "Inserted 3 paragraph(s): b1, b2, b3"

    await registry.aexecute("insert_table", {"rows": [["Q", "Rev"], ["Q1", "10"]], "after": "b2"})
    outline = await registry.aexecute("get_outline")
    assert outline.splitlines() == [
        "b1 [Title] Report",
        "b2 [heading 1] Summary",
        "b4 [table] 2x2 table: Q | Rev",
        "b3 [Normal] Revenue grew 10 percent.",
    ]

    assert await registry.aexecute("replace_text", {"find": "10 percent", "replace": "12 percent"}) == "Replaced 1 occurrence(s)."
    assert "Formatted 1 run" in await registry.aexecute("format_text", {"blocks": ["b3"], "match": "12 percent", "bold": True})
    assert "1 paragraph" in await registry.aexecute("set_paragraph_format", {"blocks": ["b1"], "alignment": "center"})
    await registry.aexecute("set_table_cell", {"table_id": "b4", "row": 1, "col": 1, "text": "11"})
    await registry.aexecute("delete_blocks", {"block_ids": ["b2"]})

    blocks = json.loads(await registry.aexecute("read_blocks"))["blocks"]
    assert [b["id"] for b in blocks] == ["b1", "b4", "b3"]
    assert blocks[1]["rows"][1] == ["Q1", "11"]
    assert blocks[2]["text"] == "Revenue grew 12 percent."
    assert "word/document.xml" in await registry.aexecute("list_parts")
    assert "Title" in await registry.aexecute("list_styles")


async def test_errors_are_returned_not_raised(doc):
    registry = ToolRegistry(default_tools(doc))
    assert (await registry.aexecute("nope")).startswith("Error: unknown tool 'nope'")
    assert "not a valid JSON object" in await registry.aexecute("get_outline", "{bad json")
    assert "No block with id 'b99'" in await registry.aexecute("update_paragraph", {"block_id": "b99", "text": "x"})
    assert (await registry.aexecute("update_paragraph", {"wrong": 1})).startswith("Error:")
    result = await registry.aexecute("insert_paragraphs", {"paragraphs": [{"text": "x", "style": "Nope"}]})
    assert "Available:" in result


async def test_plan_mode_blocks_mutating_tools(doc):
    registry = ToolRegistry(default_tools(doc))
    result = await registry.aexecute("insert_paragraphs", {"paragraphs": [{"text": "x"}]}, read_only=True)
    assert "not available in plan mode" in result
    assert len(doc) == 0
    read_only_names = {t["function"]["name"] for t in registry.to_openai_schema(read_only=True)}
    assert "insert_paragraphs" not in read_only_names and "get_outline" in read_only_names


async def test_custom_sync_async_and_blocking_tools():
    registry = ToolRegistry()

    @registry.register
    def add(a: int, b: int) -> int:
        """Add two numbers."""
        return a + b

    async def fetch(url: str) -> dict:
        """Pretend to fetch."""
        await asyncio.sleep(0)
        return {"url": url}

    registry.register(fetch)
    registry.register(tool(blocking=True)(lambda: "slow"))

    assert await registry.aexecute("add", '{"a": 2, "b": 3}') == "5"
    assert json.loads(await registry.aexecute("fetch", {"url": "x"})) == {"url": "x"}
    assert await registry.aexecute("<lambda>") == "slow"
    assert registry.names == ["add", "fetch", "<lambda>"]


def test_sync_execute():
    registry = ToolRegistry([tool(lambda: None, name="noop")])
    assert registry.execute("noop") == "OK"


async def test_insert_paragraphs_tidies_markdown_lists(doc):
    registry = ToolRegistry(default_tools(doc))
    await registry.aexecute("insert_paragraphs", {"paragraphs": [
        {"text": "- one\n- two"},                          # list crammed into one paragraph
        {"text": "- three", "style": "List Bullet"},       # marker typed in front of list-styled text
        {"text": "1. first\n2) second"},
        {"text": "Para A\n\nPara B"},
        {"text": "- 5 degrees is cold"},                    # a lone dash line is left alone
        {"text": "line one\nline two"},                    # single newline stays a line break
    ]})
    blocks = [(b["style"], b["text"]) for b in doc.read(count=50)["blocks"]]
    assert blocks == [
        ("List Bullet", "one"), ("List Bullet", "two"), ("List Bullet", "three"),
        ("List Number", "first"), ("List Number", "second"),
        ("Normal", "Para A"), ("Normal", "Para B"),
        ("Normal", "- 5 degrees is cold"),
        ("Normal", "line one\nline two"),
    ]


def test_normalize_keeps_markers_when_document_has_no_list_style():
    from docx_harness.tools.builtin import normalize_paragraphs

    assert normalize_paragraphs([{"text": "- a\n- b"}], available_styles={"normal"}) == [{"text": "- a\n- b", "style": None}]


async def test_teaching_messages(doc):
    registry = ToolRegistry(default_tools(doc))
    await registry.aexecute("insert_paragraphs", {"paragraphs": [{"text": "## Scope"}]})
    assert doc.read()["blocks"][0] == {"id": "b1", "type": "paragraph", "style": "heading 2", "text": "Scope"}

    result = await registry.aexecute("set_paragraph_style", {"blocks": ["b1"], "style": "Heading 2"})
    assert "Nothing changed" in result and "format_text" in result

    result = await registry.aexecute("insert_paragraphs", {"paragraphs": [{"text": "| a | b |\n|---|---|\n| 1 | 2 |"}]})
    assert result.startswith("Error:") and "insert_table" in result
    assert len(doc) == 1


# ------------------------------------------------ content tools (section 6)

def _png(width: int = 96, height: int = 48) -> bytes:
    import struct
    import zlib

    def chunk(kind: bytes, payload: bytes) -> bytes:
        return struct.pack(">I", len(payload)) + kind + payload + struct.pack(">I", zlib.crc32(kind + payload))

    rows = b"".join(b"\x00" + b"\x7f" * (3 * width) for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def test_new_tool_schemas(doc):
    tools = {t.name: t for t in default_tools(doc)}
    sides = tools["format_table_cells"].parameters["properties"]["border_sides"]
    assert sides["type"] == "array" and "outer" in sides["items"]["enum"]
    assert tools["insert_paragraphs"].parameters["properties"]["paragraphs"]["items"]["properties"]["level"]["type"] == "integer"
    assert tools["set_header_footer"].parameters["required"] == ["kind", "text"]
    for name in ("insert_image", "set_image", "set_header_footer", "add_hyperlink", "remove_hyperlink",
                 "create_style", "set_list_format"):
        assert not tools[name].read_only


def test_insert_paragraphs_nests_lists_and_makes_links(doc, tmp_path):
    registry = ToolRegistry(default_tools(doc))
    result = registry.execute("insert_paragraphs", {"paragraphs": [
        {"text": "Plan", "style": "Heading 4"},
        {"text": "See [the guide](https://example.com/guide) and [mail us](mailto:a@example.com)."},
        {"text": "1. First\n2. Second\n   1. Second, part a\n3. Third"},
        {"text": "Loose bullet", "style": "List Bullet", "level": 1},
    ]})
    assert result.startswith("Inserted 7 paragraph(s)")
    blocks = doc.read()["blocks"]
    assert blocks[0]["style"] == "heading 4"
    assert blocks[1]["text"] == "See the guide and mail us."
    assert blocks[1]["links"] == [{"text": "the guide", "url": "https://example.com/guide"},
                                  {"text": "mail us", "url": "mailto:a@example.com"}]
    assert [b["list"]["level"] for b in blocks[2:]] == [0, 0, 1, 0, 1]
    assert [b["text"] for b in blocks[2:6]] == ["First", "Second", "Second, part a", "Third"]
    outline = registry.execute("get_outline", {})
    assert "[List Number, level 1] Second, part a" in outline and "[List Number] Third" in outline

    ids = [b["id"] for b in blocks]
    assert registry.execute("set_list_format", {"blocks": [ids[5]], "level": 1, "start_at": 4}) == (
        f"Set 1 item(s) to level 1, restarted numbering at 4 from {ids[5]} (1 item(s) renumbered).")
    assert "Give level, start_at or both" in registry.execute("set_list_format", {"blocks": [ids[5]]})
    assert "not a list item" in registry.execute("set_list_format", {"blocks": [ids[0]], "level": 1})
    assert registry.execute("remove_hyperlink", {"blocks": [ids[1]], "match": "mail"}) == "Removed 1 link(s)."
    assert registry.execute("add_hyperlink", {"blocks": [ids[0]], "url": "https://example.com"}) == (
        "Linked 1 piece(s) of text to https://example.com.")
    assert registry.execute("add_hyperlink", {"blocks": [ids[0]], "url": "ftp://x"}).startswith("Error: Invalid link address")
    doc.save(tmp_path / "out.docx")


def test_style_tools(doc):
    registry = ToolRegistry(default_tools(doc))
    listing = registry.execute("list_styles", {})
    assert "Also usable (added to the document when first used): caption, footer, header, heading 4" in listing
    assert "create_style" in listing
    assert registry.execute("create_style", {"name": "Callout", "based_on": "Quote", "bold": True, "color": "red"}) == (
        "Created paragraph style 'Callout'.")
    block = doc.insert_paragraph("text")
    assert registry.execute("set_paragraph_style", {"blocks": [block], "style": "Callout"}).startswith("Applied")
    assert registry.execute("set_paragraph_style", {"blocks": [block], "style": "Caption"}).startswith("Applied")
    error = registry.execute("set_paragraph_style", {"blocks": [block], "style": "Nope"})
    assert error.startswith("Error:") and "create_style" in error
    assert "already exists" in registry.execute("create_style", {"name": "callout"})
    assert "Character styles" not in listing and "Callout" in registry.execute("list_styles", {})


def test_image_tool_and_image_root(doc, tmp_path):
    inside = tmp_path / "assets"
    inside.mkdir()
    (inside / "logo.png").write_bytes(_png())
    (tmp_path / "outside.png").write_bytes(_png())

    registry = ToolRegistry(default_tools(doc, image_root=inside))
    result = registry.execute("insert_image", {"path": "logo.png", "alt_text": "Logo", "alignment": "center"})
    assert result == "Inserted image b1 (1.0 x 0.5 in)."
    assert registry.execute("insert_image", {"path": str(inside / "logo.png"), "width_in": 2}).endswith("(2.0 x 1.0 in).")
    for escape in ("../outside.png", str(tmp_path / "outside.png")):
        assert registry.execute("insert_image", {"path": escape}).startswith("Error: Images can only be read from")
    assert registry.execute("insert_image", {"path": "missing.png"}).startswith("Error: Cannot read the image file")
    assert registry.execute("set_image", {"block_id": "b1", "height_in": 1}) == "Updated the image in b1 (2.0 x 1.0 in)."
    assert "[Normal] [image: Logo]" in registry.execute("get_outline", {})
    assert len(doc) == 2

    anywhere = ToolRegistry(default_tools(doc))
    assert anywhere.execute("insert_image", {"path": str(tmp_path / "outside.png")}).startswith("Inserted image")


def test_header_footer_and_table_format_tools(doc):
    registry = ToolRegistry(default_tools(doc))
    result = registry.execute("set_header_footer", {"kind": "footer", "text": "Page {PAGE} of {NUMPAGES}", "alignment": "center"})
    assert result.startswith("Set the footer (word/footer1.xml).")
    registry.execute("set_header_footer", {"kind": "header", "text": "", "page": "first"})
    listing = registry.execute("read_headers_footnotes", {})
    assert "word/footer1.xml (footer):\n  Page {PAGE} of {NUMPAGES}" in listing
    assert "word/header1.xml (header, first page only):\n  (empty)" in listing
    assert registry.execute("format_text", {"blocks": ["word/footer1.xml"], "size_pt": 9}).startswith("Formatted")

    registry.execute("insert_table", {"rows": [["a", "b"], ["c", "d"]]})
    assert registry.execute("format_table_cells", {
        "cells": ["b1[0,*]"], "border": "thick", "border_sides": ["bottom"], "margins_in": 0.08,
    }) == "Formatted 2 cell(s)."
    assert registry.execute("set_table_format", {"table_id": "b1", "borders": "horizontal", "cell_margins_in": 0.05}) == (
        "Set borders horizontal, cell padding 0.05 in on b1.")
    layout = json.loads(registry.execute("read_blocks", {"include_formatting": True}))["blocks"][0]["format"]
    assert layout["borders"] == "horizontal" and layout["cells"][0]["borders"]["bottom"]["style"] == "thick"
    assert registry.execute("format_table_cells", {"cells": ["b1"], "border": "wavy"}).startswith("Error: Invalid border")
