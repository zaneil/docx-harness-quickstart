from __future__ import annotations

import asyncio
import json
from typing import Literal, Optional

from word_agent import ToolRegistry, default_tools, tool


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

    read_only = {t.name for t in tools if t.read_only}
    assert read_only == {"get_outline", "read_blocks", "search_text", "list_styles", "list_parts", "read_part_xml"}


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
    from word_agent.tools.builtin import normalize_paragraphs

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
