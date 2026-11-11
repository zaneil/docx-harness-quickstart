"""The default Word tools: thin wrappers binding a WordDocument to Tool objects.

Several tools take lists (paragraphs, block ids) so the model can do more per
call and finish in fewer round trips.
"""

from __future__ import annotations

import json
from typing import Literal, Optional

from ..document import WordDocument
from .base import Tool, tool

_PARAGRAPHS_SCHEMA = {
    "type": "object",
    "properties": {
        "paragraphs": {
            "type": "array",
            "description": "Paragraphs to insert, in order.",
            "items": {
                "type": "object",
                "properties": {
                    "text": {"type": "string", "description": "Paragraph text. One paragraph per item; do not put blank lines inside."},
                    "style": {"type": "string", "description": "Paragraph style name, e.g. 'Title', 'Heading 1', 'List Bullet'. Omit for normal body text."},
                },
                "required": ["text"],
            },
        },
        "after": {
            "type": "string",
            "description": "Where to insert: a block id (insert after it), 'start' or 'end'. Default 'end'.",
        },
    },
    "required": ["paragraphs"],
}


def default_tools(doc: WordDocument, include_raw: bool = False) -> list[Tool]:
    """Build the built-in tools for ``doc``.

    ``include_raw=True`` adds ``write_part_xml``, which lets the model replace
    raw XML parts. It is off by default because a bad write can corrupt the file.
    """

    # ------------------------------------------------------------------ read

    @tool(read_only=True)
    def get_outline(start: Optional[str] = None, limit: int = 200) -> str:
        """List the document's blocks in order, one line each: id, style and a short text preview. Call this first to learn the block ids.

        Args:
            start: Block id to start from. Omit to start at the top.
            limit: Maximum number of blocks to list.
        """
        entries = doc.outline(start, limit)
        if not entries:
            return "The document is empty. Insert content with after='end'."
        lines = []
        for e in entries:
            label = e.get("style") or e["type"]
            lines.append(f"{e['id']} [{label}] {e['preview']}")
        total = len(doc)
        if len(entries) < total:
            lines.append(f"({len(entries)} of {total} blocks shown; pass start to continue)")
        return "\n".join(lines)

    @tool(read_only=True)
    def read_blocks(start: Optional[str] = None, count: int = 20, include_formatting: bool = False) -> str:
        """Read the full content of consecutive blocks. Paragraphs return their text and style; tables return their rows.

        Args:
            start: Id of the first block to read. Omit to start at the top.
            count: Number of blocks to read.
            include_formatting: Also return each paragraph's runs with bold, italic, color and size.
        """
        return json.dumps(doc.read(start, count, include_formatting), ensure_ascii=False)

    @tool(read_only=True)
    def search_text(query: str, case_sensitive: bool = False) -> str:
        """Find blocks and table cells containing a piece of text.

        Args:
            query: Text to look for.
            case_sensitive: Match case exactly.
        """
        hits = doc.search(query, case_sensitive)
        return json.dumps(hits, ensure_ascii=False) if hits else f"No matches for {query!r}."

    @tool(read_only=True)
    def list_styles() -> str:
        """List the paragraph and table styles available in this document."""
        styles = doc.list_styles()
        paragraph = sorted(s["name"] for s in styles if s["type"] == "paragraph")
        table = sorted(s["name"] for s in styles if s["type"] == "table")
        return f"Paragraph styles: {', '.join(paragraph) or '(none)'}\nTable styles: {', '.join(table) or '(none)'}"

    # ----------------------------------------------------------------- write

    @tool(parameters=_PARAGRAPHS_SCHEMA)
    def insert_paragraphs(paragraphs: list, after: str = "end") -> str:
        """Insert one or more paragraphs (headings, body text, list items) at a position. Use a style for headings and lists rather than manual formatting."""
        if not isinstance(paragraphs, list) or not paragraphs:
            raise ValueError("paragraphs must be a non-empty list of {text, style} objects")
        ids = []
        anchor = after
        for item in paragraphs:
            if isinstance(item, str):
                item = {"text": item}
            anchor = doc.insert_paragraph(str(item.get("text", "")), item.get("style"), anchor)
            ids.append(anchor)
        return f"Inserted {len(ids)} paragraph(s): {', '.join(ids)}"

    @tool
    def insert_table(rows: list[list[str]], after: str = "end", header: bool = True, style: str = "Table Grid") -> str:
        """Insert a table.

        Args:
            rows: Table content as a list of rows; each row is a list of cell texts.
            after: Where to insert: a block id (insert after it), 'start' or 'end'.
            header: Make the first row bold.
            style: Table style name.
        """
        block_id = doc.insert_table(rows, after, style, header)
        return f"Inserted table {block_id} ({len(rows)} rows). Cells are addressed as {block_id}[row,col]."

    @tool
    def insert_page_break(after: str = "end") -> str:
        """Insert a page break.

        Args:
            after: Where to insert: a block id (insert after it), 'start' or 'end'.
        """
        return f"Inserted page break {doc.insert_page_break(after)}"

    # ---------------------------------------------------------------- update

    @tool
    def update_paragraph(block_id: str, text: str) -> str:
        """Replace the entire text of one paragraph. Its style is kept.

        Args:
            block_id: Id of the paragraph.
            text: The new text.
        """
        doc.set_paragraph_text(block_id, text)
        return f"Updated {block_id}"

    @tool
    def replace_text(find: str, replace: str, block: Optional[str] = None, case_sensitive: bool = True) -> str:
        """Find and replace text, preserving formatting. Best for small edits inside a paragraph.

        Args:
            find: Exact text to find.
            replace: Replacement text. Use an empty string to delete.
            block: Limit to one block id or table cell like 'b3[0,1]'. Omit for the whole document.
            case_sensitive: Match case exactly.
        """
        count = doc.replace_text(find, replace, block, case_sensitive)
        return f"Replaced {count} occurrence(s)." if count else f"No occurrences of {find!r} found; nothing changed."

    @tool
    def set_table_cell(table_id: str, row: int, col: int, text: str) -> str:
        """Set the text of one table cell.

        Args:
            table_id: Id of the table block.
            row: Zero-based row index.
            col: Zero-based column index.
            text: The new cell text.
        """
        doc.set_cell_text(table_id, row, col, text)
        return f"Updated {table_id}[{row},{col}]"

    @tool
    def insert_table_row(table_id: str, cells: list[str], index: Optional[int] = None) -> str:
        """Add a row to an existing table.

        Args:
            table_id: Id of the table block.
            cells: Cell texts for the new row, left to right.
            index: Zero-based position for the new row. Omit to append at the bottom.
        """
        return f"Inserted row {doc.insert_table_row(table_id, cells, index)} in {table_id}"

    # ---------------------------------------------------------------- delete

    @tool
    def delete_table_row(table_id: str, row: int) -> str:
        """Delete one row of a table.

        Args:
            table_id: Id of the table block.
            row: Zero-based row index.
        """
        doc.delete_table_row(table_id, row)
        return f"Deleted row {row} of {table_id}"

    @tool
    def delete_blocks(block_ids: list[str]) -> str:
        """Delete blocks (paragraphs or whole tables) from the document.

        Args:
            block_ids: Ids of the blocks to delete.
        """
        if isinstance(block_ids, str):
            block_ids = [block_ids]
        for block_id in block_ids:
            doc.delete_block(block_id)
        return f"Deleted {len(block_ids)} block(s)."

    # --------------------------------------------------------------- styling

    @tool
    def set_paragraph_style(blocks: list[str], style: str) -> str:
        """Apply a named paragraph style (e.g. 'Heading 2', 'List Bullet', 'Normal') to blocks.

        Args:
            blocks: Block ids or table cells like 'b3[0,1]'.
            style: Style name. Use list_styles to see what is available.
        """
        if isinstance(blocks, str):
            blocks = [blocks]
        count = sum(doc.set_paragraph_style(ref, style) for ref in blocks)
        return f"Applied style {style!r} to {count} paragraph(s)."

    @tool
    def format_text(
        blocks: list[str],
        match: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        strike: Optional[bool] = None,
        color: Optional[str] = None,
        size_pt: Optional[float] = None,
        font: Optional[str] = None,
        highlight: Optional[str] = None,
    ) -> str:
        """Apply character formatting to whole blocks, or only to occurrences of some text inside them.

        Args:
            blocks: Block ids or table cells like 'b3[0,1]'. A table id formats the whole table.
            match: Only format this exact text within the blocks. Omit to format all text.
            bold: Turn bold on or off.
            italic: Turn italic on or off.
            underline: Turn underline on or off.
            strike: Turn strikethrough on or off.
            color: Text color as 6-digit hex (e.g. '1F4E79') or a basic color name.
            size_pt: Font size in points.
            font: Font family name.
            highlight: Highlight color name such as 'yellow', or 'none' to remove.
        """
        if isinstance(blocks, str):
            blocks = [blocks]
        count = sum(
            doc.format_text(
                ref, match, bold=bold, italic=italic, underline=underline, strike=strike,
                color=color, size_pt=size_pt, font=font, highlight=highlight,
            )
            for ref in blocks
        )
        return f"Formatted {count} run(s) in {len(blocks)} block(s)."

    @tool
    def set_paragraph_format(
        blocks: list[str],
        alignment: Optional[Literal["left", "center", "right", "justify"]] = None,
        space_before_pt: Optional[float] = None,
        space_after_pt: Optional[float] = None,
        line_spacing: Optional[float] = None,
        indent_left_in: Optional[float] = None,
        indent_right_in: Optional[float] = None,
        first_line_indent_in: Optional[float] = None,
        page_break_before: Optional[bool] = None,
        keep_with_next: Optional[bool] = None,
    ) -> str:
        """Set paragraph layout: alignment, spacing and indentation.

        Args:
            blocks: Block ids or table cells like 'b3[0,1]'.
            alignment: Horizontal alignment.
            space_before_pt: Space above the paragraph in points.
            space_after_pt: Space below the paragraph in points.
            line_spacing: Line spacing multiple, e.g. 1.0, 1.5 or 2.0.
            indent_left_in: Left indent in inches.
            indent_right_in: Right indent in inches.
            first_line_indent_in: First-line indent in inches; negative for a hanging indent.
            page_break_before: Start the paragraph on a new page.
            keep_with_next: Keep on the same page as the next paragraph.
        """
        if isinstance(blocks, str):
            blocks = [blocks]
        count = sum(
            doc.set_paragraph_format(
                ref, alignment=alignment, space_before_pt=space_before_pt, space_after_pt=space_after_pt,
                line_spacing=line_spacing, indent_left_in=indent_left_in, indent_right_in=indent_right_in,
                first_line_indent_in=first_line_indent_in, page_break_before=page_break_before,
                keep_with_next=keep_with_next,
            )
            for ref in blocks
        )
        return f"Formatted {count} paragraph(s)."

    # ------------------------------------------------------------- raw parts

    @tool(read_only=True)
    def list_parts() -> str:
        """List the files (parts) inside the unzipped .docx package."""
        return "\n".join(doc.list_parts())

    @tool(read_only=True)
    def read_part_xml(name: str, offset: int = 0, max_chars: int = 6000) -> str:
        """Read the raw XML of a part of the .docx package. Only needed for things the other tools do not cover.

        Args:
            name: Part name, e.g. 'word/document.xml' or 'word/styles.xml'.
            offset: Character offset to start from.
            max_chars: Maximum number of characters to return.
        """
        xml = doc.read_part_xml(name)
        chunk = xml[offset:offset + max_chars]
        remaining = len(xml) - offset - len(chunk)
        return chunk + (f"\n... ({remaining} more characters; continue with offset={offset + len(chunk)})" if remaining > 0 else "")

    @tool
    def write_part_xml(name: str, xml: str) -> str:
        """Replace the entire raw XML of an existing part. Last resort: invalid WordprocessingML will corrupt the document. Rewriting word/document.xml resets all block ids.

        Args:
            name: Part name, e.g. 'word/styles.xml'.
            xml: Complete, well-formed XML for the part.
        """
        doc.write_part_xml(name, xml)
        return f"Wrote {name}. Call get_outline again before further edits."

    tools = [
        get_outline, read_blocks, search_text, list_styles,
        insert_paragraphs, insert_table, insert_page_break,
        update_paragraph, replace_text, set_table_cell, insert_table_row,
        delete_table_row, delete_blocks,
        set_paragraph_style, format_text, set_paragraph_format,
        list_parts, read_part_xml,
    ]
    if include_raw:
        tools.append(write_part_xml)
    return tools
