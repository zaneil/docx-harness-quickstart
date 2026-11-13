"""The default Word tools: thin wrappers binding a WordDocument to Tool objects.

Several tools take lists (paragraphs, block ids) so the model can do more per
call and finish in fewer round trips.
"""

from __future__ import annotations

import json
import re
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
                    "text": {"type": "string", "description": "Paragraph text. One paragraph (or one list item) per entry, without '-' or '1.' markers."},
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


_BULLET = re.compile(r"^\s*[-*•–]\s+")
_HEADING = re.compile(r"^\s*(#{1,3})\s+(.*)$")
_NUMBER = re.compile(r"^\s*\d+[.)]\s+")


def normalize_paragraphs(items: list, available_styles: set[str]) -> list[dict]:
    """Tidy what models commonly send: markdown-style lists crammed into one
    paragraph, list markers typed in front of list-styled text, and blank-line
    separated paragraphs. ``available_styles`` are lower-cased style names."""
    out: list[dict] = []
    for item in items:
        if isinstance(item, str):
            item = {"text": item}
        text = str(item.get("text", ""))
        style = item.get("style") or None
        lines = [line for line in text.split("\n") if line.strip()]
        if len(lines) > 1 and all(line.strip().startswith("|") for line in lines):
            raise ValueError(
                "Paragraph text contains a markdown table. Nothing was inserted. "
                "Use insert_table with rows as a list of lists instead, and resend the other paragraphs."
            )
        heading = _HEADING.match(text) if len(lines) == 1 else None
        if heading and f"heading {len(heading.group(1))}" in available_styles:
            out.append({"text": heading.group(2).strip(), "style": style or f"Heading {len(heading.group(1))}"})
            continue
        is_list_style = bool(style) and "list" in style.lower()
        handled = False
        for marker, auto_style in ((_BULLET, "List Bullet"), (_NUMBER, "List Number")):
            if lines and all(marker.match(line) for line in lines) and (len(lines) > 1 or is_list_style):
                target = style if is_list_style else auto_style
                if is_list_style or auto_style.lower() in available_styles:
                    out.extend({"text": marker.sub("", line).strip(), "style": target} for line in lines)
                    handled = True
                break
        if handled:
            continue
        if "\n\n" in text:
            out.extend({"text": part.strip(), "style": style} for part in re.split(r"\n\s*\n", text) if part.strip())
        else:
            out.append({"text": text, "style": style})
    return out


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
        """Read the full content of consecutive blocks. Paragraphs return text and style. Tables return rows on the column grid, plus 'merged' (cells spanning several positions) and 'nested' (ids of tables inside cells) when present.

        Args:
            start: Id of the first block to read, or the id of a nested table to read just that table. Omit to start at the top.
            count: Number of blocks to read.
            include_formatting: Also return paragraph alignment, spacing, indents and runs (bold, italic, color, size), and for tables column widths, header rows and cell shading.
        """
        return json.dumps(doc.read(start, count, include_formatting), ensure_ascii=False)

    @tool(read_only=True)
    def search_text(query: str, case_sensitive: bool = False) -> str:
        """Find blocks and table cells containing a piece of text. Headers, footers and notes are searched too; those hits name the part instead of a block.

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

    @tool(read_only=True)
    def read_headers_footnotes() -> str:
        """Read the text of the document's headers, footers, footnotes and endnotes, grouped by part name. Edit that text with replace_text, passing the part name as block."""
        sections = []
        for story in doc.stories():
            lines = [f"  {text}" for text in story["paragraphs"] if text.strip()]
            sections.append(f"{story['part']} ({story['kind']}):\n" + ("\n".join(lines) or "  (empty)"))
        return "\n".join(sections) or "The document has no headers, footers, footnotes or endnotes."

    @tool(read_only=True)
    def list_revisions(limit: int = 100) -> str:
        """List the tracked changes (redlines) pending in the document: insertions, deletions and formatting changes, each with its revision id in brackets, the author and the affected block.

        Args:
            limit: Maximum number of revisions to list.
        """
        revisions = doc.revisions()
        if not revisions:
            return "There are no tracked changes in the document."
        lines = []
        for r in revisions[:max(limit, 0)]:
            where = r["block"] or r.get("part") or ("its old place" if r.get("move") == "from" else "deleted block")
            what = repr(r["text"]) if r["text"] else f"whole {r.get('target', 'paragraph')}"
            if r.get("target") == "paragraph format":
                what = "paragraph formatting"
            lines.append(f"[{r['id']}] {r['type']} by {r['author']} in {where}: {what}")
        if len(revisions) > len(lines):
            lines.append(f"({len(lines)} of {len(revisions)} revisions shown)")
        return "\n".join(lines)

    # ----------------------------------------------------------------- write

    @tool(parameters=_PARAGRAPHS_SCHEMA)
    def insert_paragraphs(paragraphs: list, after: str = "end") -> str:
        """Insert one or more paragraphs (headings, body text, list items) at a position. Use a style for headings and lists rather than manual formatting."""
        if not isinstance(paragraphs, list) or not paragraphs:
            raise ValueError("paragraphs must be a non-empty list of {text, style} objects")
        available = {s["name"].lower() for s in doc.list_styles("paragraph")}
        ids = []
        anchor = after
        for item in normalize_paragraphs(paragraphs, available):
            anchor = doc.insert_paragraph(item["text"], item["style"], anchor)
            ids.append(anchor)
        return f"Inserted {len(ids)} paragraph(s): {', '.join(ids)}"

    @tool
    def insert_table(
        rows: list[list[str]],
        after: str = "end",
        header: bool = True,
        style: str = "Table Grid",
        column_widths_in: Optional[list[float]] = None,
    ) -> str:
        """Insert a table.

        Args:
            rows: Table content as a list of rows; each row is a list of cell texts. A newline inside a cell text starts a new paragraph.
            after: Where to insert: a block id (insert after it), 'start' or 'end'.
            header: Make the first row a header: bold, and repeated at the top of each page.
            style: Table style name.
            column_widths_in: Width of each column in inches. Omit to share the page width equally.
        """
        block_id = doc.insert_table(rows, after, style, header, column_widths_in)
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
            block: Limit to one block id, a table cell like 'b3[0,1]', or a part name from read_headers_footnotes. Omit for the whole document, headers, footers and notes included.
            case_sensitive: Match case exactly.
        """
        count = doc.replace_text(find, replace, block, case_sensitive)
        return f"Replaced {count} occurrence(s)." if count else f"No occurrences of {find!r} found; nothing changed."

    @tool
    def set_table_cell(table_id: str, row: int, col: int, text: str) -> str:
        """Set the text of one table cell. Each line of the text is one paragraph, matched to the cell's existing paragraphs so they keep their formatting. For a small wording change prefer replace_text with block='b3[row,col]'.

        Args:
            table_id: Id of the table block.
            row: Zero-based row index.
            col: Zero-based column index on the table's grid. Any position inside a merged cell edits that merged cell.
            text: The new cell text.
        """
        doc.set_cell_text(table_id, row, col, text)
        return f"Updated {table_id}[{row},{col}]"

    @tool
    def insert_table_row(table_id: str, cells: list[str], index: Optional[int] = None) -> str:
        """Add a row to an existing table. The row copies the layout of the row next to it.

        Args:
            table_id: Id of the table block.
            cells: Cell texts for the new row, left to right.
            index: Zero-based position for the new row. Omit to append at the bottom.
        """
        return f"Inserted row {doc.insert_table_row(table_id, cells, index)} in {table_id}"

    @tool
    def insert_table_column(
        table_id: str,
        cells: Optional[list[str]] = None,
        index: Optional[int] = None,
        width_in: Optional[float] = None,
    ) -> str:
        """Add a column to an existing table. The new cells copy the look of the column beside them.

        Args:
            table_id: Id of the table block.
            cells: Cell texts for the new column, top to bottom, one per row starting at row 0; use '' for rows to leave empty, such as a title row above the headings. Omit for an empty column.
            index: Zero-based position for the new column. Omit to add it on the right.
            width_in: Width of the new column in inches. Omit to match its neighbour and keep the table's overall width.
        """
        position = doc.insert_table_column(table_id, index, cells, width_in)
        return f"Inserted column {position} in {table_id}; columns to its right moved up by one."

    @tool
    def merge_table_cells(table_id: str, row: int, col: int, end_row: int, end_col: int) -> str:
        """Merge a rectangle of table cells into one cell, e.g. a heading across several columns. Text from all the cells is kept. The merged cell is addressed by its top-left position.

        Args:
            table_id: Id of the table block.
            row: Zero-based row of the top-left cell.
            col: Zero-based column of the top-left cell.
            end_row: Zero-based row of the bottom-right cell.
            end_col: Zero-based column of the bottom-right cell.
        """
        doc.merge_cells(table_id, row, col, end_row, end_col)
        top, left = min(row, end_row), min(col, end_col)
        return f"Merged into {table_id}[{top},{left}] ({abs(end_row - row) + 1} row(s) by {abs(end_col - col) + 1} column(s))."

    @tool
    def unmerge_table_cells(table_id: str, row: int, col: int) -> str:
        """Split a merged table cell back into single cells. Its text stays in the top-left one.

        Args:
            table_id: Id of the table block.
            row: Zero-based row of any position inside the merged cell.
            col: Zero-based column of any position inside the merged cell.
        """
        doc.unmerge_cells(table_id, row, col)
        return f"Split the merged cell at {table_id}[{row},{col}] into single cells."

    @tool
    def move_blocks(block_ids: list[str], after: str) -> str:
        """Move blocks (paragraphs or whole tables) to another place in the document, keeping their order and their ids. Use this to reorder content instead of deleting and re-inserting it.

        Args:
            block_ids: Ids of the blocks to move, in the order they should appear.
            after: Where to put them: a block id (they go after it), 'start' or 'end'.
        """
        if isinstance(block_ids, str):
            block_ids = [block_ids]
        if after in block_ids:
            raise ValueError(f"after={after!r} is one of the blocks being moved; pick a block that stays put, or 'start' / 'end'.")
        anchor = after
        for block_id in block_ids:
            doc.move_block(block_id, anchor)
            anchor = block_id
        return f"Moved {len(block_ids)} block(s) after {after}."

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
    def delete_table_column(table_id: str, col: int) -> str:
        """Delete one column of a table.

        Args:
            table_id: Id of the table block.
            col: Zero-based column index on the table's grid.
        """
        doc.delete_table_column(table_id, col)
        return f"Deleted column {col} of {table_id}; columns to its right moved down by one."

    @tool
    def delete_blocks(block_ids: list[str]) -> str:
        """Delete blocks (paragraphs or whole tables) from the document.

        Args:
            block_ids: Ids of the blocks to delete. The id of a nested table deletes that table from its cell.
        """
        if isinstance(block_ids, str):
            block_ids = [block_ids]
        for block_id in block_ids:
            doc.delete_block(block_id)
        return f"Deleted {len(block_ids)} block(s)."

    # --------------------------------------------------------------- styling

    @tool
    def set_paragraph_style(blocks: list[str], style: str) -> str:
        """Change which named paragraph style blocks use (e.g. turn body text into 'Heading 2' or 'List Bullet'). To change color, size or font, use format_text instead.

        Args:
            blocks: Block ids, table cells like 'b3[0,1]', table rows like 'b3[0,*]' or table columns like 'b3[*,1]'.
            style: Style name. Use list_styles to see what is available.
        """
        if isinstance(blocks, str):
            blocks = [blocks]
        count = sum(doc.set_paragraph_style(ref, style) for ref in blocks)
        if count == 0:
            return (
                f"Nothing changed: these blocks already use style {style!r}. "
                "To change color, size, font or bold, call format_text instead."
            )
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
            blocks: Block ids, table cells like 'b3[0,1]', table rows like 'b3[0,*]' or table columns like 'b3[*,1]'. A table id formats the whole table.
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
            blocks: Block ids, table cells like 'b3[0,1]', table rows like 'b3[0,*]' or table columns like 'b3[*,1]'.
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

    @tool
    def format_table_cells(
        cells: list[str],
        shading: Optional[str] = None,
        vertical_align: Optional[Literal["top", "center", "bottom"]] = None,
    ) -> str:
        """Set the background color and vertical alignment of table cells. For the text in the cells (bold, color, horizontal alignment) use format_text and set_paragraph_format with the same references.

        Args:
            cells: Cells like 'b3[0,1]', rows like 'b3[0,*]', columns like 'b3[*,1]', or a table id for every cell.
            shading: Background color as 6-digit hex (e.g. 'D9E2F3') or a basic color name; 'none' removes it.
            vertical_align: Where the text sits in the cell.
        """
        if isinstance(cells, str):
            cells = [cells]
        count = sum(doc.format_cells(ref, shading, vertical_align) for ref in cells)
        return f"Formatted {count} cell(s)."

    @tool
    def set_table_format(
        table_id: str,
        column_widths_in: Optional[list[float]] = None,
        header_rows: Optional[int] = None,
        alignment: Optional[Literal["left", "center", "right"]] = None,
    ) -> str:
        """Set a table's layout: column widths, repeating header rows and position on the page. Give at least one.

        Args:
            table_id: Id of the table block.
            column_widths_in: Width of every column in inches, left to right, one per column.
            header_rows: How many rows at the top repeat on each page the table runs onto; 0 for none.
            alignment: Where the table sits between the margins.
        """
        doc.set_table_format(table_id, column_widths_in, header_rows, alignment)
        done = []
        if column_widths_in is not None:
            done.append(f"{len(column_widths_in)} column width(s)")
        if header_rows is not None:
            done.append(f"{header_rows} header row(s)")
        if alignment is not None:
            done.append(f"alignment {alignment}")
        return f"Set {', '.join(done)} on {table_id}."

    # ------------------------------------------------------------- revisions

    def resolve_revisions(accept: bool, ids, author, block, all_revisions) -> str:
        if isinstance(ids, (str, int)):
            ids = [ids]
        if not (ids or author or block or all_revisions):
            raise ValueError(
                "Say which revisions: pass ids (the numbers in brackets from list_revisions), "
                "author, block, or all_revisions=true. Nothing was changed."
            )
        resolve = doc.accept_revisions if accept else doc.reject_revisions
        count = resolve(ids=[str(i) for i in ids] if ids else None, author=author or None, block=block or None)
        if count == 0:
            return "No tracked changes matched; nothing changed."
        left = len(doc.revisions())
        verb = "Accepted" if accept else "Rejected"
        return f"{verb} {count} revision(s); {left} still pending. Block ids may have changed: call get_outline before further edits."

    @tool
    def accept_revisions(
        ids: Optional[list[str]] = None,
        author: Optional[str] = None,
        block: Optional[str] = None,
        all_revisions: bool = False,
    ) -> str:
        """Accept tracked changes, making them a permanent part of the document. Only when the user asks for it. Give at least one of the arguments; several narrow the selection together.

        Args:
            ids: Revision ids, the numbers in brackets from list_revisions.
            author: Accept only this author's revisions.
            block: Accept only revisions inside this block id.
            all_revisions: Accept every tracked change in the document.
        """
        return resolve_revisions(True, ids, author, block, all_revisions)

    @tool
    def reject_revisions(
        ids: Optional[list[str]] = None,
        author: Optional[str] = None,
        block: Optional[str] = None,
        all_revisions: bool = False,
    ) -> str:
        """Reject tracked changes, restoring the text and formatting from before them. Only when the user asks for it. Give at least one of the arguments; several narrow the selection together.

        Args:
            ids: Revision ids, the numbers in brackets from list_revisions.
            author: Reject only this author's revisions.
            block: Reject only revisions inside this block id.
            all_revisions: Reject every tracked change in the document.
        """
        return resolve_revisions(False, ids, author, block, all_revisions)

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
        get_outline, read_blocks, search_text, list_styles, read_headers_footnotes, list_revisions,
        insert_paragraphs, insert_table, insert_page_break,
        update_paragraph, replace_text, set_table_cell, insert_table_row, insert_table_column,
        merge_table_cells, unmerge_table_cells, move_blocks,
        delete_table_row, delete_table_column, delete_blocks,
        set_paragraph_style, format_text, set_paragraph_format, format_table_cells, set_table_format,
        accept_revisions, reject_revisions,
        list_parts, read_part_xml,
    ]
    if include_raw:
        tools.append(write_part_xml)
    return tools
