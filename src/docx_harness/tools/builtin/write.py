"""The "write" group: add, rewrite, move and delete content.

Structural edits (insert, move, delete) answer with the outline around the
change, so the model sees where things landed without reading again.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Literal, Optional

from ...document import WordDocument
from ..base import Tool, tool
from .common import (
    CREATABLE_STYLES,
    around,
    as_list,
    normalize_paragraphs,
    split_links,
    with_context,
)

_PARAGRAPHS_SCHEMA = {
    "type": "object",
    "properties": {
        "paragraphs": {
            "type": "array",
            "description": "Paragraphs to insert, in order.",
            "items": {
                "type": "object",
                "properties": {
                    "text": {"type": "string", "description": "One paragraph or list item, without '-' or '1.' markers. A link can be written as [text](https://address)."},
                    "style": {"type": "string", "description": "Style name, e.g. 'Heading 1', 'List Bullet'. Omit for body text."},
                    "level": {"type": "integer", "description": "List items only: 0 (default) outer list, 1 sub-item, up to 8."},
                },
                "required": ["text"],
            },
        },
        "after": {"type": "string", "description": "A block id to insert after, 'start' or 'end' (default)."},
    },
    "required": ["paragraphs"],
}


def write_tools(doc: WordDocument, image_root: str | os.PathLike | None = None) -> list[Tool]:
    def image_path(path: str) -> Path:
        candidate = Path(str(path)).expanduser()
        if image_root is None:
            return candidate
        root = Path(image_root).resolve()
        resolved = (candidate if candidate.is_absolute() else root / candidate).resolve()
        if root != resolved and root not in resolved.parents:
            raise ValueError(f"Images can only be read from {root}. Give a path inside that folder.")
        return resolved

    # ----------------------------------------------------------------- insert

    @tool(parameters=_PARAGRAPHS_SCHEMA)
    def insert_paragraphs(paragraphs: list, after: str = "end") -> str:
        """Insert paragraphs (headings, body text, list items) at a position. Use styles for headings and lists, not manual formatting. Numbered items placed after a numbered item continue its list; elsewhere they start at 1."""
        if not isinstance(paragraphs, list) or not paragraphs:
            raise ValueError("paragraphs must be a non-empty list of {text, style} objects")
        available = {s["name"].lower() for s in doc.list_styles("paragraph")} | set(CREATABLE_STYLES)
        ids = []
        anchor = after
        for item in normalize_paragraphs(paragraphs, available):
            text, found = split_links(item["text"])
            anchor = doc.insert_paragraph(text, item["style"], anchor, item.get("level"))
            for label, url in dict.fromkeys(found):
                doc.add_hyperlink(anchor, url, match=label)
            ids.append(anchor)
        return with_context(f"Inserted {len(ids)} paragraph(s): {', '.join(ids)}.", doc, ids)

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
            rows: List of rows, each a list of cell texts. A newline in a cell starts a new paragraph.
            after: A block id to insert after, 'start' or 'end'.
            header: Make the first row a bold header that repeats on each page.
            style: Table style name.
            column_widths_in: Width of each column in inches. Omit to share the page width.
        """
        block_id = doc.insert_table(rows, after, style, header, column_widths_in)
        summary = f"Inserted table {block_id} ({len(rows)} rows). Cells are addressed as {block_id}[row,col]."
        return with_context(summary, doc, [block_id])

    @tool
    def insert_page_break(after: str = "end") -> str:
        """Insert a page break.

        Args:
            after: A block id to insert after, 'start' or 'end'.
        """
        block_id = doc.insert_page_break(after)
        return with_context(f"Inserted page break {block_id}.", doc, [block_id])

    @tool
    def insert_section_break(after: str, start: Literal["new_page", "continuous", "even_page", "odd_page"] = "new_page") -> str:
        """Start a new section after a block, so that part of the document can have its own header or footer (set_header_footer with section). Both sections keep the current headers and footers until changed.

        Args:
            after: Id of the last block of the current section; the new section starts after it.
            start: Where the new section begins: on a new page, on the same page (continuous), or on the next even or odd page.
        """
        before = set(doc.block_ids())
        number = doc.insert_section_break(after, start)
        added = [b for b in doc.block_ids() if b not in before]
        return with_context(
            f"Inserted section break {added[0]}; section {number} starts after it. "
            f"To give it its own header, call set_header_footer with section={number}.",
            doc, added,
        )

    @tool
    def insert_image(
        path: str,
        after: str = "end",
        width_in: Optional[float] = None,
        height_in: Optional[float] = None,
        alt_text: str = "",
        alignment: Optional[Literal["left", "center", "right"]] = None,
    ) -> str:
        """Insert a PNG, JPEG or GIF picture in a paragraph of its own. Without a size it keeps its natural size, shrunk to fit the margins. Delete it with delete_blocks.

        Args:
            path: Path of the image file.
            after: A block id to insert after, 'start' or 'end'.
            width_in: Width in inches; the height follows unless also given.
            height_in: Height in inches; the width follows unless also given.
            alt_text: Short description for readers who cannot see it.
            alignment: Position between the margins.
        """
        block_id = doc.insert_image(image_path(path), after, width_in, height_in, alt_text, alignment)
        size = doc.read(block_id, 1)["blocks"][0]["images"][0]
        return with_context(f"Inserted image {block_id} ({size['width_in']} x {size['height_in']} in).", doc, [block_id])

    @tool
    def set_image(block_id: str, width_in: Optional[float] = None, height_in: Optional[float] = None,
                  alt_text: Optional[str] = None) -> str:
        """Resize a picture or change its alt text. Give one dimension to keep the proportions.

        Args:
            block_id: Id of the block holding the picture.
            width_in: New width in inches.
            height_in: New height in inches.
            alt_text: New description of the picture.
        """
        doc.set_image(block_id, width_in, height_in, alt_text)
        size = doc.read(block_id, 1)["blocks"][0]["images"][0]
        return f"Updated the image in {block_id} ({size['width_in']} x {size['height_in']} in)."

    @tool
    def set_header_footer(
        kind: Literal["header", "footer"],
        text: str,
        page: Literal["default", "first", "even"] = "default",
        alignment: Optional[Literal["left", "center", "right"]] = None,
        section: Optional[int] = None,
    ) -> str:
        """Set the text of the page header or footer, creating it if needed. Replaces what was there.

        Args:
            kind: Header (top of the page) or footer (bottom).
            text: Each line is a paragraph. {PAGE} is the page number and {NUMPAGES} the page count. A tab moves to the centre, a second to the right edge. Empty for a blank one.
            page: 'default' for all pages, 'first' for a different first page (e.g. blank on a cover), 'even' for even pages (even pages apply to the whole document).
            alignment: Alignment of the text.
            section: Section number (from 1, see list_sections) to change only that section. Omit to set it for the whole document.
        """
        setter = doc.set_header if kind == "header" else doc.set_footer
        part = setter(text, page, alignment, section)
        where = f" for section {section}" if section is not None else ""
        return f"Set the {kind}{where} ({part}). To restyle it, pass '{part}' in blocks to format_text or set_paragraph_format."

    @tool
    def add_footnote(block_id: str, text: str, after_text: Optional[str] = None,
                     kind: Literal["footnote", "endnote"] = "footnote") -> str:
        """Add a footnote (or endnote) to a paragraph. Its number is placed at the end of the paragraph, or right after after_text. Use this for citations and asides instead of writing them into the text.

        Args:
            block_id: Id of the paragraph (or a table cell like 'b3[0,1]').
            text: The note's text, without a number; Word numbers notes itself.
            after_text: Exact text in the paragraph that the number should follow, e.g. the end of a sentence. Omit for the end of the paragraph.
            kind: 'footnote' (bottom of the page) or 'endnote' (end of the document).
        """
        note_id = doc.add_note(block_id, text, after_text, kind)
        return (f"Added {kind} {note_id} in {block_id}. Change its text with replace_text on the notes part "
                f"(read_headers_footnotes), or remove it with delete_footnote.")

    @tool
    def delete_footnote(note_id: str, kind: Literal["footnote", "endnote"] = "footnote") -> str:
        """Delete a footnote or endnote: its number in the text and the note itself.

        Args:
            note_id: The note's id, as shown in read_blocks ('notes') or read_headers_footnotes.
            kind: 'footnote' or 'endnote'.
        """
        doc.delete_note(note_id, kind)
        return f"Deleted {kind} {note_id}."

    # ----------------------------------------------------------------- update

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
            replace: Replacement text; empty to delete.
            block: Limit to one block id, a table cell like 'b3[0,1]', or a part name from read_headers_footnotes. Omit for the whole document, headers, footers and notes included.
            case_sensitive: Match case exactly.
        """
        count = doc.replace_text(find, replace, block, case_sensitive)
        return f"Replaced {count} occurrence(s)." if count else f"No occurrences of {find!r} found; nothing changed."

    @tool
    def add_hyperlink(blocks: list[str], url: str, match: Optional[str] = None) -> str:
        """Turn existing text into a link. Text that is already a link is pointed at the new address.

        Args:
            blocks: Block ids, table cells like 'b3[0,1]' or part names.
            url: Full address starting with https://, http:// or mailto:.
            match: Only link this exact text. Omit to link all the text of each block.
        """
        count = sum(doc.add_hyperlink(ref, url, match) for ref in as_list(blocks))
        return f"Linked {count} piece(s) of text to {url}."

    @tool
    def remove_hyperlink(blocks: list[str], match: Optional[str] = None) -> str:
        """Turn links back into plain text. The text itself stays.

        Args:
            blocks: Block ids, table cells like 'b3[0,1]' or part names.
            match: Only remove links whose text contains this. Omit to remove every link in the blocks.
        """
        count = sum(doc.remove_hyperlink(ref, match) for ref in as_list(blocks))
        return f"Removed {count} link(s)."

    # ------------------------------------------------------------ move/delete

    @tool
    def move_blocks(block_ids: list[str], after: Optional[str] = None, before: Optional[str] = None) -> str:
        """Move blocks (paragraphs or whole tables) elsewhere, keeping their order and ids. Use this to reorder instead of deleting and re-inserting. Give after or before: to put a section ahead of another, pass before=<that section's heading>.

        Args:
            block_ids: Ids of the blocks to move, in the order they should appear.
            after: Put them after this block id, or at 'start' or 'end'.
            before: Put them just before this block id.
        """
        block_ids = as_list(block_ids)
        if (after is None) == (before is None):
            raise ValueError("Give exactly one of after and before. Nothing was moved.")
        anchor = after if after is not None else before
        if anchor in block_ids:
            where = "after" if after is not None else "before"
            raise ValueError(f"{where}={anchor!r} is one of the blocks being moved; pick a block that stays put. Nothing was moved.")
        first, *rest = block_ids
        if before is not None:
            doc.move_block(first, before=before)
        else:
            doc.move_block(first, after)
        previous = first
        for block_id in rest:
            doc.move_block(block_id, previous)
            previous = block_id
        where = f"after {after}" if after is not None else f"before {before}"
        return with_context(f"Moved {len(block_ids)} block(s) {where}.", doc, block_ids)

    @tool
    def delete_blocks(block_ids: list[str]) -> str:
        """Delete blocks (paragraphs or whole tables).

        Args:
            block_ids: Ids of the blocks to delete. The id of a nested table deletes that table from its cell.
        """
        block_ids = as_list(block_ids)
        order = doc.block_ids()
        gone = set(block_ids)
        positions = [i for i, block_id in enumerate(order) if block_id in gone]
        for block_id in block_ids:
            doc.delete_block(block_id)
        summary = f"Deleted {len(block_ids)} block(s)."
        if not positions:  # nested tables only: nothing changed in the body's order
            return summary
        neighbours = [order[i] for i in (positions[0] - 1, positions[-1] + 1)
                      if 0 <= i < len(order) and order[i] not in gone]
        context = around(doc, neighbours, heading="Now around the gap:", mark=False)
        return f"{summary} {context}" if context else summary

    return [
        insert_paragraphs, insert_table, insert_page_break, insert_section_break, insert_image, set_image,
        set_header_footer, add_footnote, delete_footnote, update_paragraph, replace_text, add_hyperlink, remove_hyperlink, move_blocks, delete_blocks,
    ]
