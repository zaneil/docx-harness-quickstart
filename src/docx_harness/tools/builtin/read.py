"""The "read" group: tools that inspect the document and change nothing."""

from __future__ import annotations

import json
from typing import Optional

from ...document import WordDocument
from ..base import Tool, tool
from .common import CREATABLE_STYLES, outline_line


def read_tools(doc: WordDocument) -> list[Tool]:
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
        lines = [outline_line(e) for e in entries]
        total = len(doc)
        if len(entries) < total:
            lines.append(f"({len(entries)} of {total} blocks shown; pass start to continue)")
        return "\n".join(lines)

    @tool(read_only=True)
    def read_blocks(start: Optional[str] = None, count: int = 20, include_formatting: bool = False) -> str:
        """Read the full content of consecutive blocks. Paragraphs give text and style, plus 'list', 'links', 'images', 'notes' (footnotes and endnotes) and 'section_break' when present. Tables give rows on the column grid, plus 'merged' and 'nested' (ids of tables inside cells) when present.

        Args:
            start: Id of the first block, or of a nested table to read just that table. Omit to start at the top.
            count: Number of blocks to read.
            include_formatting: Also return paragraph layout and runs (bold, italic, color, size), and table widths, header rows, borders and cell shading.
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
        """List the styles available in this document, and the standard ones added automatically the first time they are used."""
        styles = doc.list_styles()
        paragraph = sorted(s["name"] for s in styles if s["type"] == "paragraph")
        character = sorted(s["name"] for s in styles if s["type"] == "character" and not s["default"])
        table = sorted(s["name"] for s in styles if s["type"] == "table")
        have = {name.lower() for name in paragraph}
        extra = [name for name in CREATABLE_STYLES if name not in have]
        lines = [f"Paragraph styles: {', '.join(paragraph) or '(none)'}"]
        if character:
            lines.append(f"Character styles: {', '.join(character)}")
        lines.append(f"Table styles: {', '.join(table) or '(none)'}")
        if extra:
            lines.append(f"Also usable (added to the document when first used): {', '.join(extra)}")
        lines.append("For any other look, define a style with create_style.")
        return "\n".join(lines)

    @tool(read_only=True)
    def read_headers_footnotes() -> str:
        """Read the text of headers, footers, footnotes and endnotes, grouped by part name. Page numbers show as {PAGE} and {NUMPAGES}; each note starts with its [id]. A part name works as the block in replace_text and in blocks for format_text and set_paragraph_format; set_header_footer rewrites a header or footer."""
        parts = []
        for story in doc.stories():
            if "notes" in story:
                lines = [f"  [{note['id']}] {note['text']}" for note in story["notes"]]
            else:
                lines = [f"  {text}" for text in story["paragraphs"] if text.strip()]
            pages = {"first": ", first page only", "even": ", even pages"}.get(story.get("page"), "")
            if story.get("sections"):
                numbers = story["sections"]
                pages += f", section{'s' * (len(numbers) > 1)} {', '.join(map(str, numbers))}"
            parts.append(f"{story['part']} ({story['kind']}{pages}):\n" + ("\n".join(lines) or "  (empty)"))
        return "\n".join(parts) or (
            "The document has no headers, footers, footnotes or endnotes. Create a header or footer with "
            "set_header_footer, a note with add_footnote."
        )

    @tool(read_only=True, available=lambda: doc.section_count() > 1)
    def list_sections() -> str:
        """List the document's sections: the blocks each covers, how it starts, and which header and footer parts it shows ('from section N' when it shows an earlier section's). Pass a section number to set_header_footer to change only that section."""
        lines = []
        for s in doc.sections():
            span = f"{s['start']}..{s['end']}" if s["start"] else "(empty)"
            first = ", different first page" if s["first_page_different"] else ""
            lines.append(f"Section {s['section']}: blocks {span}, starts {s['start_type'].replace('_', ' ')}{first}")
            for kind in ("header", "footer"):
                shown = [
                    f"{page} {info['part']}" + (f" (from section {info['from_section']})" if "from_section" in info else "")
                    for page, info in s[kind].items()
                ]
                lines.append(f"  {kind}: {'; '.join(shown) or 'none'}")
        return "\n".join(lines)

    return [get_outline, read_blocks, search_text, list_styles, read_headers_footnotes, list_sections]
