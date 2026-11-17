"""The "raw" group: the unzipped package's XML parts, for what the semantic
tools do not cover. Writing is off unless ``include_raw=True``."""

from __future__ import annotations

from ...document import WordDocument
from ..base import Tool, tool


def raw_tools(doc: WordDocument, include_raw: bool = False) -> list[Tool]:
    @tool(read_only=True)
    def list_parts() -> str:
        """List the files (parts) inside the unzipped .docx package."""
        return "\n".join(doc.list_parts())

    @tool(read_only=True)
    def read_part_xml(name: str, offset: int = 0, max_chars: int = 6000) -> str:
        """Read the raw XML of a package part. Only for things the other tools do not cover.

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

    return [list_parts, read_part_xml, *([write_part_xml] if include_raw else [])]
