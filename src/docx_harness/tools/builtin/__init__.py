"""The default Word tools: thin wrappers binding a WordDocument to Tool objects.

They come in groups, one module each, so you can take only the ones you need:

| Group       | Module          | Tools                                                        |
| ----------- | --------------- | ------------------------------------------------------------ |
| "read"      | ``read.py``     | outline, read, search, styles, headers and notes             |
| "write"     | ``write.py``    | insert, rewrite, replace, links, images, headers, move, delete |
| "tables"    | ``tables.py``   | edit existing tables (offered only while there is one)       |
| "format"    | ``formatting.py`` | styles, lists, character and paragraph formatting          |
| "revisions" | ``revisions.py`` | tracked changes (offered only while tracking or pending)    |
| "raw"       | ``raw.py``      | package parts and raw XML                                    |

Several tools take lists (paragraphs, block ids) so the model can do more per
call and finish in fewer round trips.
"""

from __future__ import annotations

import os
from typing import Iterable

from ...document import WordDocument
from ..base import Tool
from .common import normalize_paragraphs
from .formatting import format_tools
from .raw import raw_tools
from .read import read_tools
from .revisions import revision_tools
from .tables import table_tools
from .write import write_tools

TOOL_GROUPS = ("read", "write", "tables", "format", "revisions", "raw")


def default_tools(doc: WordDocument, include_raw: bool = False,
                  image_root: str | os.PathLike | None = None,
                  groups: Iterable[str] | None = None) -> list[Tool]:
    """Build the built-in tools for ``doc``.

    ``groups`` picks which groups to include (see :data:`TOOL_GROUPS`); all of
    them by default. The "tables" tools are offered to the model only while the
    document has a table, and the "revisions" tools only while it tracks changes
    or has some pending, so they cost no prompt space otherwise.

    ``include_raw=True`` adds ``write_part_xml``, which lets the model replace
    raw XML parts. It is off by default because a bad write can corrupt the file.

    ``image_root`` limits ``insert_image`` to files inside that folder
    (relative paths are taken from it). Without it any readable image file on
    this machine can be inserted, which suits a command line run by its user.
    """
    chosen = TOOL_GROUPS if groups is None else tuple(groups)
    unknown = sorted(set(chosen) - set(TOOL_GROUPS))
    if unknown:
        raise ValueError(f"Unknown tool group(s) {', '.join(unknown)}. Choose from {', '.join(TOOL_GROUPS)}.")
    builders = {
        "read": lambda: read_tools(doc),
        "write": lambda: write_tools(doc, image_root),
        "tables": lambda: _when(table_tools(doc), doc.has_tables),
        "format": lambda: format_tools(doc),
        "revisions": lambda: _when(revision_tools(doc), lambda: doc.track_changes or doc.has_revisions()),
        "raw": lambda: raw_tools(doc, include_raw),
    }
    return [t for group in TOOL_GROUPS if group in chosen for t in builders[group]()]


def _when(tools: list[Tool], check) -> list[Tool]:
    for t in tools:
        t.available = check
    return tools


__all__ = ["default_tools", "normalize_paragraphs", "TOOL_GROUPS"]
