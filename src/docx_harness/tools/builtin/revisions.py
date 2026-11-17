"""The "revisions" group: tracked changes (redlines). Offered only while the
document tracks changes or has some pending."""

from __future__ import annotations

from typing import Optional

from ...document import WordDocument
from ..base import Tool, tool
from .common import as_list


def revision_tools(doc: WordDocument) -> list[Tool]:
    @tool(read_only=True)
    def list_revisions(limit: int = 100) -> str:
        """List the tracked changes pending in the document, each with its revision id in brackets, author and block.

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

    def resolve(accept: bool, ids, author, block, all_revisions) -> str:
        ids = as_list(ids)
        if not (ids or author or block or all_revisions):
            raise ValueError(
                "Say which revisions: pass ids (the numbers in brackets from list_revisions), "
                "author, block, or all_revisions=true. Nothing was changed."
            )
        method = doc.accept_revisions if accept else doc.reject_revisions
        count = method(ids=[str(i) for i in ids] or None, author=author or None, block=block or None)
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
        """Accept tracked changes, making them permanent. Only when the user asks. Give at least one argument; several narrow the selection together.

        Args:
            ids: Revision ids from list_revisions.
            author: Only this author's revisions.
            block: Only revisions inside this block id.
            all_revisions: Every tracked change in the document.
        """
        return resolve(True, ids, author, block, all_revisions)

    @tool
    def reject_revisions(
        ids: Optional[list[str]] = None,
        author: Optional[str] = None,
        block: Optional[str] = None,
        all_revisions: bool = False,
    ) -> str:
        """Reject tracked changes, restoring what was there before them. Only when the user asks. Give at least one argument; several narrow the selection together.

        Args:
            ids: Revision ids from list_revisions.
            author: Only this author's revisions.
            block: Only revisions inside this block id.
            all_revisions: Every tracked change in the document.
        """
        return resolve(False, ids, author, block, all_revisions)

    return [list_revisions, accept_revisions, reject_revisions]
