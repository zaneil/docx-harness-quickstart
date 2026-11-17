"""Tracked changes: listing, accepting and rejecting revisions.

Recording them happens in ``edits.py`` (``EditContext``).

Mixed into ``WordDocument`` (``document.py``), which provides the block-id
helpers it relies on (``_get``, ``_body_block``, ``_table``, ``_paragraphs_for``,
``_describe``...) and ``self.edits``, the single path for every mutation."""

from __future__ import annotations

from lxml import etree

from . import tables
from .common import (
    DocumentError,
)
from .ooxml import (
    AUTHOR,
    CELL_DEL,
    CELL_INS,
    DEL,
    DEL_TEXT,
    HYPERLINK,
    INS,
    MOVE_FROM,
    MOVE_RANGE_STARTS,
    MOVE_TO,
    PPR_CHANGE,
    PROPERTY_CHANGES,
    RPR,
    RPR_CHANGE,
    TBL,
    TCPR,
    TRPR,
    P,
    T,
    is_deleted_block,
    paragraph_text,
    qn,
    run_text,
)

_REVISION_TYPES = {INS: "insert", DEL: "delete", MOVE_FROM: "move", MOVE_TO: "move",
                   CELL_INS: "insert", CELL_DEL: "delete",
                   **{tag: "format" for tag in PROPERTY_CHANGES}}


class RevisionsMixin:
    # ------------------------------------------------------------- revisions

    def revisions(self) -> list[dict]:
        """Tracked changes pending in the document, in document order: the
        body first, then headers, footers and notes.

        Each entry has ``id``, ``type`` (insert, delete, format or move; a move
        has two halves, told apart by ``move``: "from" or "to"), ``author``,
        ``date``, ``block`` (the block id, or None when the block itself is
        deleted or the revision is outside the body, in which case ``part``
        names where it is) and ``text`` (the inserted, deleted or reformatted
        text; a whole inserted or deleted paragraph or row adds an entry with
        empty text and a ``target``)."""
        out: list[dict] = []
        for part, node in self._revision_nodes():
            parent = node.getparent()
            entry = {
                "id": node.get(qn("w:id")),
                "type": _REVISION_TYPES[node.tag],
                "author": node.get(AUTHOR),
                "date": node.get(qn("w:date")),
                "block": None,
                "text": "",
            }
            if node.tag in (MOVE_FROM, MOVE_TO):
                entry["move"] = "from" if node.tag == MOVE_FROM else "to"
            if node.tag == PPR_CHANGE:
                entry["target"] = "paragraph format"
                entry["text"] = paragraph_text(parent.getparent())
            elif node.tag == RPR_CHANGE:
                entry["text"] = run_text(parent.getparent())
            elif node.tag in (CELL_INS, CELL_DEL):
                entry["target"] = "table cell"
            elif node.tag in PROPERTY_CHANGES:
                entry["target"] = "table format"
            elif parent.tag == TRPR:
                entry["target"] = "table row"
            elif parent.tag == RPR:
                entry["target"] = "paragraph"
            else:
                entry["text"] = "".join(t.text or "" for t in node.iter(T, DEL_TEXT))
            if part is None:
                top = self._top(node)
                entry["block"] = self._by_element.get(top) if not is_deleted_block(top) else None
            else:
                entry["part"] = part
            out.append(entry)
        return out

    def has_revisions(self) -> bool:
        """Whether any tracked change is pending. Cheaper than :meth:`revisions`:
        it stops at the first one."""
        if next(self._body.iter(*_REVISION_TYPES), None) is not None:
            return True
        return any(next(root.iter(*_REVISION_TYPES), None) is not None for _, _, root in self._story_roots())

    def _revision_nodes(self) -> list[tuple[str | None, etree._Element]]:
        """Every revision element with the story part it is in (None: the body)."""
        nodes = [(None, node) for node in self._body.iter(*_REVISION_TYPES)]
        for part, _, root in self._story_roots():
            nodes += [(part, node) for node in root.iter(*_REVISION_TYPES)]
        return nodes

    def _top(self, node: etree._Element) -> etree._Element:
        """The body-level element that contains ``node``."""
        while node.getparent() is not None and node.getparent() is not self._body:
            node = node.getparent()
        return node

    def accept_revisions(self, ids: list[str] | None = None, author: str | None = None, block: str | None = None) -> int:
        """Make tracked changes permanent. With no arguments every revision is
        accepted; ``ids`` (from :meth:`revisions`), ``author`` and ``block``
        narrow the selection. Returns the number of revisions resolved."""
        return self._resolve(True, ids, author, block)

    def reject_revisions(self, ids: list[str] | None = None, author: str | None = None, block: str | None = None) -> int:
        """Undo tracked changes, restoring the text and formatting from before
        them. Selection works as in :meth:`accept_revisions`."""
        return self._resolve(False, ids, author, block)

    def _resolve(self, accept: bool, ids: list[str] | None, author: str | None, block: str | None) -> int:
        wanted = None if ids is None else {str(i).strip() for i in ([ids] if isinstance(ids, (str, int)) else ids)}
        scope = self._get(block) if block else None
        nodes = [
            node for part, node in self._revision_nodes()
            if (wanted is None or node.get(qn("w:id")) in wanted)
            and (author is None or node.get(AUTHOR) == author)
            and (scope is None or (part is None and any(a is scope for a in node.iterancestors())))
        ]
        roots = [self._body] + [root for _, _, root in self._story_roots()]
        # The two halves of a move are resolved together.
        names = {self._move_name(node) for node in nodes if node.tag in (MOVE_FROM, MOVE_TO)} - {None}
        if names:
            chosen = set(nodes)
            nodes += [node for _, node in self._revision_nodes()
                      if node not in chosen and node.tag in (MOVE_FROM, MOVE_TO) and self._move_name(node) in names]
            for root in roots:
                self.edits.remove_move_ranges(root, names)
        if wanted is not None:
            missing = wanted - {node.get(qn("w:id")) for node in nodes}
            if missing:
                raise DocumentError(
                    f"No pending revision with id {', '.join(sorted(missing))}. Nothing was changed. "
                    "List the revisions again to see current ids."
                )

        def is_mark(node: etree._Element) -> bool:  # on a paragraph mark, a table row or a cell
            return node.tag not in PROPERTY_CHANGES and node.getparent().tag in (RPR, TRPR, TCPR)

        # A cell that goes (a rejected insertion, an accepted deletion) takes
        # its grid column with it once no row has a cell there any more.
        emptied: dict[etree._Element, set[int]] = {}
        for node in nodes:
            if node.tag == (CELL_DEL if accept else CELL_INS):
                cell = node.getparent().getparent()
                table = next(a for a in cell.iterancestors() if a.tag == TBL)
                emptied.setdefault(table, set()).add(tables.physical_col(cell))

        # Text first, then paragraph marks, rows and cells: a paragraph whose
        # inserted text was rejected is empty by the time its mark is, and
        # goes entirely.
        links = {node.getparent() for node in nodes if node.getparent().tag == HYPERLINK}
        for node in sorted(nodes, key=is_mark):
            if any(ancestor is root for ancestor in node.iterancestors() for root in roots):
                self.edits.resolve(node, accept)
        for link in links:
            self.edits.remove_if_empty(link)  # a link whose text was rejected or deleted
        for table, columns in emptied.items():
            spare = len(tables.grid_cols(table)) - tables.physical_width(table)
            for column in sorted(columns, reverse=True)[:max(spare, 0)]:
                tables.remove_grid_col(table, column)
        for root in roots:
            if next(root.iter(MOVE_FROM, MOVE_TO), None) is None:
                self.edits.remove_move_ranges(root)  # no moves left: no stray range markers either
        self._sync_ids()
        return len(nodes)

    @staticmethod
    def _move_name(node: etree._Element) -> str | None:
        """The name of the move range a ``w:moveFrom`` / ``w:moveTo`` belongs to."""
        start_tag = MOVE_RANGE_STARTS[0 if node.tag == MOVE_FROM else 1]
        paragraph = next((a for a in node.iterancestors() if a.tag == P), None)
        start = paragraph.find(start_tag) if paragraph is not None else None
        return start.get(qn("w:name")) if start is not None else None

    def _sync_ids(self) -> None:
        """Forget blocks that left the body and give ids to blocks that came back."""
        for element in [e for e in self._by_element if not self._attached(e)]:
            self._ids.pop(self._by_element.pop(element), None)
        for element in self._block_elements():
            self._id_of(element)
