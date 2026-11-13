"""EditContext: the single path through which the document is mutated.

Every insert and delete of a run, a block or a table cell goes through six
primitives (``insert_run``, ``delete_run``, ``insert_block``, ``delete_block``,
``insert_cell``, ``delete_cell``), and every formatting change announces
itself through a hook (``before_run_format``, ``before_paragraph_format``,
``before_cell_format``, ``before_row_format``, ``before_table_format``).

With ``track_changes=False`` the primitives edit the XML directly. With
``track_changes=True`` they record Word revisions instead, so the result opens
in Word as redlines the reader can accept or reject:

- inserted runs are wrapped in ``w:ins``; deleted runs stay in the file,
  wrapped in ``w:del`` with their text turned into ``w:delText``
- inserted and deleted paragraphs and table rows carry ``w:ins`` / ``w:del``
  on the paragraph mark (``w:pPr/w:rPr``) or the row (``w:trPr``)
- inserted and deleted table cells (a column is one cell per row) carry
  ``w:cellIns`` / ``w:cellDel`` in ``w:tcPr``
- formatting changes keep the previous properties in ``w:rPrChange`` /
  ``w:pPrChange`` / ``w:tcPrChange`` / ``w:trPrChange`` / ``w:tblPrChange``
- a moved paragraph stays where it was as ``w:moveFrom`` and appears at its
  new place as ``w:moveTo``, the two linked by a named move range

``resolve`` accepts or rejects one such revision; it is never tracked itself.

Undoing your own pending insertion removes it outright rather than recording
a deletion of an insertion, which is what Word does.
"""

from __future__ import annotations

import copy
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Iterator

from lxml import etree

from .ooxml import (
    AUTHOR,
    CELL_DEL,
    CELL_INS,
    CELL_MERGE,
    DEL,
    DEL_INSTR_TEXT,
    DEL_TEXT,
    INS,
    INSTR_TEXT,
    MOVE_FROM,
    MOVE_RANGE_ENDS,
    MOVE_RANGE_STARTS,
    MOVE_TO,
    P,
    PPR,
    PPR_CHANGE,
    PPR_ORDER,
    PROPERTY_CHANGES,
    RPR,
    RPR_CHANGE,
    SECTPR,
    T,
    TBLPR,
    TBLPREX,
    TC,
    TCPR,
    TCPR_ORDER,
    TR,
    TRPR,
    TRPR_CHANGE,
    el,
    ensure_child,
    get_or_add_props,
    insert_in_order,
    get_or_add_ppr,
    get_or_add_rpr,
    iter_runs,
    paragraph_mark,
    qn,
    row_mark,
    strip_revisions,
    unwrap,
)

_ID = qn("w:id")
_WRAPPERS = (INS, DEL, MOVE_FROM, MOVE_TO)
_TR_CHANGE = TRPR_CHANGE
_MARK_ORDER = (INS, DEL, MOVE_FROM, MOVE_TO)  # order of revision marks on a paragraph mark
_NAME = qn("w:name")
# What survives in a property element when a rejected change restores the
# previous properties, and whether it goes before or after them.
_KEPT_ON_RESTORE = {
    RPR: ((INS, DEL, MOVE_FROM, MOVE_TO), "before"),  # revision marks on a paragraph mark
    PPR: ((RPR, SECTPR), "after"),
    TRPR: ((INS, DEL), "after"),
    TCPR: ((CELL_INS, CELL_DEL, CELL_MERGE), "after"),
    TBLPR: ((), "after"),
}


@dataclass
class EditContext:
    track_changes: bool = False
    author: str = "docx-harness"
    #: Timestamp written on revisions (``YYYY-MM-DDTHH:MM:SSZ``). Defaults to
    #: the current UTC time; set it for reproducible output.
    date: str | None = None
    _roots: tuple = field(default=(), init=False, repr=False)
    _next_id: int | None = field(default=None, init=False, repr=False)
    _suspended: int = field(default=0, init=False, repr=False)

    # ----------------------------------------------------------------- state

    def bind(self, *roots: etree._Element) -> None:
        """Attach to the roots of the parts that will be edited (the body,
        headers, footers, notes) so new revision ids avoid existing ones."""
        self._roots = roots
        self._next_id = None

    @property
    def tracking(self) -> bool:
        return self.track_changes and not self._suspended

    @contextmanager
    def untracked(self) -> Iterator[None]:
        """Edit directly for the duration, e.g. while preparing content that is
        not in the document yet and will be inserted as one tracked block."""
        self._suspended += 1
        try:
            yield
        finally:
            self._suspended -= 1

    def _new_id(self) -> str:
        if self._next_id is None:
            highest = 0
            for root in self._roots:
                for element in root.iter():
                    value = element.get(_ID) if isinstance(element.tag, str) else None
                    if value and value.lstrip("-").isdigit():
                        highest = max(highest, int(value))
            self._next_id = highest + 1
        self._next_id += 1
        return str(self._next_id - 1)

    def _revision(self, tag: str) -> etree._Element:
        date = self.date or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        return el(tag, {"w:id": self._new_id(), "w:author": self.author, "w:date": date})

    def _mine(self, element: etree._Element | None) -> bool:
        """Is this a pending insertion made by this author?"""
        return element is not None and element.tag == INS and element.get(AUTHOR) == self.author

    # ------------------------------------------------------------------ runs

    def insert_run(self, parent: etree._Element, index: int, run: etree._Element) -> None:
        if not self.tracking:
            parent.insert(index, run)
            return
        # Revisions do not nest (no w:ins inside w:ins or w:del), so step out
        # of another author's wrapper, splitting it when inserting mid-way.
        while parent.tag in _WRAPPERS and not self._mine(parent):
            parent, index = self._step_out(parent, index)
        if self._mine(parent):
            parent.insert(index, run)  # already inside our own insertion
            return
        wrapper = self._revision("w:ins")
        wrapper.append(run)
        parent.insert(index, wrapper)

    def _step_out(self, wrapper: etree._Element, index: int) -> tuple[etree._Element, int]:
        """The position in ``wrapper``'s parent equivalent to ``index`` inside it."""
        outer = wrapper.getparent()
        position = outer.index(wrapper)
        if index <= 0:
            return outer, position
        if index < len(wrapper):
            rest = etree.Element(wrapper.tag, attrib=dict(wrapper.attrib), nsmap=wrapper.nsmap)
            rest.set(_ID, self._new_id())
            for child in list(wrapper)[index:]:
                rest.append(child)
            wrapper.addnext(rest)
        return outer, position + 1

    def delete_run(self, run: etree._Element) -> None:
        parent = run.getparent()
        if parent is None:
            return
        if not self.tracking or self._mine(parent):
            parent.remove(run)
            # Do not leave an empty tracked-insertion wrapper behind.
            if parent.tag == INS and len(parent) == 0 and parent.getparent() is not None:
                parent.getparent().remove(parent)
            return
        if parent.tag in (DEL, MOVE_FROM):
            return  # already deleted
        wrapper = self._revision("w:del")
        wrapper.tail, run.tail = run.tail, None
        parent.insert(parent.index(run), wrapper)
        wrapper.append(run)
        for text in run.iter(T, INSTR_TEXT):
            text.tag = DEL_TEXT if text.tag == T else DEL_INSTR_TEXT

    # ---------------------------------------------------------------- blocks

    def insert_block(self, parent: etree._Element, index: int, block: etree._Element) -> None:
        """Insert a paragraph, table, table row or content control."""
        parent.insert(index, block)
        if not self.tracking:
            return
        for tr in block.iter(TR):
            if row_mark(tr, INS) is None:
                self._mark_row(tr, "w:ins")
        self._mark_content_inserted(block)

    def _mark_content_inserted(self, element: etree._Element) -> None:
        """Mark every paragraph in ``element``, text and paragraph mark, as inserted."""
        for p in element.iter(P):
            for run in list(iter_runs(p)):
                holder = run.getparent()
                if holder.tag not in _WRAPPERS:
                    wrapper = self._revision("w:ins")
                    holder.insert(holder.index(run), wrapper)
                    wrapper.append(run)
            if paragraph_mark(p, INS) is None:
                self._mark_paragraph(p, "w:ins")

    def delete_block(self, block: etree._Element) -> None:
        parent = block.getparent()
        if parent is None:
            return
        if not self.tracking or self._inserted_here(block):
            parent.remove(block)
            return
        # The block stays in the file, marked as deleted, until the reader
        # accepts the change. Readers skip it (see ``is_deleted_block``).
        for tr in block.iter(TR):
            if row_mark(tr, DEL) is None:
                self._mark_row(tr, "w:del")
        self._mark_content_deleted(block)

    def _mark_content_deleted(self, element: etree._Element) -> None:
        for p in element.iter(P):
            for run in list(iter_runs(p)):
                self.delete_run(run)
            if paragraph_mark(p, DEL) is None:
                self._mark_paragraph(p, "w:del")

    # ----------------------------------------------------------- table cells

    def insert_cell(self, parent: etree._Element, index: int, tc: etree._Element) -> None:
        """Insert a table cell into a row (inserting a column is one call per row)."""
        parent.insert(index, tc)
        if not self.tracking:
            return
        if _cell_mark(tc, CELL_INS) is None:
            self._mark_cell(tc, "w:cellIns")
        self._mark_content_inserted(tc)

    def delete_cell(self, tc: etree._Element) -> None:
        parent = tc.getparent()
        if parent is None:
            return
        if not self.tracking or self._mine_move(_cell_mark(tc, CELL_INS)):
            parent.remove(tc)  # untracked, or our own pending insertion simply goes
            return
        if _cell_mark(tc, CELL_DEL) is not None:
            return  # already deleted
        # Like a deleted row, the cell stays in the file until the reader
        # accepts the change, and readers skip it (see ``tables.row_cells``).
        self._mark_cell(tc, "w:cellDel")
        self._mark_content_deleted(tc)

    def _mark_cell(self, tc: etree._Element, tag: str) -> None:
        insert_in_order(get_or_add_props(tc, "w:tcPr"), self._revision(tag), TCPR_ORDER)

    def move_block(self, block: etree._Element, parent: etree._Element, index: int) -> etree._Element:
        """Move a block to position ``index`` of ``parent`` (counted with the
        block still where it is). Returns the element that now holds the
        content: the block itself, or a copy when the move is tracked and the
        original has to stay behind as the "moved from" half."""
        target = parent[index] if index < len(parent) else None
        if target is block or (block.getparent() is parent and block.getnext() is target):
            return block  # already there

        def place(element: etree._Element) -> None:
            if target is None:
                parent.append(element)
            else:
                target.addprevious(element)

        own = self._inserted_here(block) or (block.tag == P and self._mine_move(paragraph_mark(block, MOVE_TO)))
        if not self.tracking or own:
            block.tail = None
            place(block)  # our own pending insertion or move simply goes elsewhere
            return block
        moved = copy.deepcopy(block)
        moved.tail = None
        strip_revisions(moved)  # the copy starts clean
        if block.tag != P:
            # Word has no move markup for tables: a deletion here, an insertion there.
            self.delete_block(block)
            place(moved)
            self.insert_block(parent, parent.index(moved), moved)
            return moved
        name = f"move{self._new_id()}"
        self._mark_move(block, "w:moveFrom", name)
        self._mark_move(moved, "w:moveTo", name)
        place(moved)
        return moved

    def _mine_move(self, element: etree._Element | None) -> bool:
        return element is not None and element.get(AUTHOR) == self.author

    def _mark_move(self, p: etree._Element, tag: str, name: str) -> None:
        """Mark a paragraph as one half of a move: its runs, its paragraph
        mark, and a named range that pairs it with the other half."""
        for run in list(iter_runs(p)):
            holder = run.getparent()
            wrapper = self._revision(tag)
            wrapper.tail, run.tail = run.tail, None
            holder.insert(holder.index(run), wrapper)
            wrapper.append(run)
        self._mark_paragraph(p, tag)
        start = self._revision(f"{tag}RangeStart")
        start.set(_NAME, name)
        p.insert(1 if len(p) and p[0].tag == PPR else 0, start)
        p.append(el(f"{tag}RangeEnd", {"w:id": start.get(_ID)}))

    def remove_move_ranges(self, root: etree._Element, names: set[str] | None = None) -> None:
        """Remove the range markers of the named moves (all of them when
        ``names`` is None) once their revisions are being resolved."""
        ids = set()
        for start in list(root.iter(*MOVE_RANGE_STARTS)):
            if names is None or start.get(_NAME) in names:
                ids.add((start.tag.replace("RangeStart", "RangeEnd"), start.get(_ID)))
                start.getparent().remove(start)
        for end in list(root.iter(*MOVE_RANGE_ENDS)):
            if names is None or (end.tag, end.get(_ID)) in ids:
                end.getparent().remove(end)

    def _inserted_here(self, block: etree._Element) -> bool:
        """Was the whole block inserted by this author and not yet accepted?"""
        if block.tag == P:
            return self._mine(paragraph_mark(block, INS))
        if block.tag == TR:
            return self._mine(row_mark(block, INS))
        if block.tag == TC:
            return self._mine_move(_cell_mark(block, CELL_INS))
        rows = list(block.iter(TR))
        if rows:
            return all(self._mine(row_mark(tr, INS)) for tr in rows)
        paragraphs = list(block.iter(P))
        return bool(paragraphs) and all(self._mine(paragraph_mark(p, INS)) for p in paragraphs)

    def _mark_paragraph(self, p: etree._Element, tag: str) -> None:
        rpr = ensure_child(get_or_add_ppr(p), "w:rPr", PPR_ORDER)
        # Revision marks lead the paragraph mark's properties, in schema order.
        earlier = _MARK_ORDER[:_MARK_ORDER.index(qn(tag))]
        index = 0
        while index < len(rpr) and rpr[index].tag in earlier:
            index += 1
        rpr.insert(index, self._revision(tag))

    def _mark_row(self, tr: etree._Element, tag: str) -> None:
        trpr = tr.find(TRPR)
        if trpr is None:
            trpr = el("w:trPr")
            tr.insert(1 if len(tr) and tr[0].tag == TBLPREX else 0, trpr)
        mark = self._revision(tag)
        change = trpr.find(_TR_CHANGE)
        if change is not None:
            change.addprevious(mark)
        else:
            trpr.append(mark)

    # ------------------------------------------------------------ formatting

    def before_run_format(self, run: etree._Element) -> None:
        """Call before changing a run's properties. When tracking, the current
        properties are kept in ``w:rPrChange`` (once per run)."""
        if not self.tracking or self._mine(run.getparent()):
            return
        rpr = get_or_add_rpr(run)
        if rpr.find(RPR_CHANGE) is not None:
            return
        previous = el("w:rPr")
        for child in rpr:
            previous.append(copy.deepcopy(child))
        change = self._revision("w:rPrChange")
        change.append(previous)
        rpr.append(change)

    def before_paragraph_format(self, p: etree._Element) -> None:
        """Call before changing a paragraph's properties or style. When
        tracking, the current properties are kept in ``w:pPrChange``."""
        if not self.tracking or self._mine(paragraph_mark(p, INS)):
            return
        ppr = get_or_add_ppr(p)
        if ppr.find(PPR_CHANGE) is not None:
            return
        previous = el("w:pPr")
        for child in ppr:
            if child.tag not in (RPR, SECTPR):
                previous.append(copy.deepcopy(child))
        change = self._revision("w:pPrChange")
        change.append(previous)
        ppr.append(change)

    def before_cell_format(self, tc: etree._Element) -> None:
        """Call before changing a cell's properties (width, span, merge,
        shading, alignment). When tracking, they are kept in ``w:tcPrChange``."""
        if self.tracking and not self._mine_move(_cell_mark(tc, CELL_INS)):
            self._keep_previous(get_or_add_props(tc, "w:tcPr"), "w:tcPrChange", (CELL_INS, CELL_DEL, CELL_MERGE))

    def before_row_format(self, tr: etree._Element) -> None:
        """Call before changing a row's properties (``w:trPrChange``)."""
        if self.tracking and not self._mine(row_mark(tr, INS)):
            self._keep_previous(get_or_add_props(tr, "w:trPr"), "w:trPrChange", (INS, DEL))

    def before_table_format(self, tbl: etree._Element) -> None:
        """Call before changing a table's properties (``w:tblPrChange``)."""
        if self.tracking and not self._inserted_here(tbl):
            self._keep_previous(get_or_add_props(tbl, "w:tblPr"), "w:tblPrChange", ())

    def _keep_previous(self, props: etree._Element, change_tag: str, skip: tuple) -> None:
        """Store a copy of ``props`` in a property-change revision, once."""
        if props.find(qn(change_tag)) is not None:
            return
        previous = el("w:" + etree.QName(props).localname)
        for child in props:
            if child.tag not in skip:
                previous.append(copy.deepcopy(child))
        change = self._revision(change_tag)
        change.append(previous)
        props.append(change)

    # ------------------------------------------------- accepting / rejecting

    def resolve(self, node: etree._Element, accept: bool) -> None:
        """Accept or reject one revision element (``w:ins``, ``w:del``,
        ``w:moveFrom``, ``w:moveTo`` or a ``w:*PrChange``), making the change
        permanent or undoing it. This is never recorded as a revision."""
        parent = node.getparent()
        if parent is None:
            return
        if node.tag in PROPERTY_CHANGES:
            parent.remove(node)
            if not accept:
                self._restore_properties(parent, node)
            return
        # Content goes when an insertion is rejected or a deletion accepted.
        goes = (node.tag in (INS, MOVE_TO, CELL_INS)) != accept
        if node.tag in (CELL_INS, CELL_DEL):
            cell = parent.getparent()
            parent.remove(node)
            if goes and cell is not None and cell.getparent() is not None:
                row = next((a for a in cell.iterancestors() if a.tag == TR), None)
                cell.getparent().remove(cell)
                if row is not None and next(row.iter(TC), None) is None:
                    self._remove_row(row)  # a row cannot be left without cells
            return
        if parent.tag == TRPR:
            row = parent.getparent()
            _remove_with_empty_holders(node)
            if goes:
                self._remove_row(row)
        elif parent.tag == RPR and parent.getparent() is not None and parent.getparent().tag == PPR:
            paragraph = parent.getparent().getparent()
            _remove_with_empty_holders(node)
            if goes:
                self._remove_paragraph_mark(paragraph)
        elif goes:
            parent.remove(node)
        else:
            if node.tag == DEL:
                for text in node.iter(DEL_TEXT, DEL_INSTR_TEXT):
                    text.tag = T if text.tag == DEL_TEXT else INSTR_TEXT
            unwrap(node)

    @staticmethod
    def _restore_properties(props: etree._Element, change: etree._Element) -> None:
        """Put back the properties stored in a ``w:*PrChange``."""
        kept_tags, where = _KEPT_ON_RESTORE.get(props.tag, ((), "after"))
        kept = [child for child in props if child.tag in kept_tags]
        previous = list(change[0]) if len(change) else []
        for child in list(props):
            props.remove(child)
        for child in (kept + previous) if where == "before" else (previous + kept):
            props.append(child)

    @staticmethod
    def _remove_row(tr: etree._Element) -> None:
        tbl = tr.getparent()
        if tbl is None:
            return
        tbl.remove(tr)
        if tbl.find(TR) is None and tbl.getparent() is not None:
            tbl.getparent().remove(tbl)  # a table cannot be left without rows

    @staticmethod
    def _remove_paragraph_mark(p: etree._Element) -> None:
        """Take out the end of a paragraph: an empty paragraph disappears, and
        one with content joins the paragraph after it."""
        holder = p.getparent()
        if holder is None:
            return
        ppr = p.find(PPR)
        if ppr is not None and ppr.find(SECTPR) is not None:
            return  # the paragraph carries a section break; keep it
        content = [child for child in p if child.tag not in (PPR, *MOVE_RANGE_STARTS, *MOVE_RANGE_ENDS)]
        following = p.getnext()
        if not content:
            # A table cell has to end with a paragraph.
            previous = p.getprevious()
            ends_cell = holder.tag == TC and following is None and (previous is None or previous.tag != P)
            if not ends_cell:
                holder.remove(p)
        elif following is not None and following.tag == P:
            index = 1 if len(following) and following[0].tag == PPR else 0
            for child in content:
                following.insert(index, child)
                index += 1
            holder.remove(p)


def _cell_mark(tc: etree._Element, tag: str) -> etree._Element | None:
    tcpr = tc.find(TCPR)
    return tcpr.find(tag) if tcpr is not None else None


def _remove_with_empty_holders(mark: etree._Element) -> None:
    """Remove a revision mark, and the property elements it leaves empty."""
    holder = mark.getparent()
    holder.remove(mark)
    while holder is not None and len(holder) == 0 and holder.tag in (RPR, PPR, TRPR):
        outer = holder.getparent()
        if outer is None:
            break
        outer.remove(holder)
        holder = outer if outer.tag in (RPR, PPR, TRPR) else None
