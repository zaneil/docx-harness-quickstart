"""Table editing on the column grid: cells, rows, columns, merges and table formatting.

The pure-XML grid helpers it calls are in ``tables.py``.

Mixed into ``WordDocument`` (``document.py``), which provides the block-id
helpers it relies on (``_get``, ``_body_block``, ``_table``, ``_paragraphs_for``,
``_describe``...) and ``self.edits``, the single path for every mutation."""

from __future__ import annotations

import copy

from lxml import etree

from . import styling, tables
from .common import (
    PAGE_WIDTH_TWIPS,
    DocumentError,
    _lines,
)
from .ooxml import (
    PPR,
    RPR,
    TBL,
    TCPR,
    TR,
    TRPR,
    P,
    el,
    is_deleted_block,
    iter_runs,
    make_paragraph,
    make_run,
    paragraph_text,
    remove_child,
    strip_revisions,
)

_VERTICAL = {"top": "top", "center": "center", "centre": "center", "middle": "center", "bottom": "bottom"}
_TABLE_ALIGNMENTS = {"left": "left", "center": "center", "centre": "center", "right": "right"}
_NOT_COPIED_TO_NEW_CELLS = ("w:gridSpan", "w:hMerge", "w:vMerge", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange")


class TablesMixin:
    # ---------------------------------------------------------------- tables

    @staticmethod
    def _grid(tbl: etree._Element) -> tables.Grid:
        if tbl.tag != TBL:
            raise DocumentError("Cell references like 'b3[0,1]' only apply to tables.")
        return tables.build_grid(tbl)

    @staticmethod
    def _position(grid: tables.Grid, row: int, col: int) -> tables.Position:
        if not 0 <= row < len(grid.rows):
            raise DocumentError(f"Row {row} is out of range; the table has {len(grid.rows)} rows (zero-based).")
        position = grid.at(row, col)
        if position is None:
            raise DocumentError(
                f"Column {col} is out of range; row {row} has {grid.row_extent(row)} columns (zero-based)."
            )
        return position

    def _cell(self, tbl: etree._Element, row: int, col: int) -> etree._Element:
        """The cell at a grid position: for a position covered by a merged
        cell, the merged cell's top-left ``w:tc``, which holds its content."""
        return self._position(self._grid(tbl), row, col).slot.tc

    def _slots(self, tbl: etree._Element, row: int | None, col: int | None) -> list[tables.Slot]:
        """The cells a reference selects; None is ``*``."""
        grid = self._grid(tbl)
        if row is not None and col is not None:
            return [self._position(grid, row, col).slot]
        if row is not None and not 0 <= row < len(grid.rows):
            raise DocumentError(f"Row {row} is out of range; the table has {len(grid.rows)} rows (zero-based).")
        if col is not None and not 0 <= col < grid.width:
            raise DocumentError(f"Column {col} is out of range; the table has {grid.width} columns (zero-based).")
        return [
            slot for slot in grid.slots()
            if (row is None or slot.row <= row < slot.row + slot.rowspan)
            and (col is None or slot.col <= col < slot.col + slot.colspan)
        ]

    @staticmethod
    def _paragraph_like(source: etree._Element | None, text: str = "") -> etree._Element:
        """A new paragraph with the properties of ``source`` and the
        character formatting of its first run."""
        if source is None:
            return make_paragraph(text)
        new = el("w:p")
        if source.find(PPR) is not None:
            new.append(copy.deepcopy(source.find(PPR)))
            strip_revisions(new)  # no revision marks or formatting history on the copy
        if text:
            new.append(make_run(text, next((r.find(RPR) for r in iter_runs(source) if r.find(RPR) is not None), None)))
        return new

    def _set_cell(self, tc: etree._Element, text: str) -> None:
        """Set a cell's text, one paragraph per line. Lines map onto the
        paragraphs the cell already has, which keep their style and
        formatting; nested tables in the cell are left as they are."""
        lines = _lines(text)
        paragraphs = tables.cell_paragraphs(tc)
        if not paragraphs:
            paragraphs = [make_paragraph()]
            self.edits.insert_block(tc, len(tc), paragraphs[0])
        for p, line in zip(paragraphs, lines):
            self._set_paragraph_text(p, line)
        anchor = paragraphs[-1]
        for line in lines[len(paragraphs):]:
            new = self._paragraph_like(anchor, line)
            parent = anchor.getparent()
            self.edits.insert_block(parent, parent.index(anchor) + 1, new)
            anchor = new
        for extra in paragraphs[len(lines):]:
            previous = extra.getprevious()
            if extra.getnext() is None and (previous is None or previous.tag != P):
                self._set_paragraph_text(extra, "")  # a cell has to end with a paragraph
            else:
                self.edits.delete_block(extra)

    def _new_cell(self, like: etree._Element, text: str = "", width: int | None = None) -> etree._Element:
        """A detached, unmerged cell that looks like ``like``."""
        tc = el("w:tc")
        if like.find(TCPR) is not None:
            tcpr = copy.deepcopy(like.find(TCPR))
            for tag in _NOT_COPIED_TO_NEW_CELLS:
                remove_child(tcpr, tag)
            tc.append(tcpr)
        if width:
            tables.set_cell_width(tc, width)
        source = next(iter(tables.cell_paragraphs(like)), None)
        for line in _lines(text):
            tc.append(self._paragraph_like(source, line))
        return tc

    def _sync_cell_widths(self, tbl: etree._Element) -> None:
        """Make cell widths (and a fixed table width) agree with the grid."""
        for tr in tbl.iter(TR):
            if next(a for a in tr.iterancestors() if a.tag == TBL) is not tbl:
                continue  # a row of a nested table
            for tc in tables.row_cells(tr, include_deleted=True):
                width = tables.width_from_grid(tbl, tc)
                if width:
                    tables.set_cell_width(tc, width)
        current = tables.table_width(tbl)
        if current is not None and current[1] == "dxa" and current[0] > 0:
            tables.set_table_width(tbl, sum(tables.grid_widths(tbl)))

    def set_cell_text(self, table_id: str, row: int, col: int, text: str) -> None:
        """Set the text of the cell at grid position (row, col). Each line of
        ``text`` is a paragraph. A position inside a merged cell sets the
        merged cell."""
        self._set_cell(self._cell(self._table(table_id), row, col), text)

    def insert_table_row(self, table_id: str, cells: list[str], index: int | None = None) -> int:
        """Insert a row at ``index`` (default: append). The new row copies the
        layout of a neighbouring row. ``cells`` fills the row's own cells left
        to right; a cell merged with the rows above and below is skipped.
        Returns the index of the new row."""
        tbl = self._table(table_id)
        self._modernise_merges(tbl)
        grid = tables.build_grid(tbl)
        rows = grid.rows
        if not rows:
            raise DocumentError("Table has no rows to copy the layout from.")
        if index is None:
            index = len(rows)
        if not 0 <= index <= len(rows):
            raise DocumentError(f"index must be between 0 and {len(rows)}.")
        # Copy the last row unless inserting at the very top, so a bold header
        # row is not used as the pattern for data rows.
        source = rows[0] if len(rows) == 1 else rows[max(index - 1, 1)] if index > 0 else rows[0]
        new_row = copy.deepcopy(source)
        strip_revisions(new_row)  # the copy starts clean; it is tracked as one inserted row
        # Header rows stay one block at the top of the table.
        if new_row.find(TRPR) is not None and not (index < len(rows) and tables.is_header_row(rows[index])):
            remove_child(new_row.find(TRPR), "w:tblHeader")
        own: list[etree._Element] = []
        col = tables.grid_before(new_row)
        with self.edits.untracked():
            for tc in tables.row_cells(new_row):
                for inner in tables.nested_tables(tc):
                    self.edits.delete_block(inner)  # a new row does not repeat a nested table
                above, below = grid.at(index - 1, col), grid.at(index, col)
                inside = above is not None and below is not None and above.slot is below.slot
                if tables.v_merge(tc) or inside:
                    # Part of a vertical merge only when the row lands inside one.
                    tables.set_v_merge(tc, "continue" if inside else None)
                if inside:
                    self._set_cell(tc, "")
                else:
                    own.append(tc)
                col += tables.span(tc)
            if len(cells) > len(own):
                raise DocumentError(f"Got {len(cells)} cells but the new row has {len(own)} cells.")
            for i, tc in enumerate(own):
                self._set_cell(tc, str(cells[i]) if i < len(cells) and cells[i] is not None else "")
        if index == len(rows):
            parent = rows[-1].getparent()
            self.edits.insert_block(parent, parent.index(rows[-1]) + 1, new_row)
        else:
            parent = rows[index].getparent()
            self.edits.insert_block(parent, parent.index(rows[index]), new_row)
        return index

    def delete_table_row(self, table_id: str, row: int) -> None:
        """Delete a row. A cell merged down from this row keeps its content:
        the merge then starts in the row below."""
        tbl = self._table(table_id)
        grid = tables.build_grid(tbl)
        rows = grid.rows
        if not 0 <= row < len(rows):
            raise DocumentError(f"Row {row} is out of range; the table has {len(rows)} rows (zero-based).")
        if len(rows) == 1:
            raise DocumentError("Cannot delete the only row; delete the table block instead.")
        for col, position in sorted(grid.positions[row].items()):
            slot = position.slot
            if (slot.row, slot.col) != (row, col) or slot.rowspan == 1:
                continue
            heir = grid.at(row + 1, col).own
            self.edits.before_cell_format(heir)
            tables.set_v_merge(heir, "restart" if slot.rowspan > 2 else None)
            old = [child for child in heir if child.tag != TCPR]
            for child in slot.tc:
                if child.tag != TCPR and not is_deleted_block(child):
                    clone = copy.deepcopy(child)
                    strip_revisions(clone)
                    self.edits.insert_block(heir, len(heir), clone)
            for child in old:
                self.edits.delete_block(child)
        self.edits.delete_block(rows[row])

    def insert_table_column(
        self,
        table_id: str,
        index: int | None = None,
        cells: list[str] | None = None,
        width_in: float | None = None,
    ) -> int:
        """Insert a column at grid position ``index`` (default: after the
        last). ``cells`` gives its text top to bottom, one entry per row. The
        new cells copy the look of the column beside them, and a merged cell
        that spans the position grows by one column instead. Without
        ``width_in`` the columns are rescaled so the table keeps its width
        (when changes are tracked, the other columns are left as they are).
        Returns the index of the new column."""
        tbl = self._table(table_id)
        self._modernise_merges(tbl)
        grid = tables.build_grid(tbl)
        if not grid.rows or not grid.width:
            raise DocumentError("Table has no cells to copy the layout from.")
        if index is None:
            index = grid.width
        if not 0 <= index <= grid.width:
            raise DocumentError(f"index must be between 0 and {grid.width}.")
        cells = list(cells or [])
        if len(cells) > len(grid.rows):
            raise DocumentError(f"Got {len(cells)} cells but the table has {len(grid.rows)} rows. Give one text per row.")
        widths = tables.grid_widths(tbl)
        total = sum(widths)
        physical = grid.physical_of(index)
        if physical is None:
            physical = (grid.physical_of(index - 1) or 0) + 1
        beside = grid.physical_of(index - 1 if index > 0 else 0)
        if width_in is not None:
            new_width = self._twips([width_in], 1)[0]
        else:
            new_width = (widths[beside] if beside is not None and beside < len(widths) else 0) or PAGE_WIDTH_TWIPS // (grid.width + 1)

        widened = []
        for r, tr in enumerate(grid.rows):
            here, left = grid.at(r, index), grid.at(r, index - 1)
            if here is not None and here.slot.col < index:
                # A merged cell straddles the new column: it grows by one.
                self.edits.before_cell_format(here.own)
                tables.set_span(here.own, tables.span(here.own) + 1)
                widened.append(here.own)
                continue
            if here is None and left is None:
                if index <= tables.grid_before(tr):  # the row starts further right
                    self.edits.before_row_format(tr)
                    tables.set_grid_before(tr, tables.grid_before(tr) + 1)
                continue
            text = str(cells[r]) if r < len(cells) and cells[r] is not None else ""
            new = self._new_cell((left or here).own, text, new_width)
            if here is not None:
                parent = here.own.getparent()
                self.edits.insert_cell(parent, parent.index(here.own), new)
            else:
                parent = left.own.getparent()
                self.edits.insert_cell(parent, parent.index(left.own) + 1, new)
        tables.insert_grid_col(tbl, physical, new_width)
        if self.edits.tracking:
            for tc in widened:
                width = tables.width_from_grid(tbl, tc)
                if width:
                    tables.set_cell_width(tc, width)
        else:
            if width_in is None:
                tables.scale_grid(tbl, total)
            self._sync_cell_widths(tbl)
        return index

    def delete_table_column(self, table_id: str, col: int) -> None:
        """Delete the column at grid position ``col``. A merged cell that
        spans it loses one column and keeps its content."""
        tbl = self._table(table_id)
        self._modernise_merges(tbl)
        grid = tables.build_grid(tbl)
        if not 0 <= col < grid.width:
            raise DocumentError(f"Column {col} is out of range; the table has {grid.width} columns (zero-based).")
        if grid.width == 1:
            raise DocumentError("Cannot delete the only column; delete the table block instead.")
        widths = tables.grid_widths(tbl)
        total = sum(widths)
        physical = grid.physical_of(col)
        gone = widths[physical] if physical is not None and physical < len(widths) else 0
        for r, tr in enumerate(grid.rows):
            here = grid.at(r, col)
            if here is None:
                if col < tables.grid_before(tr):
                    self.edits.before_row_format(tr)
                    tables.set_grid_before(tr, tables.grid_before(tr) - 1)
                continue
            own = here.own
            if tables.span(own) > 1:
                self.edits.before_cell_format(own)
                tables.set_span(own, tables.span(own) - 1)
                current = tables.cell_width(own)
                if self.edits.tracking and gone and current is not None and current[0] > gone:
                    tables.set_cell_width(own, current[0] - gone)
            else:
                self.edits.delete_cell(own)
                if not tables.row_cells(tr, include_deleted=True):
                    self.edits.delete_block(tr)  # a row cannot be left without cells
        if not self.edits.tracking:
            # A tracked deletion keeps its cells, and so its grid column, until accepted.
            if physical is not None:
                tables.remove_grid_col(tbl, physical)
            tables.scale_grid(tbl, total)
            self._sync_cell_widths(tbl)

    def _modernise_merges(self, tbl: etree._Element) -> None:
        """Rewrite legacy horizontal merges (``w:hMerge``, one cell per
        column) as spans (``w:gridSpan``, one cell) before the table's
        structure is edited. The two say the same thing, so this is not a
        tracked change."""
        runs = tables.h_merge_runs(tbl)
        if not runs:
            return
        with self.edits.untracked():
            for first, swallowed in runs:
                if first is None:
                    for tc in swallowed:
                        tables.clear_h_merge(tc)
                    continue
                columns = tables.span(first) + sum(tables.span(tc) for tc in swallowed)
                widths = [tables.cell_width(tc) for tc in [first, *swallowed]]
                for tc in swallowed:
                    for child in [c for c in tc if c.tag == TBL or (c.tag == P and paragraph_text(c))]:
                        self.edits.move_block(child, first, len(first))
                    self.edits.delete_cell(tc)
                tables.clear_h_merge(first)
                tables.set_span(first, columns)
                if all(w is not None and w[1] == "dxa" for w in widths):
                    tables.set_cell_width(first, sum(w[0] for w in widths))

    def _no_tracked_merge(self) -> None:
        if self.edits.tracking:
            raise DocumentError(
                "Merging and unmerging cells cannot be recorded as a tracked change, so nothing was changed. "
                "Leave the cells as they are and tell the user this edit needs track changes switched off."
            )

    def merge_cells(self, table_id: str, row: int, col: int, end_row: int, end_col: int) -> None:
        """Merge the rectangle of cells from (row, col) to (end_row, end_col),
        both included, into one cell. The text of all of them is kept, in
        reading order. Not available while changes are tracked."""
        tbl = self._table(table_id)
        self._modernise_merges(tbl)
        self._no_tracked_merge()
        grid = tables.build_grid(tbl)
        (r0, r1), (c0, c1) = sorted((row, end_row)), sorted((col, end_col))
        slots: list[tables.Slot] = []
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                slot = self._position(grid, r, c).slot
                if (slot.row < r0 or slot.col < c0
                        or slot.row + slot.rowspan - 1 > r1 or slot.col + slot.colspan - 1 > c1):
                    raise DocumentError(
                        f"The range cuts through the merged cell at [{slot.row},{slot.col}], which covers "
                        f"{slot.rowspan} row(s) and {slot.colspan} column(s). Widen the range to take in all "
                        "of it, or unmerge it first. Nothing was changed."
                    )
                if slot not in slots:
                    slots.append(slot)
        if len(slots) == 1:
            raise DocumentError("That range is one cell already; give two different corners to merge.")

        def has_content(tc: etree._Element) -> bool:
            return bool(self._cell_text(tc).strip()) or bool(tables.nested_tables(tc))

        anchor = slots[0].tc
        incoming = [slot.tc for slot in slots[1:] if has_content(slot.tc)]
        if incoming and not has_content(anchor):
            for child in [c for c in anchor if c.tag != TCPR]:
                self.edits.delete_block(child)
        for tc in incoming:
            for child in [c for c in tc if c.tag != TCPR]:
                self.edits.move_block(child, anchor, len(anchor))
        for r in range(r0, r1 + 1):
            owns: list[etree._Element] = []
            for c in range(c0, c1 + 1):
                own = grid.at(r, c).own
                if own not in owns:
                    owns.append(own)
            keep = owns[0]
            for extra in owns[1:]:
                self.edits.delete_cell(extra)
            tables.set_span(keep, c1 - c0 + 1)
            tables.set_v_merge(keep, None if r0 == r1 else "restart" if r == r0 else "continue")
            width = tables.width_from_grid(tbl, keep)
            if width:
                tables.set_cell_width(keep, width)
            if r > r0:  # the cells under the first hold nothing
                for child in [c for c in keep if c.tag != TCPR]:
                    self.edits.delete_block(child)
                self.edits.insert_block(keep, len(keep), make_paragraph())
        if len(anchor) == 0 or anchor[-1].tag != P:
            self.edits.insert_block(anchor, len(anchor), make_paragraph())  # a cell has to end with a paragraph

    def unmerge_cells(self, table_id: str, row: int, col: int) -> None:
        """Split the merged cell covering (row, col) back into single cells.
        Its content stays in the top-left one. Not available while changes
        are tracked."""
        tbl = self._table(table_id)
        self._modernise_merges(tbl)
        self._no_tracked_merge()
        slot = self._position(tables.build_grid(tbl), row, col).slot
        if not slot.merged:
            raise DocumentError(f"The cell at [{row},{col}] is not merged; there is nothing to split.")
        widths = tables.grid_widths(tbl)
        for own in slot.members:
            start = tables.physical_col(own)
            tables.set_v_merge(own, None)
            tables.set_span(own, 1)
            if start < len(widths) and widths[start]:
                tables.set_cell_width(own, widths[start])
            parent = own.getparent()
            for offset in range(1, slot.colspan):
                width = widths[start + offset] if start + offset < len(widths) else None
                self.edits.insert_cell(parent, parent.index(own) + offset, self._new_cell(own, "", width))

    def format_cells(
        self,
        ref: str,
        shading: str | None = None,
        vertical_align: str | None = None,
        border: str | None = None,
        border_sides: list[str] | str | None = None,
        border_color: str | None = None,
        border_width_pt: float | None = None,
        margins_in: float | None = None,
        text_direction: str | None = None,
    ) -> int:
        """Format table cells. ``ref`` is a cell (``b3[0,1]``), a row
        (``b3[0,*]``), a column (``b3[*,1]``) or a table id for every cell.

        ``shading`` is a background color (hex or a basic color name, "none"
        to clear); ``vertical_align`` is top, center or bottom. ``border`` is
        a line style (single, double, dashed, dotted, thick, or none to remove
        the line) drawn on ``border_sides``: any of top, bottom, left, right,
        all (the default), outer (the outline of the referenced cells taken
        together) and inner (the lines between them), in ``border_color``
        and ``border_width_pt``. ``margins_in`` is the padding inside the
        cell and ``text_direction`` is horizontal, vertical or vertical_up.

        Returns the number of cells formatted. For the text itself use
        ``format_text`` and ``set_paragraph_format`` with the same ``ref``."""
        if border is None and (border_sides is not None or border_color is not None or border_width_pt is not None):
            border = "single"
        if all(v is None for v in (shading, vertical_align, border, margins_in, text_direction)):
            raise DocumentError(
                "Give at least one cell property (shading, vertical_align, border, margins_in, text_direction)."
            )
        element, cell = self._parse_ref(ref)
        if element.tag != TBL:
            raise DocumentError(f"{ref} is a {self._type(element)}; cell formatting needs a table or a table cell.")
        fill = None
        if shading is not None:
            fill = "" if str(shading).strip().lower() in ("", "none", "clear", "auto") else styling.normalize_color(shading)
        align = None
        if vertical_align is not None:
            align = _VERTICAL.get(str(vertical_align).strip().lower())
            if align is None:
                raise DocumentError(f"Invalid vertical_align {vertical_align!r}. Use 'top', 'center' or 'bottom'.")
        line = self._border_spec(border, border_sides, border_color, border_width_pt) if border is not None else None
        padding = None
        if margins_in is not None:
            try:
                padding = round(float(margins_in) * tables.TWIPS_PER_INCH)
            except (TypeError, ValueError):
                padding = -1
            if not 0 <= padding <= tables.TWIPS_PER_INCH:
                raise DocumentError("margins_in is the padding inside a cell in inches, between 0 and 1.")
        direction = None
        if text_direction is not None:
            direction = str(text_direction).strip().lower().replace("-", "_").replace(" ", "_")
            if direction not in tables.TEXT_DIRECTIONS:
                raise DocumentError(f"Invalid text_direction {text_direction!r}. Use 'horizontal', 'vertical' or 'vertical_up'.")

        grid = self._grid(element)
        slots = self._slots(element, *(cell or (None, None)))
        for slot in slots:
            for tc in slot.members:
                self.edits.before_cell_format(tc)
                tables.apply_cell_format(tc, fill, align)
                if padding is not None:
                    tables.set_cell_margins(tc, padding)
                if direction is not None:
                    tables.set_text_direction(tc, direction)
        if line is not None:
            self._draw_borders(grid, slots, *line)
        return len(slots)

    @staticmethod
    def _border_spec(border, sides, color, width_pt) -> tuple[str, set[str], int, str]:
        style = str(border).strip().lower()
        if style not in tables.BORDER_STYLES:
            raise DocumentError(f"Invalid border {border!r}. Use one of: {', '.join(tables.BORDER_STYLES)}.")
        wanted = {str(s).strip().lower() for s in ([sides] if isinstance(sides, str) else sides or ["all"])}
        allowed = {*tables.SIDES, "all", "outer", "inner"}
        if not wanted or wanted - allowed:
            raise DocumentError(f"Invalid border_sides {sides!r}. Use any of: {', '.join(sorted(allowed))}.")
        try:
            size = round(float(width_pt) * 8) if width_pt is not None else 4
        except (TypeError, ValueError):
            size = 0
        if not 2 <= size <= 48:
            raise DocumentError("border_width_pt must be between 0.25 and 6 points.")
        return style, wanted, size, styling.normalize_color(color) if color else "auto"

    def _draw_borders(self, grid: tables.Grid, slots: list[tables.Slot], style: str, sides: set[str],
                      size: int, color: str) -> None:
        """Set cell borders. An edge between two cells belongs to both, and
        Word shows whichever it prefers, so the neighbour's facing edge is
        set as well."""
        chosen = {(slot.row, slot.col) for slot in slots}  # slots may come from another reading of the grid

        def beside(slot: tables.Slot, side: str) -> list[tables.Slot | None]:
            rows = range(slot.row, slot.row + slot.rowspan)
            cols = range(slot.col, slot.col + slot.colspan)
            cells = {
                "top": [(slot.row - 1, c) for c in cols], "bottom": [(slot.row + slot.rowspan, c) for c in cols],
                "left": [(r, slot.col - 1) for r in rows], "right": [(r, slot.col + slot.colspan) for r in rows],
            }[side]
            return [position.slot if position else None for position in (grid.at(r, c) for r, c in cells)]

        def edge_cells(slot: tables.Slot, side: str) -> list[etree._Element]:
            return {"top": slot.members[:1], "bottom": slot.members[-1:]}.get(side, slot.members)

        def within(inner: tables.Slot, outer: tables.Slot, side: str) -> bool:
            if side in ("top", "bottom"):
                return outer.col <= inner.col and inner.col + inner.colspan <= outer.col + outer.colspan
            return outer.row <= inner.row and inner.row + inner.rowspan <= outer.row + outer.rowspan

        for slot in slots:
            for side in tables.SIDES:
                near = beside(slot, side)
                inside = all(n is not None and (n.row, n.col) in chosen for n in near)
                if not ("all" in sides or side in sides or ("inner" if inside else "outer") in sides):
                    continue
                targets = [(tc, side) for tc in edge_cells(slot, side)]
                for neighbour in dict.fromkeys(n for n in near if n is not None):
                    if within(neighbour, slot, side):
                        targets += [(tc, tables.OPPOSITE[side]) for tc in edge_cells(neighbour, tables.OPPOSITE[side])]
                for tc, edge in targets:
                    self.edits.before_cell_format(tc)
                    tables.set_cell_border(tc, edge, style, size, color)

    def set_table_format(
        self,
        table_id: str,
        column_widths_in: list[float] | None = None,
        header_rows: int | None = None,
        alignment: str | None = None,
        borders: str | None = None,
        cell_margins_in: float | None = None,
    ) -> None:
        """Lay out a table. ``column_widths_in`` sets every column's width in
        inches and fixes the layout so Word keeps them; ``header_rows`` is how
        many rows at the top repeat on each page the table runs onto (0 for
        none); ``alignment`` places the table left, center or right.
        ``borders`` draws the table's lines: all, none, outer (a box) or
        horizontal (rules between rows, no vertical lines); borders set on
        single cells with ``format_cells`` still show. ``cell_margins_in`` is
        the padding inside every cell."""
        tbl = self._table(table_id)
        if all(v is None for v in (column_widths_in, header_rows, alignment, borders, cell_margins_in)):
            raise DocumentError(
                "Give at least one table property (column_widths_in, header_rows, alignment, borders, cell_margins_in)."
            )
        if borders is not None and str(borders).strip().lower() not in tables.BORDER_PRESETS:
            raise DocumentError(f"Invalid borders {borders!r}. Use one of: {', '.join(tables.BORDER_PRESETS)}.")
        padding = None
        if cell_margins_in is not None:
            try:
                padding = round(float(cell_margins_in) * tables.TWIPS_PER_INCH)
            except (TypeError, ValueError):
                padding = -1
            if not 0 <= padding <= tables.TWIPS_PER_INCH:
                raise DocumentError("cell_margins_in is the padding inside the cells in inches, between 0 and 1.")
        grid = tables.build_grid(tbl)
        if alignment is not None and str(alignment).strip().lower() not in _TABLE_ALIGNMENTS:
            raise DocumentError(f"Invalid alignment {alignment!r}. Use 'left', 'center' or 'right'.")
        if header_rows is not None and not 0 <= int(header_rows) <= len(grid.rows):
            raise DocumentError(f"header_rows must be between 0 and {len(grid.rows)}, the number of rows.")
        twips = self._twips(column_widths_in, grid.width) if column_widths_in is not None else None

        if twips is not None or alignment is not None or borders is not None or padding is not None:
            self.edits.before_table_format(tbl)
        if borders is not None:
            tables.set_table_borders(tbl, str(borders).strip().lower())
        if padding is not None:
            tables.set_table_cell_margins(tbl, padding)
        if twips is not None:
            widths = tables.grid_widths(tbl)
            for col, width in enumerate(twips):
                physical = grid.physical_of(col)
                widths.extend([0] * (physical + 1 - len(widths)))
                widths[physical] = width
            tables.set_grid_widths(tbl, widths)
            for tr in grid.rows:
                for tc in tables.row_cells(tr):
                    width = tables.width_from_grid(tbl, tc)
                    if width and tables.cell_width(tc) != (width, "dxa"):
                        self.edits.before_cell_format(tc)
                        tables.set_cell_width(tc, width, create=True)
            tables.set_table_width(tbl, sum(twips), fixed=True)
        if alignment is not None:
            tables.set_table_alignment(tbl, _TABLE_ALIGNMENTS[str(alignment).strip().lower()])
        if header_rows is not None:
            for index, tr in enumerate(grid.rows):
                if tables.is_header_row(tr) != (index < int(header_rows)):
                    self.edits.before_row_format(tr)
                    tables.set_header_row(tr, index < int(header_rows))
