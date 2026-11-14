"""Tables as a grid.

A Word table is a list of rows of cells, but cells can span several grid
columns (``w:gridSpan``) and merge down several rows (``w:vMerge``), and rows
and cells can sit inside content controls. This module turns a ``w:tbl`` into
a grid in which every (row, column) position knows the cell that covers it,
so the rest of the document layer never indexes ``w:tc`` elements directly.

Tracked-deleted rows and cells (``w:del`` on the row, ``w:cellDel`` on the
cell) are still in the file but are not part of the grid, like everywhere
else in the document layer.

The helpers that write here only change layout properties (grid columns,
widths, spans). Inserting and removing rows and cells is the job of
``EditContext``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Iterator

from lxml import etree

from .ooxml import (
    CELL_DEL,
    CUSTOM_XML,
    GRID_BEFORE,
    GRID_COL,
    GRID_SPAN,
    H_MERGE,
    P,
    SDT,
    SDT_CONTENT,
    TBL,
    TBL_GRID,
    TBL_HEADER,
    TBLPR_ORDER,
    TC,
    TCPR,
    TCPR_ORDER,
    TR,
    TRPR,
    TRPR_ORDER,
    V_MERGE,
    el,
    ensure_child,
    get_or_add_props,
    is_deleted_block,
    qn,
    remove_child,
)

VAL = qn("w:val")
_W = qn("w:w")
_TYPE = qn("w:type")
_TRANSPARENT = (SDT, SDT_CONTENT, CUSTOM_XML)
TWIPS_PER_INCH = 1440


# ------------------------------------------------------------------- reading

def _through(parent: etree._Element, tag: str) -> Iterator[etree._Element]:
    """Children of ``parent`` with ``tag``, looking through content controls."""
    for child in parent:
        if child.tag == tag:
            yield child
        elif child.tag in _TRANSPARENT:
            yield from _through(child, tag)


def live_rows(tbl: etree._Element) -> list[etree._Element]:
    """Rows that have not been removed by a tracked deletion."""
    return [tr for tr in _through(tbl, TR) if not is_deleted_block(tr)]


def cell_mark(tc: etree._Element, tag: str) -> etree._Element | None:
    """The ``w:cellIns`` or ``w:cellDel`` (``tag``) recorded on a cell."""
    tcpr = tc.find(TCPR)
    return tcpr.find(tag) if tcpr is not None else None


def row_cells(tr: etree._Element, include_deleted: bool = False) -> list[etree._Element]:
    """The cells of a row, without those removed by a tracked deletion."""
    return [tc for tc in _through(tr, TC) if include_deleted or cell_mark(tc, CELL_DEL) is None]


def _int(element: etree._Element | None, attribute: str = VAL, default: int = 0) -> int:
    try:
        return int(float(element.get(attribute))) if element is not None else default
    except (TypeError, ValueError):
        return default


def span(tc: etree._Element) -> int:
    """Number of grid columns a cell covers."""
    tcpr = tc.find(TCPR)
    return max(_int(tcpr.find(GRID_SPAN) if tcpr is not None else None, default=1), 1)


def v_merge(tc: etree._Element) -> str | None:
    """``"restart"`` for the top cell of a vertical merge, ``"continue"`` for
    the cells under it, None for an ordinary cell."""
    tcpr = tc.find(TCPR)
    merge = tcpr.find(V_MERGE) if tcpr is not None else None
    if merge is None:
        return None
    return "restart" if merge.get(VAL) == "restart" else "continue"


def h_merge(tc: etree._Element) -> str | None:
    """The legacy horizontal merge mark (``w:hMerge``) on a cell: "restart"
    for the first cell, "continue" for those it swallows, None for none."""
    tcpr = tc.find(TCPR)
    merge = tcpr.find(H_MERGE) if tcpr is not None else None
    if merge is None:
        return None
    return "restart" if merge.get(VAL) == "restart" else "continue"


def h_merge_runs(tbl: etree._Element) -> list[tuple[etree._Element | None, list[etree._Element]]]:
    """Every legacy horizontal merge in the table's own rows as ``(first
    cell, the cells it swallows)``. A stray mark with no partner comes back
    as ``(cell, [])`` or ``(None, [cell])``."""
    runs: list[tuple[etree._Element | None, list[etree._Element]]] = []
    for tr in _through(tbl, TR):
        open_run = None
        for tc in row_cells(tr, include_deleted=True):
            state = h_merge(tc)
            if state == "restart":
                open_run = (tc, [])
                runs.append(open_run)
            elif state == "continue":
                if open_run is None:
                    runs.append((None, [tc]))
                else:
                    open_run[1].append(tc)
            else:
                open_run = None
    return runs


def clear_h_merge(tc: etree._Element) -> None:
    if tc.find(TCPR) is not None:
        remove_child(tc.find(TCPR), "w:hMerge")


def grid_before(tr: etree._Element) -> int:
    """Grid columns a row skips before its first cell."""
    trpr = tr.find(TRPR)
    return _int(trpr.find(GRID_BEFORE) if trpr is not None else None)


def physical_col(tc: etree._Element) -> int:
    """The cell's column in ``w:tblGrid``, counting tracked-deleted cells."""
    tr = next(a for a in tc.iterancestors() if a.tag == TR)
    col = grid_before(tr)
    for other in row_cells(tr, include_deleted=True):
        if other is tc:
            break
        col += span(other)
    return col


def physical_width(tbl: etree._Element) -> int:
    """Grid columns the widest row occupies, tracked-deleted content included."""
    return max(
        (grid_before(tr) + sum(span(tc) for tc in row_cells(tr, include_deleted=True)) for tr in _through(tbl, TR)),
        default=0,
    )


@dataclass(eq=False)
class Slot:
    """One cell as the reader sees it: a plain cell or a whole merged area.
    ``tc`` is its top-left cell, which holds the content."""

    tc: etree._Element
    row: int
    col: int
    rowspan: int = 1
    colspan: int = 1
    members: list = field(default_factory=list)  # the w:tc of every row it covers

    @property
    def merged(self) -> bool:
        return self.rowspan > 1 or self.colspan > 1


@dataclass(eq=False)
class Position:
    """What sits at one (row, column) of the grid."""

    slot: Slot
    own: etree._Element  # the w:tc of this row (the anchor, or a vMerge continuation)
    physical: int  # column of ``own`` in w:tblGrid


@dataclass(eq=False)
class Grid:
    rows: list  # live w:tr elements
    positions: list  # per row: {column: Position}
    width: int

    def at(self, row: int, col: int) -> Position | None:
        return self.positions[row].get(col) if 0 <= row < len(self.positions) else None

    def slots(self) -> list[Slot]:
        """Every cell once, in reading order."""
        seen: list[Slot] = []
        for r, row in enumerate(self.positions):
            for c in sorted(row):
                slot = row[c].slot
                if slot.row == r and slot.col == c:
                    seen.append(slot)
        return seen

    def physical_of(self, col: int) -> int | None:
        """The ``w:tblGrid`` column behind grid column ``col``."""
        for row in self.positions:
            position = row.get(col)
            if position is not None:
                return position.physical + col - position.slot.col
        return None

    def row_extent(self, row: int) -> int:
        """One past the last column the row occupies."""
        return max(self.positions[row], default=-1) + 1


def build_grid(tbl: etree._Element) -> Grid:
    rows = live_rows(tbl)
    positions: list[dict[int, Position]] = []
    for r, tr in enumerate(rows):
        row: dict[int, Position] = {}
        col = physical = grid_before(tr)
        merging: Position | None = None  # the open legacy horizontal merge of this row
        for tc in row_cells(tr, include_deleted=True):
            width = span(tc)
            if cell_mark(tc, CELL_DEL) is not None:
                physical += width
                continue
            if h_merge(tc) == "continue" and merging is not None:
                # Written by old versions of Word: the cell is swallowed by
                # the one that opened the merge, as with a gridSpan.
                merging.slot.colspan += width
                for c in range(col, col + width):
                    row[c] = merging
                col += width
                physical += width
                continue
            above = positions[r - 1].get(col) if r > 0 else None
            if v_merge(tc) == "continue" and above is not None and above.slot.col == col and above.slot.colspan == width:
                slot = above.slot
                slot.rowspan = r - slot.row + 1
                slot.members.append(tc)
            else:
                slot = Slot(tc, r, col, 1, width, [tc])
            position = Position(slot, tc, physical)
            for c in range(col, col + width):
                row[c] = position
            merging = position if h_merge(tc) == "restart" else None
            col += width
            physical += width
        positions.append(row)
    return Grid(rows, positions, max((max(row, default=-1) + 1 for row in positions), default=0))


def _enclosing_cell(element: etree._Element) -> etree._Element | None:
    return next((a for a in element.iterancestors() if a.tag == TC), None)


def cell_paragraphs(tc: etree._Element) -> list[etree._Element]:
    """The live paragraphs of a cell, not counting those of nested tables."""
    return [p for p in tc.iter(P) if not is_deleted_block(p) and _enclosing_cell(p) is tc]


def nested_tables(tc: etree._Element) -> list[etree._Element]:
    """Tables placed directly in a cell that have not been tracked-deleted."""
    return [t for t in tc.iter(TBL) if _enclosing_cell(t) is tc and not is_deleted_block(t)]


def is_header_row(tr: etree._Element) -> bool:
    trpr = tr.find(TRPR)
    flag = trpr.find(TBL_HEADER) if trpr is not None else None
    return flag is not None and flag.get(VAL) not in ("0", "false", "off")


# -------------------------------------------------------------------- layout

def grid_cols(tbl: etree._Element) -> list[etree._Element]:
    grid = tbl.find(TBL_GRID)
    return grid.findall(GRID_COL) if grid is not None else []


def grid_widths(tbl: etree._Element) -> list[int]:
    """Column widths in twips, as stored in ``w:tblGrid``."""
    return [_int(col, _W) for col in grid_cols(tbl)]


def insert_grid_col(tbl: etree._Element, index: int, width: int) -> None:
    grid = tbl.find(TBL_GRID)
    if grid is None:
        return
    cols = grid.findall(GRID_COL)
    new = el("w:gridCol", {"w:w": str(width)})
    if index < len(cols):
        cols[index].addprevious(new)
    elif cols:
        cols[-1].addnext(new)
    else:
        grid.insert(0, new)


def remove_grid_col(tbl: etree._Element, index: int) -> None:
    cols = grid_cols(tbl)
    if 0 <= index < len(cols):
        cols[index].getparent().remove(cols[index])


def set_grid_widths(tbl: etree._Element, widths: list[int]) -> None:
    """Write the widths of the first ``len(widths)`` grid columns, adding
    columns (and the grid itself) when they are missing."""
    grid = tbl.find(TBL_GRID)
    if grid is None:
        grid = el("w:tblGrid")
        get_or_add_props(tbl, "w:tblPr").addnext(grid)
    cols = grid.findall(GRID_COL)
    for index, width in enumerate(widths):
        if index < len(cols):
            cols[index].set(_W, str(width))
        else:
            insert_grid_col(tbl, index, width)


def scale_grid(tbl: etree._Element, total: int) -> None:
    """Scale the grid columns so they add up to ``total`` twips."""
    widths = grid_widths(tbl)
    current = sum(widths)
    if total <= 0 or current <= 0 or current == total:
        return
    scaled = [max(round(w * total / current), 1) for w in widths]
    scaled[-1] += total - sum(scaled)
    set_grid_widths(tbl, scaled)


def cell_width(tc: etree._Element) -> tuple[int, str] | None:
    """A cell's preferred width and its unit (``dxa``, ``pct``, ``auto``)."""
    tcpr = tc.find(TCPR)
    tcw = tcpr.find(qn("w:tcW")) if tcpr is not None else None
    return (_int(tcw, _W), tcw.get(_TYPE) or "dxa") if tcw is not None else None


def set_cell_width(tc: etree._Element, width: int, create: bool = False) -> bool:
    """Set a cell's width in twips. Widths given in other units are left
    alone unless ``create`` is true. Returns whether anything changed."""
    current = cell_width(tc)
    if current == (width, "dxa") or (not create and (current is None or current[1] != "dxa")):
        return False
    tcw = ensure_child(_tcpr(tc), "w:tcW", TCPR_ORDER)
    tcw.set(_W, str(width))
    tcw.set(_TYPE, "dxa")
    return True


def width_from_grid(tbl: etree._Element, tc: etree._Element) -> int | None:
    """The width the grid gives a cell (the sum over the columns it spans)."""
    widths = grid_widths(tbl)
    start = physical_col(tc)
    covered = widths[start:start + span(tc)]
    return sum(covered) if len(covered) == span(tc) and all(covered) else None


def _tcpr(tc: etree._Element) -> etree._Element:
    return get_or_add_props(tc, "w:tcPr")


def set_span(tc: etree._Element, columns: int) -> None:
    if columns > 1:
        ensure_child(_tcpr(tc), "w:gridSpan", TCPR_ORDER).set(VAL, str(columns))
    elif tc.find(TCPR) is not None:
        remove_child(tc.find(TCPR), "w:gridSpan")


def set_v_merge(tc: etree._Element, state: str | None) -> None:
    """``state`` is "restart", "continue" or None (not merged)."""
    if state is None:
        if tc.find(TCPR) is not None:
            remove_child(tc.find(TCPR), "w:vMerge")
        return
    merge = ensure_child(_tcpr(tc), "w:vMerge", TCPR_ORDER)
    if state == "restart":
        merge.set(VAL, "restart")
    else:
        merge.attrib.pop(VAL, None)


def set_grid_before(tr: etree._Element, columns: int) -> None:
    trpr = tr.find(TRPR)
    flag = trpr.find(GRID_BEFORE) if trpr is not None else None
    if flag is None:
        return
    if columns > 0:
        flag.set(VAL, str(columns))
    else:
        trpr.remove(flag)


def set_header_row(tr: etree._Element, on: bool) -> None:
    """Mark a row to repeat at the top of every page the table runs onto."""
    if on:
        ensure_child(get_or_add_props(tr, "w:trPr"), "w:tblHeader", TRPR_ORDER).attrib.pop(VAL, None)
    elif tr.find(TRPR) is not None:
        remove_child(tr.find(TRPR), "w:tblHeader")


def set_table_width(tbl: etree._Element, width: int, fixed: bool = False) -> None:
    """Set the table's width in twips; ``fixed`` stops Word resizing columns to fit content."""
    tblpr = get_or_add_props(tbl, "w:tblPr")
    tblw = ensure_child(tblpr, "w:tblW", TBLPR_ORDER)
    tblw.set(_W, str(width))
    tblw.set(_TYPE, "dxa")
    if fixed:
        ensure_child(tblpr, "w:tblLayout", TBLPR_ORDER).set(_TYPE, "fixed")


def table_width(tbl: etree._Element) -> tuple[int, str] | None:
    tblpr = tbl.find(qn("w:tblPr"))
    tblw = tblpr.find(qn("w:tblW")) if tblpr is not None else None
    return (_int(tblw, _W), tblw.get(_TYPE) or "dxa") if tblw is not None else None


def table_alignment(tbl: etree._Element) -> str | None:
    tblpr = tbl.find(qn("w:tblPr"))
    jc = tblpr.find(qn("w:jc")) if tblpr is not None else None
    return {"start": "left", "end": "right"}.get(jc.get(VAL), jc.get(VAL)) if jc is not None else None


def set_table_alignment(tbl: etree._Element, alignment: str) -> None:
    ensure_child(get_or_add_props(tbl, "w:tblPr"), "w:jc", TBLPR_ORDER).set(VAL, alignment)


# ------------------------------------------------- borders, padding, direction

SIDES = ("top", "left", "bottom", "right")
OPPOSITE = {"top": "bottom", "bottom": "top", "left": "right", "right": "left"}
BORDER_STYLES = {"none": "nil", "single": "single", "double": "double", "dashed": "dashed",
                 "dotted": "dotted", "thick": "thick"}
_BORDER_NAMES = {value: key for key, value in BORDER_STYLES.items()} | {"none": "none"}
TEXT_DIRECTIONS = {"horizontal": None, "vertical": "tbRl", "vertical_up": "btLr"}
_DIRECTION_NAMES = {"tbRl": "vertical", "tbRlV": "vertical", "btLr": "vertical_up", "lrTb": "horizontal"}
BORDER_PRESETS = {
    "all": ("top", "left", "bottom", "right", "insideH", "insideV"),
    "none": (),
    "outer": ("top", "left", "bottom", "right"),
    "horizontal": ("top", "bottom", "insideH"),
}
# Each side also has a writing-direction name (start / end) in the schema.
_EDGE_ORDER = ["w:top", "w:start", "w:left", "w:bottom", "w:end", "w:right", "w:insideH", "w:insideV",
               "w:tl2br", "w:tr2bl"]
_LOGICAL = {"left": "start", "right": "end"}
_SZ = qn("w:sz")
_COLOR = qn("w:color")


def _set_edge(holder: etree._Element, side: str, style: str, size: int = 4, color: str = "auto") -> None:
    """One edge of a ``w:tcBorders`` / ``w:tblBorders``. ``style`` is a key of
    BORDER_STYLES; ``size`` is in eighths of a point."""
    if side in _LOGICAL:
        remove_child(holder, "w:" + _LOGICAL[side])
    edge = ensure_child(holder, "w:" + side, _EDGE_ORDER)
    edge.attrib.clear()
    edge.set(VAL, BORDER_STYLES[style])
    edge.set(_SZ, "0" if style == "none" else str(size))
    edge.set(qn("w:space"), "0")
    edge.set(_COLOR, color)


def _edges(holder: etree._Element | None, sides: tuple) -> dict:
    out: dict = {}
    for side in sides:
        edge = None
        if holder is not None:
            edge = holder.find(qn("w:" + side))
            if edge is None and side in _LOGICAL:
                edge = holder.find(qn("w:" + _LOGICAL[side]))
        if edge is None:
            continue
        style = _BORDER_NAMES.get(edge.get(VAL, ""), edge.get(VAL, ""))
        entry: dict = {"style": style}
        if style != "none":
            entry["width_pt"] = _int(edge, _SZ) / 8
            if edge.get(_COLOR) not in (None, "auto"):
                entry["color"] = edge.get(_COLOR).upper()
        out[side] = entry
    return out


def set_cell_border(tc: etree._Element, side: str, style: str, size: int = 4, color: str = "auto") -> None:
    _set_edge(ensure_child(_tcpr(tc), "w:tcBorders", TCPR_ORDER), side, style, size, color)


def _set_margins(holder: etree._Element, twips: int) -> None:
    for side in SIDES:
        if side in _LOGICAL:
            remove_child(holder, "w:" + _LOGICAL[side])
        edge = ensure_child(holder, "w:" + side, _EDGE_ORDER)
        edge.set(_W, str(twips))
        edge.set(_TYPE, "dxa")


def _margins(holder: etree._Element | None) -> float | None:
    """Padding in inches when all four sides are set to one value."""
    if holder is None:
        return None
    values = set()
    for side in SIDES:
        edge = holder.find(qn("w:" + side))
        if edge is None and side in _LOGICAL:
            edge = holder.find(qn("w:" + _LOGICAL[side]))
        if edge is None or edge.get(_TYPE, "dxa") != "dxa":
            return None
        values.add(_int(edge, _W))
    return round(values.pop() / TWIPS_PER_INCH, 3) if len(values) == 1 else None


def set_cell_margins(tc: etree._Element, twips: int) -> None:
    """Padding between the cell's edges and its text, the same on all sides."""
    _set_margins(ensure_child(_tcpr(tc), "w:tcMar", TCPR_ORDER), twips)


def set_text_direction(tc: etree._Element, direction: str) -> None:
    """``direction`` is a key of TEXT_DIRECTIONS."""
    value = TEXT_DIRECTIONS[direction]
    if value is None:
        if tc.find(TCPR) is not None:
            remove_child(tc.find(TCPR), "w:textDirection")
    else:
        ensure_child(_tcpr(tc), "w:textDirection", TCPR_ORDER).set(VAL, value)


def set_table_borders(tbl: etree._Element, preset: str) -> None:
    """Draw the table's lines: a key of BORDER_PRESETS. Lines set on single
    cells still win over these."""
    borders = ensure_child(get_or_add_props(tbl, "w:tblPr"), "w:tblBorders", TBLPR_ORDER)
    for side in BORDER_PRESETS["all"]:
        _set_edge(borders, side, "single" if side in BORDER_PRESETS[preset] else "none")


def table_borders(tbl: etree._Element) -> str | None:
    """The preset the table's own borders match, if they match one."""
    tblpr = tbl.find(qn("w:tblPr"))
    borders = tblpr.find(qn("w:tblBorders")) if tblpr is not None else None
    if borders is None:
        return None
    drawn = tuple(side for side, edge in _edges(borders, BORDER_PRESETS["all"]).items() if edge["style"] != "none")
    return next((name for name, sides in BORDER_PRESETS.items() if set(sides) == set(drawn)), None)


def set_table_cell_margins(tbl: etree._Element, twips: int) -> None:
    _set_margins(ensure_child(get_or_add_props(tbl, "w:tblPr"), "w:tblCellMar", TBLPR_ORDER), twips)


def table_cell_margins(tbl: etree._Element) -> float | None:
    tblpr = tbl.find(qn("w:tblPr"))
    return _margins(tblpr.find(qn("w:tblCellMar")) if tblpr is not None else None)


def apply_cell_format(tc: etree._Element, fill: str | None = None, vertical_align: str | None = None) -> None:
    """``fill`` is a 6-digit hex color, or "" to remove the shading."""
    tcpr = _tcpr(tc)
    if fill:
        shd = ensure_child(tcpr, "w:shd", TCPR_ORDER)
        shd.set(VAL, "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), fill)
    elif fill is not None:
        remove_child(tcpr, "w:shd")
    if vertical_align:
        ensure_child(tcpr, "w:vAlign", TCPR_ORDER).set(VAL, vertical_align)


def cell_format(tc: etree._Element) -> dict:
    """Direct cell formatting: ``shading`` (hex), ``vertical_align``,
    ``borders`` (per side: style, width_pt, color), ``margins_in`` (when the
    four sides agree) and ``text_direction`` (when not horizontal)."""
    out: dict = {}
    tcpr = tc.find(TCPR)
    if tcpr is None:
        return out
    shd = tcpr.find(qn("w:shd"))
    fill = shd.get(qn("w:fill")) if shd is not None else None
    if fill and fill != "auto":
        out["shading"] = fill.upper()
    valign = tcpr.find(qn("w:vAlign"))
    if valign is not None and valign.get(VAL):
        out["vertical_align"] = valign.get(VAL)
    borders = _edges(tcpr.find(qn("w:tcBorders")), SIDES)
    if borders:
        out["borders"] = borders
    margins = _margins(tcpr.find(qn("w:tcMar")))
    if margins is not None:
        out["margins_in"] = margins
    direction = tcpr.find(qn("w:textDirection"))
    if direction is not None and _DIRECTION_NAMES.get(direction.get(VAL, ""), "horizontal") != "horizontal":
        out["text_direction"] = _DIRECTION_NAMES[direction.get(VAL)]
    return out
