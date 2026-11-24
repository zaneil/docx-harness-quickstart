"""WordDocument: block-level operations over a WordPackage.

The document body is exposed as an ordered list of *blocks* (paragraphs,
tables and content controls). Each block gets a stable id (``b1``, ``b2``...)
when the document is loaded or the block is created. Ids live in memory only
and never shift when other blocks are inserted or deleted.

A cell of a table is addressed as ``"<table id>[row,col]"``, for example
``"b7[0,1]"`` (zero-based). Rows and columns are positions on the table's
grid, so a merged cell answers to every position it covers. ``"b7[0,*]"`` is
a whole row and ``"b7[*,1]"`` a whole column. A table nested in a cell gets a
block id of its own the first time it is read.
"""

from __future__ import annotations

import difflib
import os
import re
from pathlib import Path

from lxml import etree

from . import builtin_styles, numbering, styling, tables
from .common import (
    BLOCK_TAGS,
    PAGE_WIDTH_TWIPS,
    DocumentError,
    _lines,
    _live_paragraphs,
    _truncate,
    make_run_text,
)
from .edits import EditContext
from .lists import ListsMixin
from .media import MediaMixin
from .notes import NotesMixin
from .ooxml import (
    BODY,
    BR,
    CR,
    INS,
    MOVE_TO,
    RPR,
    SDT,
    SECTPR,
    SETTINGS_BEFORE_TRACK_REVISIONS,
    TAB,
    TBL,
    TRACK_REVISIONS,
    W_NS,
    P,
    R,
    T,
    el,
    insert_after_any,
    is_deleted_block,
    isolate_range,
    iter_runs,
    make_paragraph,
    make_run,
    paragraph_text,
    qn,
    run_text,
)
from .package import WordPackage
from .revisions import RevisionsMixin
from .stories import StoriesMixin, _ends_section
from .table_ops import TablesMixin

_TOKEN = re.compile(r"\s+|\w+|[^\w\s]")
_REF = re.compile(r"^(b\d+)(?:\[\s*(\d+|\*)\s*,\s*(\d+|\*)\s*\])?$")
_TEXT_CHILDREN = (RPR, T, TAB, BR, CR)


class WordDocument(StoriesMixin, NotesMixin, ListsMixin, TablesMixin, MediaMixin, RevisionsMixin):
    """``track_changes=True`` records every edit as a Word revision by
    ``author`` (a redline the reader can accept or reject) instead of changing
    the text directly."""

    def __init__(self, package: WordPackage, *, track_changes: bool = False, author: str | None = None):
        self.package = package
        self.edits = EditContext(track_changes=track_changes, author=author or "docx-harness")
        self._counter = 0
        self._load()
        if track_changes:
            self.track_changes = True

    @property
    def track_changes(self) -> bool:
        return self.edits.track_changes

    @track_changes.setter
    def track_changes(self, on: bool) -> None:
        """Like the Track Changes switch in Word: edits from now on are
        recorded (or not), and the saved file tells Word to keep tracking."""
        self.edits.track_changes = bool(on)
        part = self.package.settings_part
        if part is None:
            return  # nowhere to store the switch; edits are still recorded
        settings = self.package.xml(part)
        flag = settings.find(TRACK_REVISIONS)
        if not on:
            if flag is not None:
                settings.remove(flag)
        elif flag is None:
            insert_after_any(settings, el("w:trackRevisions"), SETTINGS_BEFORE_TRACK_REVISIONS)
        else:
            flag.attrib.pop(qn("w:val"), None)  # an explicit "false" becomes on

    # ------------------------------------------------------------- lifecycle

    @classmethod
    def open(cls, path: str | os.PathLike | None = None, **kwargs) -> "WordDocument":
        """Open a .docx file, or a new document from the bundled template."""
        return cls(WordPackage.open(path), **kwargs)

    def save(self, path: str | os.PathLike) -> Path:
        return self.package.save(path)

    def _load(self) -> None:
        root = self.package.xml(self.package.main_document_part)
        body = root.find(BODY)
        if body is None:
            raise DocumentError("Document has no body")
        self._body = body
        self._bind()
        self._ids: dict[str, etree._Element] = {}
        self._by_element: dict[etree._Element, str] = {}
        for element in self._block_elements():
            self._id_of(element)

    def _bind(self) -> None:
        """Tell the edit context which parts hold content (for revision ids)."""
        self.edits.bind(self._body.getroottree().getroot(), *(story for _, _, story in self._story_roots()))

    @property
    def _styles(self) -> etree._Element | None:
        part = self.package.styles_part
        return self.package.xml(part) if part else None

    def _part_beside_main(self, filename: str) -> str:
        folder = self.package.main_document_part.rpartition("/")[0]
        return f"{folder}/{filename}" if folder else filename

    def _styles_root(self) -> etree._Element:
        """The styles part, created with the default styles when the file has none."""
        if self.package.styles_part is None:
            xml = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="{W_NS}">'
                + builtin_styles.DOC_DEFAULTS + "".join(style.xml() for style in builtin_styles.BASE) + "</w:styles>"
            )
            name = self.package.add_part(
                self._part_beside_main("styles.xml"), xml,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            )
            self.package.relate(self.package.main_document_part, "styles", name)
        return self.package.xml(self.package.styles_part)

    @property
    def _numbering(self) -> etree._Element | None:
        part = self.package.numbering_part
        return self.package.xml(part) if part else None

    def _numbering_root(self) -> etree._Element:
        """The numbering part, created empty when the file has none."""
        if self.package.numbering_part is None:
            name = self.package.add_part(self._part_beside_main("numbering.xml"), numbering.EMPTY_PART,
                                         numbering.CONTENT_TYPE)
            self.package.relate(self.package.main_document_part, "numbering", name)
        return self.package.xml(self.package.numbering_part)

    # -------------------------------------------------------------- block ids

    def _block_elements(self) -> list[etree._Element]:
        # Blocks removed by a tracked deletion are still in the file but are
        # not part of the document as it reads now.
        return [child for child in self._body if child.tag in BLOCK_TAGS and not is_deleted_block(child)]

    def _id_of(self, element: etree._Element) -> str:
        block_id = self._by_element.get(element)
        if block_id is None:
            self._counter += 1
            block_id = f"b{self._counter}"
            self._ids[block_id] = element
            self._by_element[element] = block_id
        return block_id

    def _get(self, block_id: str) -> etree._Element:
        element = self._ids.get(str(block_id).strip())
        cell = _REF.match(str(block_id).strip())
        if cell is not None and cell.group(2) is not None:
            raise DocumentError(
                f"{block_id!r} is a table cell, not a block. Set its text with set_table_cell, "
                f"or change part of it with replace_text limited to that cell."
            )
        if element is None or not self._attached(element):
            raise DocumentError(
                f"No block with id {block_id!r}. Use the outline to see current block ids."
            )
        return element

    def _attached(self, element: etree._Element) -> bool:
        return any(ancestor is self._body for ancestor in element.iterancestors())

    def _body_block(self, block_id: str) -> etree._Element:
        """A block of the body itself, for operations on the sequence of
        blocks. A table nested in a cell has an id but no place in it."""
        element = self._get(block_id)
        if element.getparent() is not self._body:
            raise DocumentError(
                f"Block {block_id} is a table nested inside a cell of another table, so it cannot be moved "
                "or used as a position. Use the id of a block from the outline instead."
            )
        return element

    def _parse_ref(self, ref: str) -> tuple[etree._Element, tuple[int | None, int | None] | None]:
        """A block and, for a cell reference, its (row, column); None in
        either place stands for ``*``, every row or every column."""
        match = _REF.match(str(ref).strip())
        if not match:
            raise DocumentError(
                f"Invalid block reference {ref!r}. Use a block id like 'b3', a table cell like 'b3[0,1]', "
                "a table row like 'b3[0,*]' or a table column like 'b3[*,1]'."
            )
        element = self._get(match.group(1))
        if match.group(2) is None:
            return element, None
        row, col = (None if g == "*" else int(g) for g in match.group(2, 3))
        return element, (row, col)

    def _paragraphs_for(self, ref: str) -> list[etree._Element]:
        for part, kind, root in self._story_roots():
            if part == str(ref).strip():
                return self._story_paragraphs(kind, root)
        element, cell = self._parse_ref(ref)
        if cell is not None:
            return [p for slot in self._slots(element, *cell) for p in tables.cell_paragraphs(slot.tc)]
        return _live_paragraphs(element)

    def _paragraph(self, block_id: str) -> etree._Element:
        element = self._get(block_id)
        if element.tag != P:
            raise DocumentError(f"Block {block_id} is a {self._type(element)}, not a paragraph.")
        return element

    def _table(self, block_id: str) -> etree._Element:
        element = self._get(block_id)
        if element.tag != TBL:
            raise DocumentError(f"Block {block_id} is a {self._type(element)}, not a table.")
        return element

    @staticmethod
    def _type(element: etree._Element) -> str:
        return {P: "paragraph", TBL: "table", SDT: "content_control"}[element.tag]

    # ------------------------------------------------------------------ read

    @staticmethod
    def _cell_text(tc: etree._Element) -> str:
        """A cell's text, one line per paragraph. Tables nested in the cell
        are read through their own ids."""
        return "\n".join(paragraph_text(p) for p in tables.cell_paragraphs(tc))

    @staticmethod
    def _live_rows(tbl: etree._Element) -> list[etree._Element]:
        """Rows that have not been removed by a tracked deletion."""
        return tables.live_rows(tbl)

    def _rows(self, tbl: etree._Element, grid: tables.Grid | None = None) -> list[list[str]]:
        """The table's text on its grid: every row has one entry per column,
        and a position covered by a merged cell (other than its top-left) is ""."""
        grid = grid or tables.build_grid(tbl)
        texts = {slot: self._cell_text(slot.tc) for slot in grid.slots()}
        out = []
        for r, row in enumerate(grid.positions):
            cells = [""] * grid.width
            for c, position in row.items():
                if (position.slot.row, position.slot.col) == (r, c):
                    cells[c] = texts[position.slot]
            out.append(cells)
        return out

    def _describe_table(self, tbl: etree._Element, include_formatting: bool) -> dict:
        block_id = self._id_of(tbl)
        grid = tables.build_grid(tbl)
        info: dict = {"rows": self._rows(tbl, grid)}
        merged, nested, cells = [], [], []
        for slot in grid.slots():
            ref = f"{block_id}[{slot.row},{slot.col}]"
            if slot.merged:
                merged.append({"cell": ref, "rows": slot.rowspan, "cols": slot.colspan})
            inner = tables.nested_tables(slot.tc)
            if inner:
                nested.append({"cell": ref, "tables": [self._id_of(t) for t in inner]})
            look = tables.cell_format(slot.tc) if include_formatting else {}
            if look:
                cells.append({"cell": ref, **look})
        if merged:
            info["merged"] = merged
        if nested:
            info["nested"] = nested
        if include_formatting:
            layout: dict = {}
            widths = tables.grid_widths(tbl)
            columns = [grid.physical_of(c) for c in range(grid.width)]
            if columns and all(c is not None and c < len(widths) and widths[c] for c in columns):
                layout["column_widths_in"] = [round(widths[c] / tables.TWIPS_PER_INCH, 2) for c in columns]
            header_rows = next((i for i, tr in enumerate(grid.rows) if not tables.is_header_row(tr)), len(grid.rows))
            if header_rows:
                layout["header_rows"] = header_rows
            if tables.table_alignment(tbl):
                layout["alignment"] = tables.table_alignment(tbl)
            if tables.table_borders(tbl):
                layout["borders"] = tables.table_borders(tbl)
            if tables.table_cell_margins(tbl) is not None:
                layout["cell_margins_in"] = tables.table_cell_margins(tbl)
            if cells:
                layout["cells"] = cells
            if layout:
                info["format"] = layout
        return info

    def _list_numbers(self) -> dict[etree._Element, int]:
        """The number each numbered list item of the body shows, counted the
        way Word does: per list and level, deeper levels starting over under
        each new item above them."""
        counters: dict[int, dict[int, int]] = {}
        shown: dict[etree._Element, int] = {}
        for block in self._block_elements():
            item = self._list_item(block)
            if item is None:
                continue
            num_id, level, kind = item
            levels = counters.setdefault(num_id, {})
            for deeper in [l for l in levels if l > level]:
                del levels[deeper]
            levels[level] = levels[level] + 1 if level in levels else numbering.start_of(self._numbering, num_id, level)
            if kind == "number":
                shown[block] = levels[level]
        return shown

    def _describe(self, element: etree._Element, include_formatting: bool = False, resolve_styles: bool = False,
                  numbers: dict | None = None) -> dict:
        info: dict = {"id": self._id_of(element), "type": self._type(element)}
        if element.tag == P:
            info["style"] = styling.style_name(self._styles, styling.paragraph_style_id(element))
            info["text"] = paragraph_text(element)
            item = self._list_item(element)
            if item is not None:
                info["list"] = {"kind": item[2], "level": item[1]}
                if numbers and element in numbers:
                    info["list"]["number"] = numbers[element]
            links = self._links(element)
            if links:
                info["links"] = links
            pictures = self._images(element)
            if pictures:
                info["images"] = pictures
            notes = self._note_references(element)
            if notes:
                info["notes"] = notes
            if _ends_section(element) and element.getparent() is self._body:
                ended = 1 + sum(1 for sibling in element.itersiblings(preceding=True)
                                if _ends_section(sibling) and not is_deleted_block(sibling))
                info["section_break"] = {"ends_section": ended}
            if include_formatting:
                layout = styling.paragraph_format(element)
                if layout:
                    info["format"] = layout
                if resolve_styles:
                    info["style_format"] = styling.style_format(self._styles, styling.paragraph_style_id(element))
                info["runs"] = [
                    {"text": run_text(r), **styling.run_format(r, explicit_off=resolve_styles),
                     **({"inserted": True} if r.getparent().tag in (INS, MOVE_TO) else {})}
                    for r in iter_runs(element) if run_text(r)
                ]
        elif element.tag == TBL:
            info.update(self._describe_table(element, include_formatting))
        else:
            info["text"] = "\n".join(paragraph_text(p) for p in _live_paragraphs(element))
        return info

    def __len__(self) -> int:
        return len(self._block_elements())

    def block_ids(self) -> list[str]:
        """The ids of the body's blocks, in document order."""
        return [self._id_of(element) for element in self._block_elements()]

    def has_tables(self) -> bool:
        """Whether the body holds at least one table."""
        return any(element.tag == TBL for element in self._block_elements())

    def outline(self, start: str | None = None, limit: int = 200, preview_chars: int = 80) -> list[dict]:
        """One short entry per block: id, type, style and a text preview."""
        out = []
        for element in self._slice(start, limit):
            info = self._describe(element)
            if element.tag == TBL:
                rows = info["rows"]
                cols = max((len(r) for r in rows), default=0)
                preview = f"{len(rows)}x{cols} table: " + " | ".join(rows[0] if rows else [])
                entry = {"id": info["id"], "type": "table", "preview": preview}
            else:
                entry = {"id": info["id"], "type": info["type"], "preview": info["text"]}
                for picture in info.get("images", []):
                    entry["preview"] = f"{entry['preview']} [image: {picture['alt'] or picture['name']}]".strip()
                for note in info.get("notes", []):
                    entry["preview"] = f"{entry['preview']} [{note['kind']} {note['id']}]"
                if "section_break" in info:
                    entry["preview"] = f"{entry['preview']} [end of section {info['section_break']['ends_section']}]".strip()
                for key in ("style", "list"):
                    if key in info:
                        entry[key] = info[key]
            entry["preview"] = _truncate(entry["preview"].replace("\n", " "), preview_chars)
            out.append(entry)
        return out

    def _slice(self, start: str | None, count: int) -> list[etree._Element]:
        elements = self._block_elements()
        index = elements.index(self._body_block(start)) if start else 0
        return elements[index:index + max(count, 0)]

    def read(self, start: str | None = None, count: int = 20, include_formatting: bool = False,
             resolve_styles: bool = False) -> dict:
        """Full content of ``count`` blocks beginning at ``start`` (inclusive).

        ``include_formatting`` adds each paragraph's direct formatting:
        ``format`` (alignment, spacing, indents; left out when there is none)
        and ``runs``. ``resolve_styles`` also adds ``style_format``, what the
        paragraph's style supplies underneath, and the ``number`` each
        numbered list item shows, for rendering a preview.

        A paragraph also returns ``list`` (kind and nesting level) when it is
        a list item, ``links``, ``images`` and ``notes`` (footnote and endnote
        marks with the note text) when it has any, and ``section_break``
        when it ends a section.

        A table returns ``rows`` (its text on the column grid), ``merged``
        (cells covering more than one position) and ``nested`` (ids of tables
        inside its cells) when it has any, and with ``include_formatting`` a
        ``format`` with column widths, header rows, alignment and cell shading.
        ``start`` may also be the id of a nested table, which is then read alone."""
        elements = self._block_elements()
        if start and self._get(start).getparent() is not self._body:
            return {"blocks": [self._describe(self._get(start), include_formatting, resolve_styles)][:max(count, 0)],
                    "next": None, "total_blocks": len(elements)}
        index = elements.index(self._get(start)) if start else 0
        chunk = elements[index:index + max(count, 0)]
        following = elements[index + len(chunk):index + len(chunk) + 1]
        numbers = self._list_numbers() if resolve_styles else None
        return {
            "blocks": [self._describe(e, include_formatting, resolve_styles, numbers) for e in chunk],
            "next": self._id_of(following[0]) if following else None,
            "total_blocks": len(elements),
        }

    def text(self) -> str:
        """Plain text of the whole body, one paragraph per line."""
        return "\n".join(paragraph_text(p) for p in _live_paragraphs(self._body))

    def search(self, query: str, case_sensitive: bool = False, limit: int = 50) -> list[dict]:
        """Find blocks (or table cells) containing ``query``. Headers, footers
        and notes are searched too; their hits have the part name as ``ref``."""
        if not query:
            raise DocumentError("query must not be empty")
        pattern = re.compile(re.escape(query), 0 if case_sensitive else re.IGNORECASE)
        hits: list[dict] = []

        def check(ref: str, text: str) -> None:
            matches = list(pattern.finditer(text))
            if matches and len(hits) < limit:
                first = matches[0]
                lo, hi = max(first.start() - 40, 0), min(first.end() + 40, len(text))
                hits.append({"ref": ref, "matches": len(matches), "snippet": text[lo:hi].replace("\n", " ")})

        def scan(tbl: etree._Element) -> None:
            block_id = self._id_of(tbl)
            for slot in tables.build_grid(tbl).slots():
                check(f"{block_id}[{slot.row},{slot.col}]", self._cell_text(slot.tc))
                for inner in tables.nested_tables(slot.tc):
                    scan(inner)

        for element in self._block_elements():
            if element.tag == TBL:
                scan(element)
            else:
                check(self._id_of(element), "\n".join(paragraph_text(p) for p in _live_paragraphs(element)))
        for story in self.stories():
            check(story["part"], "\n".join(story["paragraphs"]))
        return hits

    # ----------------------------------------------------------------- write

    def _index_after(self, after: str) -> int:
        """Position in the body for content placed after ``after``."""
        after = str(after).strip()
        if after == "end":
            last = self._body[-1] if len(self._body) else None
            return len(self._body) - 1 if last is not None and last.tag == SECTPR else len(self._body)
        if after == "start":
            return 0
        return self._body.index(self._body_block(after)) + 1

    def _insert(self, new: etree._Element, after: str) -> str:
        self.edits.insert_block(self._body, self._index_after(after), new)
        return self._id_of(new)

    def move_block(self, block_id: str, after: str | None = None, *, before: str | None = None) -> None:
        """Move a block to after block ``after`` ("start", "end" or a block
        id), or to just before block ``before``; give one of the two. The block
        keeps its id. When changes are tracked, a paragraph is recorded as a
        move and a table as a deletion plus an insertion."""
        if (after is None) == (before is None):
            raise DocumentError("Give exactly one of after and before.")
        element = self._body_block(block_id)
        if before is not None:
            index = self._body.index(self._body_block(before))
        else:
            index = self._index_after(after)
        moved = self.edits.move_block(element, self._body, index)
        if moved is not element:  # tracked: the id follows the content to its new place
            block_id = self._by_element.pop(element)
            self._ids[block_id] = moved
            self._by_element[moved] = block_id

    def _style_id(self, style: str | None, style_type: str = "paragraph") -> str | None:
        """The id of a style by name or id. A well-known style the file does
        not define (Heading 4, Caption, List Number...) is added to it."""
        if not style:
            return None
        return styling.ensure_style(
            self._styles_root(), style, style_type,
            lambda kind: numbering.base_num(self._numbering_root(), kind),
        )

    def insert_paragraph(self, text: str = "", style: str | None = None, after: str = "end",
                         level: int | None = None) -> str:
        """Insert a paragraph after block ``after`` ("start", "end" or a block id).
        Returns the new block id. Newlines in ``text`` become line breaks.

        With a list style, ``level`` nests the item (0 is the outer level). A
        numbered item carries on the numbered list it is placed in or next
        to; otherwise it starts a new list at 1."""
        p = make_paragraph(text, self._style_id(style))
        index = self._index_after(after)
        self._place_in_list(p, index, level)
        self.edits.insert_block(self._body, index, p)
        return self._id_of(p)

    def insert_page_break(self, after: str = "end") -> str:
        p = el("w:p", None, el("w:r", None, el("w:br", {"w:type": "page"})))
        return self._insert(p, after)

    def insert_table(
        self,
        rows: list[list[str]],
        after: str = "end",
        style: str | None = "Table Grid",
        header: bool = True,
        column_widths_in: list[float] | None = None,
    ) -> str:
        """Insert a table. ``rows`` is a list of rows, each a list of cell
        texts; a newline in a cell's text starts a new paragraph. When
        ``header`` is true the first row is bold and repeats at the top of
        each page. ``column_widths_in`` fixes the column widths (inches);
        otherwise the columns share the page width."""
        if not rows or not all(isinstance(r, (list, tuple)) for r in rows):
            raise DocumentError("rows must be a non-empty list of lists of cell text")
        cols = max(len(r) for r in rows)
        if cols == 0:
            raise DocumentError("rows must contain at least one cell")
        fixed = column_widths_in is not None
        widths = self._twips(column_widths_in, cols) if fixed else [PAGE_WIDTH_TWIPS // cols] * cols

        tbl_pr = el("w:tblPr")
        style_id = None
        if style:
            try:
                style_id = self._style_id(style, "table")
            except styling.StyleError:
                style_id = None
        if style_id:
            tbl_pr.append(el("w:tblStyle", {"w:val": style_id}))
        tbl_pr.append(el("w:tblW", {"w:w": str(sum(widths)), "w:type": "dxa"} if fixed else {"w:w": "5000", "w:type": "pct"}))
        if not style_id:
            # No usable table style in this document: draw explicit borders.
            borders = el("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                borders.append(el(f"w:{side}", {"w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto"}))
            tbl_pr.append(borders)
        if fixed:
            tbl_pr.append(el("w:tblLayout", {"w:type": "fixed"}))
        tbl_pr.append(el("w:tblLook", {"w:val": "04A0"}))

        tbl = el("w:tbl", None, tbl_pr)
        grid = el("w:tblGrid")
        for width in widths:
            grid.append(el("w:gridCol", {"w:w": str(width)}))
        tbl.append(grid)
        for r_index, row in enumerate(rows):
            tr = el("w:tr")
            if header and r_index == 0:
                tr.append(el("w:trPr", None, el("w:tblHeader")))
            for c_index in range(cols):
                text = str(row[c_index]) if c_index < len(row) and row[c_index] is not None else ""
                tc = el("w:tc", None, el("w:tcPr", None, el("w:tcW", {"w:w": str(widths[c_index]), "w:type": "dxa"})))
                for line in _lines(text):
                    p = make_paragraph(line)
                    if header and r_index == 0:
                        for run in p.iter(R):
                            styling.apply_run_format(run, bold=True)
                    tc.append(p)
                tr.append(tc)
            tbl.append(tr)
        return self._insert(tbl, after)

    @staticmethod
    def _twips(widths_in: list[float], columns: int) -> list[int]:
        if isinstance(widths_in, (str, int, float)) or len(widths_in) != columns:
            raise DocumentError(f"Give one width in inches per column: the table has {columns} columns.")
        try:
            twips = [round(float(w) * tables.TWIPS_PER_INCH) for w in widths_in]
        except (TypeError, ValueError):
            raise DocumentError("Column widths must be numbers, in inches.") from None
        if any(w <= 0 for w in twips):
            raise DocumentError("Column widths must be greater than zero.")
        return twips

    # ---------------------------------------------------------------- update

    def _replace_range(self, p: etree._Element, start: int, end: int, text: str) -> None:
        """Replace characters ``[start, end)`` of a paragraph with ``text``,
        which takes the formatting of the text it replaces or follows."""
        if end > start:
            runs = isolate_range(p, start, end)
            if text:
                new = make_run(text, runs[0].find(RPR))
                if self.edits.tracking:
                    # Redlines read best as old text, then new text.
                    parent = runs[-1].getparent()
                    self.edits.insert_run(parent, parent.index(runs[-1]) + 1, new)
                else:
                    parent = runs[0].getparent()
                    self.edits.insert_run(parent, parent.index(runs[0]), new)
            for run in runs:
                self.edits.delete_run(run)
        elif text:
            before = isolate_range(p, 0, start) if start > 0 else []
            if before:
                anchor, offset = before[-1], 1
            else:
                runs = list(iter_runs(p))
                anchor = next((r for r in runs if run_text(r)), runs[0] if runs else None)
                offset = 0
            if anchor is None:
                self.edits.insert_run(p, len(p), make_run(text))
            else:
                parent = anchor.getparent()
                self.edits.insert_run(parent, parent.index(anchor) + offset, make_run(text, anchor.find(RPR)))

    def _revise_paragraph_text(self, p: etree._Element, text: str) -> None:
        """Change a paragraph's text by editing only the words that differ, so
        a tracked rewrite shows as a few revisions rather than the whole
        paragraph struck out and retyped."""
        old, new = paragraph_text(p), make_run_text(text)
        a, b = _TOKEN.findall(old), _TOKEN.findall(new)
        matcher = difflib.SequenceMatcher(None, a, b, autojunk=False)
        if matcher.ratio() < 0.5:
            self._replace_range(p, 0, len(old), new)  # too different: one clean replacement
            return
        offsets = [0]
        for token in a:
            offsets.append(offsets[-1] + len(token))
        # Right to left, so earlier offsets stay valid as the text changes.
        for tag, i1, i2, j1, j2 in reversed(matcher.get_opcodes()):
            if tag != "equal":
                self._replace_range(p, offsets[i1], offsets[i2], "".join(b[j1:j2]))

    def _set_paragraph_text(self, p: etree._Element, text: str) -> None:
        if self.edits.tracking:
            self._revise_paragraph_text(p, text)
            return
        runs = list(iter_runs(p))
        rpr = next((r.find(RPR) for r in runs if run_text(r) and r.find(RPR) is not None), None)
        for run in runs:
            # A run holding a picture, a field or a note reference is not text: it stays.
            if all(child.tag in _TEXT_CHILDREN for child in run):
                self.edits.delete_run(run)
        if text:
            self.edits.insert_run(p, len(p), make_run(text, rpr))

    def set_paragraph_text(self, block_id: str, text: str) -> None:
        """Replace the whole text of a paragraph, keeping its style and the
        character formatting of its first run."""
        self._set_paragraph_text(self._paragraph(block_id), text)

    def _replace_in_paragraph(self, p: etree._Element, pattern: re.Pattern, replacement: str) -> int:
        count = 0
        pos = 0
        while True:
            match = pattern.search(paragraph_text(p), pos)
            if match is None or match.end() == match.start():
                return count
            self._replace_range(p, match.start(), match.end(), replacement)
            pos = match.start() + len(make_run_text(replacement))
            count += 1

    def replace_text(self, find: str, replace: str, ref: str | None = None, case_sensitive: bool = True) -> int:
        """Replace every occurrence of ``find``. Matches may span differently
        formatted runs; the replacement takes the formatting at the start of
        the match. Scope is the whole document (headers, footers and notes
        included, as with Replace All in Word), or one block, table cell or
        story part when ``ref`` is given. Returns the number of replacements."""
        if not find:
            raise DocumentError("find must not be empty")
        pattern = re.compile(re.escape(find), 0 if case_sensitive else re.IGNORECASE)
        if ref:
            paragraphs = self._paragraphs_for(ref)
        else:
            paragraphs = _live_paragraphs(self._body)
            for _, kind, root in self._story_roots():
                paragraphs += self._story_paragraphs(kind, root)
        return sum(self._replace_in_paragraph(p, pattern, replace) for p in paragraphs)

    # ---------------------------------------------------------------- delete

    def delete_block(self, block_id: str) -> None:
        element = self._get(block_id)
        self.edits.delete_block(element)
        self._ids.pop(self._by_element.pop(element), None)

    # --------------------------------------------------------------- styling

    def list_styles(self, style_type: str | None = None) -> list[dict]:
        styles = styling.list_styles(self._styles)
        return [s for s in styles if style_type is None or s["type"] == style_type]

    def create_style(self, name: str, based_on: str | None = "Normal", style_type: str = "paragraph", **fmt) -> str:
        """Define a new paragraph or character style and return its name.
        ``based_on`` is the style it inherits from (the document's default
        when None or missing). ``fmt`` takes the keywords of
        ``styling.apply_run_format`` and, for a paragraph style,
        ``styling.apply_paragraph_format``. Style definitions are not tracked
        changes; applying the style to text is."""
        root = self._styles_root()
        based_on_id = None
        if based_on:
            try:
                based_on_id = self._style_id(based_on, style_type)
            except styling.StyleError:
                if styling._key(based_on) != "normal":
                    raise
        if based_on_id is None:
            based_on_id = styling.default_style_id(root, style_type)
        style_id = styling.create_style(root, name, style_type, based_on_id, **fmt)
        return styling.style_name(root, style_id, style_type)

    def set_paragraph_style(self, ref: str, style: str) -> int:
        """Apply a paragraph style by name or id. Returns the number of
        paragraphs whose style actually changed."""
        style_id = self._style_id(style)
        changed = 0
        for p in self._paragraphs_for(ref):
            if styling.paragraph_style_id(p) != style_id:
                self.edits.before_paragraph_format(p)
                styling.set_paragraph_style(p, style_id)
                changed += 1
        return changed

    def format_text(self, ref: str, match: str | None = None, case_sensitive: bool = True, **fmt) -> int:
        """Apply character formatting to a block, a table cell, or only the
        occurrences of ``match`` inside it. Returns the number of runs changed.

        Formatting keywords: bold, italic, underline, strike, color, size_pt,
        font, highlight (see ``styling.apply_run_format``)."""
        fmt = {k: v for k, v in fmt.items() if v is not None}
        if not fmt:
            raise DocumentError("Give at least one formatting property (bold, italic, color, size_pt...).")
        changed = 0
        for p in self._paragraphs_for(ref):
            if match:
                pattern = re.compile(re.escape(match), 0 if case_sensitive else re.IGNORECASE)
                spans = [(m.start(), m.end()) for m in pattern.finditer(paragraph_text(p))]
                runs = [run for start, end in spans for run in isolate_range(p, start, end)]
            else:
                runs = list(iter_runs(p))
            for run in runs:
                self.edits.before_run_format(run)
                styling.apply_run_format(run, **fmt)
            changed += len(runs)
        if match and changed == 0:
            raise DocumentError(f"Text {match!r} was not found in {ref}.")
        return changed

    def set_paragraph_format(self, ref: str, **fmt) -> int:
        """Apply paragraph formatting (alignment, spacing, indents...) to a
        block or table cell. See ``styling.apply_paragraph_format``."""
        fmt = {k: v for k, v in fmt.items() if v is not None}
        if not fmt:
            raise DocumentError("Give at least one paragraph property (alignment, space_after_pt...).")
        paragraphs = self._paragraphs_for(ref)
        for p in paragraphs:
            self.edits.before_paragraph_format(p)
            styling.apply_paragraph_format(p, **fmt)
        return len(paragraphs)

    # ------------------------------------------------------------- raw parts

    def list_parts(self) -> list[str]:
        return self.package.list_parts()

    def read_part_xml(self, name: str) -> str:
        return self.package.read_xml(name)

    def write_part_xml(self, name: str, xml: str) -> None:
        """Replace a whole XML part. Rewriting the main document part resets
        block ids."""
        self.package.write_xml(name, xml)
        if name == self.package.main_document_part:
            self._load()
