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

import copy
import difflib
import os
import re
from pathlib import Path

from lxml import etree

from . import styling, tables
from .edits import EditContext
from .ooxml import (
    AUTHOR,
    BODY,
    CELL_DEL,
    CELL_INS,
    DEL,
    DEL_TEXT,
    INS,
    P,
    PPR,
    MOVE_FROM,
    MOVE_RANGE_STARTS,
    MOVE_TO,
    PPR_CHANGE,
    PROPERTY_CHANGES,
    R,
    RPR,
    RPR_CHANGE,
    SDT,
    SECTPR,
    SETTINGS_BEFORE_TRACK_REVISIONS,
    T,
    TBL,
    TCPR,
    TR,
    TRACK_REVISIONS,
    TRPR,
    el,
    insert_after_any,
    is_deleted_block,
    isolate_range,
    iter_runs,
    make_paragraph,
    make_run,
    paragraph_text,
    qn,
    remove_child,
    run_text,
    strip_revisions,
)
from .package import PackageError, WordPackage

BLOCK_TAGS = (P, TBL, SDT)
_TOKEN = re.compile(r"\s+|\w+|[^\w\s]")
_REVISION_TYPES = {INS: "insert", DEL: "delete", MOVE_FROM: "move", MOVE_TO: "move",
                   CELL_INS: "insert", CELL_DEL: "delete",
                   **{tag: "format" for tag in PROPERTY_CHANGES}}
_REF = re.compile(r"^(b\d+)(?:\[\s*(\d+|\*)\s*,\s*(\d+|\*)\s*\])?$")
_VERTICAL = {"top": "top", "center": "center", "centre": "center", "middle": "center", "bottom": "bottom"}
_TABLE_ALIGNMENTS = {"left": "left", "center": "center", "centre": "center", "right": "right"}
_NOT_COPIED_TO_NEW_CELLS = ("w:gridSpan", "w:hMerge", "w:vMerge", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange")
# Text outside the body: relationship kind -> the element holding each note
# (None when the whole part is one story).
_STORY_KINDS = {"header": None, "footer": None, "footnotes": qn("w:footnote"), "endnotes": qn("w:endnote")}
_NOTE_SEPARATORS = ("separator", "continuationSeparator", "continuationNotice")
PAGE_WIDTH_TWIPS = 9360  # 6.5in text width on Letter with 1in margins


class DocumentError(ValueError):
    """A document operation could not be performed (bad id, bad argument)."""


class WordDocument:
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
        self.edits.bind(root, *(story for _, _, story in self._story_roots()))
        self._ids: dict[str, etree._Element] = {}
        self._by_element: dict[etree._Element, str] = {}
        for element in self._block_elements():
            self._id_of(element)

    @property
    def _styles(self) -> etree._Element | None:
        part = self.package.styles_part
        return self.package.xml(part) if part else None

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

    # ---------------------------------------------- headers, footers, notes

    def _story_roots(self) -> list[tuple[str, str, etree._Element]]:
        """(part name, kind, root) of every header, footer, footnotes and
        endnotes part of the document."""
        found = []
        for kind in _STORY_KINDS:
            for part in self.package.related_parts(kind):
                try:
                    found.append((part, kind, self.package.xml(part)))
                except PackageError:
                    continue  # unreadable part: leave it alone
        return found

    @staticmethod
    def _story_paragraphs(kind: str, root: etree._Element) -> list[etree._Element]:
        note = _STORY_KINDS[kind]
        if note is None:
            return _live_paragraphs(root)
        return [
            p for holder in root.iter(note) if holder.get(qn("w:type")) not in _NOTE_SEPARATORS
            for p in _live_paragraphs(holder)
        ]

    def stories(self) -> list[dict]:
        """The text outside the body: one entry per header, footer, footnotes
        or endnotes part, with its ``part`` name, ``kind`` and ``paragraphs``
        (their text). The part name works as ``ref`` in :meth:`replace_text`."""
        return [
            {"part": part, "kind": kind, "paragraphs": [paragraph_text(p) for p in self._story_paragraphs(kind, root)]}
            for part, kind, root in self._story_roots()
        ]

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
            if cells:
                layout["cells"] = cells
            if layout:
                info["format"] = layout
        return info

    def _describe(self, element: etree._Element, include_formatting: bool = False, resolve_styles: bool = False) -> dict:
        info: dict = {"id": self._id_of(element), "type": self._type(element)}
        if element.tag == P:
            info["style"] = styling.style_name(self._styles, styling.paragraph_style_id(element))
            info["text"] = paragraph_text(element)
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
                if "style" in info:
                    entry["style"] = info["style"]
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
        paragraph's style supplies underneath, for rendering a preview.

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
        return {
            "blocks": [self._describe(e, include_formatting, resolve_styles) for e in chunk],
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

    def move_block(self, block_id: str, after: str) -> None:
        """Move a block to after block ``after`` ("start", "end" or a block
        id). The block keeps its id. When changes are tracked, a paragraph is
        recorded as a move and a table as a deletion plus an insertion."""
        element = self._body_block(block_id)
        moved = self.edits.move_block(element, self._body, self._index_after(after))
        if moved is not element:  # tracked: the id follows the content to its new place
            block_id = self._by_element.pop(element)
            self._ids[block_id] = moved
            self._by_element[moved] = block_id

    def _style_id(self, style: str | None, style_type: str = "paragraph") -> str | None:
        if not style:
            return None
        return styling.resolve_style(self._styles, style, style_type)

    def insert_paragraph(self, text: str = "", style: str | None = None, after: str = "end") -> str:
        """Insert a paragraph after block ``after`` ("start", "end" or a block id).
        Returns the new block id. Newlines in ``text`` become line breaks."""
        return self._insert(make_paragraph(text, self._style_id(style)), after)

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

    def format_cells(self, ref: str, shading: str | None = None, vertical_align: str | None = None) -> int:
        """Format table cells. ``ref`` is a cell (``b3[0,1]``), a row
        (``b3[0,*]``), a column (``b3[*,1]``) or a table id for every cell.
        ``shading`` is a background color (hex or a basic color name, "none"
        to clear); ``vertical_align`` is top, center or bottom. Returns the
        number of cells formatted. For the text itself use ``format_text``
        and ``set_paragraph_format`` with the same ``ref``."""
        if shading is None and vertical_align is None:
            raise DocumentError("Give at least one cell property (shading, vertical_align).")
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
        slots = self._slots(element, *(cell or (None, None)))
        for slot in slots:
            for tc in slot.members:
                self.edits.before_cell_format(tc)
                tables.apply_cell_format(tc, fill, align)
        return len(slots)

    def set_table_format(
        self,
        table_id: str,
        column_widths_in: list[float] | None = None,
        header_rows: int | None = None,
        alignment: str | None = None,
    ) -> None:
        """Lay out a table. ``column_widths_in`` sets every column's width in
        inches and fixes the layout so Word keeps them; ``header_rows`` is how
        many rows at the top repeat on each page the table runs onto (0 for
        none); ``alignment`` places the table left, center or right."""
        tbl = self._table(table_id)
        if column_widths_in is None and header_rows is None and alignment is None:
            raise DocumentError("Give at least one table property (column_widths_in, header_rows, alignment).")
        grid = tables.build_grid(tbl)
        if alignment is not None and str(alignment).strip().lower() not in _TABLE_ALIGNMENTS:
            raise DocumentError(f"Invalid alignment {alignment!r}. Use 'left', 'center' or 'right'.")
        if header_rows is not None and not 0 <= int(header_rows) <= len(grid.rows):
            raise DocumentError(f"header_rows must be between 0 and {len(grid.rows)}, the number of rows.")
        twips = self._twips(column_widths_in, grid.width) if column_widths_in is not None else None

        if twips is not None or alignment is not None:
            self.edits.before_table_format(tbl)
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

    # ---------------------------------------------------------------- delete

    def delete_block(self, block_id: str) -> None:
        element = self._get(block_id)
        self.edits.delete_block(element)
        self._ids.pop(self._by_element.pop(element), None)

    # --------------------------------------------------------------- styling

    def list_styles(self, style_type: str | None = None) -> list[dict]:
        styles = styling.list_styles(self._styles)
        return [s for s in styles if style_type is None or s["type"] == style_type]

    def set_paragraph_style(self, ref: str, style: str) -> int:
        """Apply a paragraph style by name or id. Returns the number of
        paragraphs whose style actually changed."""
        style_id = styling.resolve_style(self._styles, style, "paragraph")
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
        for node in sorted(nodes, key=is_mark):
            if any(ancestor is root for ancestor in node.iterancestors() for root in roots):
                self.edits.resolve(node, accept)
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


def _live_paragraphs(element: etree._Element) -> list[etree._Element]:
    """Paragraphs in ``element`` (itself included) not removed by a tracked deletion."""
    return [p for p in element.iter(P) if not is_deleted_block(p)]


def _lines(text: str) -> list[str]:
    """Cell text as paragraphs: one per line."""
    return str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")


def make_run_text(text: str) -> str:
    """The text a run built from ``text`` will actually contain."""
    return run_text(make_run(text))


def _truncate(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"
