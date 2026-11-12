"""WordDocument: block-level operations over a WordPackage.

The document body is exposed as an ordered list of *blocks* (paragraphs,
tables and content controls). Each block gets a stable id (``b1``, ``b2``...)
when the document is loaded or the block is created. Ids live in memory only
and never shift when other blocks are inserted or deleted.

A cell of a table is addressed as ``"<table id>[row,col]"``, for example
``"b7[0,1]"`` (zero-based).
"""

from __future__ import annotations

import copy
import difflib
import os
import re
from pathlib import Path

from lxml import etree

from . import styling
from .edits import EditContext
from .ooxml import (
    AUTHOR,
    BODY,
    DEL,
    DEL_TEXT,
    INS,
    P,
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
    TC,
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
    run_text,
    strip_revisions,
)
from .package import PackageError, WordPackage

BLOCK_TAGS = (P, TBL, SDT)
_TOKEN = re.compile(r"\s+|\w+|[^\w\s]")
_REVISION_TYPES = {INS: "insert", DEL: "delete", MOVE_FROM: "move", MOVE_TO: "move",
                   **{tag: "format" for tag in PROPERTY_CHANGES}}
_REF = re.compile(r"^(b\d+)(?:\[\s*(\d+)\s*,\s*(\d+)\s*\])?$")
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
        if element is None or element.getparent() is not self._body:
            raise DocumentError(
                f"No block with id {block_id!r}. Use the outline to see current block ids."
            )
        return element

    def _parse_ref(self, ref: str) -> tuple[etree._Element, tuple[int, int] | None]:
        match = _REF.match(str(ref).strip())
        if not match:
            raise DocumentError(
                f"Invalid block reference {ref!r}. Use a block id like 'b3' or a table cell like 'b3[0,1]'."
            )
        element = self._get(match.group(1))
        if match.group(2) is None:
            return element, None
        return element, (int(match.group(2)), int(match.group(3)))

    def _paragraphs_for(self, ref: str) -> list[etree._Element]:
        for part, kind, root in self._story_roots():
            if part == str(ref).strip():
                return self._story_paragraphs(kind, root)
        element, cell = self._parse_ref(ref)
        if cell is not None:
            return _live_paragraphs(self._cell(element, *cell))
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
        return "\n".join(paragraph_text(p) for p in _live_paragraphs(tc))

    @staticmethod
    def _live_rows(tbl: etree._Element) -> list[etree._Element]:
        """Rows that have not been removed by a tracked deletion."""
        return [tr for tr in tbl.findall(TR) if not is_deleted_block(tr)]

    def _rows(self, tbl: etree._Element) -> list[list[str]]:
        return [[self._cell_text(tc) for tc in tr.findall(TC)] for tr in self._live_rows(tbl)]

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
            info["rows"] = self._rows(element)
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
        index = elements.index(self._get(start)) if start else 0
        return elements[index:index + max(count, 0)]

    def read(self, start: str | None = None, count: int = 20, include_formatting: bool = False,
             resolve_styles: bool = False) -> dict:
        """Full content of ``count`` blocks beginning at ``start`` (inclusive).

        ``include_formatting`` adds each paragraph's direct formatting:
        ``format`` (alignment, spacing, indents; left out when there is none)
        and ``runs``. ``resolve_styles`` also adds ``style_format``, what the
        paragraph's style supplies underneath, for rendering a preview."""
        elements = self._block_elements()
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

        for element in self._block_elements():
            block_id = self._id_of(element)
            if element.tag == TBL:
                for r, row in enumerate(self._rows(element)):
                    for c, cell in enumerate(row):
                        check(f"{block_id}[{r},{c}]", cell)
            else:
                check(block_id, "\n".join(paragraph_text(p) for p in _live_paragraphs(element)))
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
        return self._body.index(self._get(after)) + 1

    def _insert(self, new: etree._Element, after: str) -> str:
        self.edits.insert_block(self._body, self._index_after(after), new)
        return self._id_of(new)

    def move_block(self, block_id: str, after: str) -> None:
        """Move a block to after block ``after`` ("start", "end" or a block
        id). The block keeps its id. When changes are tracked, a paragraph is
        recorded as a move and a table as a deletion plus an insertion."""
        element = self._get(block_id)
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
    ) -> str:
        """Insert a table. ``rows`` is a list of rows, each a list of cell texts.
        When ``header`` is true the first row is bold."""
        if not rows or not all(isinstance(r, (list, tuple)) for r in rows):
            raise DocumentError("rows must be a non-empty list of lists of cell text")
        cols = max(len(r) for r in rows)
        if cols == 0:
            raise DocumentError("rows must contain at least one cell")
        width = PAGE_WIDTH_TWIPS // cols

        tbl_pr = el("w:tblPr")
        style_id = None
        if style:
            try:
                style_id = self._style_id(style, "table")
            except styling.StyleError:
                style_id = None
        if style_id:
            tbl_pr.append(el("w:tblStyle", {"w:val": style_id}))
        tbl_pr.append(el("w:tblW", {"w:w": "5000", "w:type": "pct"}))
        if not style_id:
            # No usable table style in this document: draw explicit borders.
            borders = el("w:tblBorders")
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                borders.append(el(f"w:{side}", {"w:val": "single", "w:sz": "4", "w:space": "0", "w:color": "auto"}))
            tbl_pr.append(borders)
        tbl_pr.append(el("w:tblLook", {"w:val": "04A0"}))

        tbl = el("w:tbl", None, tbl_pr)
        grid = el("w:tblGrid")
        for _ in range(cols):
            grid.append(el("w:gridCol", {"w:w": str(width)}))
        tbl.append(grid)
        for r_index, row in enumerate(rows):
            tr = el("w:tr")
            for c_index in range(cols):
                text = str(row[c_index]) if c_index < len(row) and row[c_index] is not None else ""
                p = make_paragraph(text)
                if header and r_index == 0:
                    for run in p.iter(R):
                        styling.apply_run_format(run, bold=True)
                tr.append(el("w:tc", None, el("w:tcPr", None, el("w:tcW", {"w:w": str(width), "w:type": "dxa"})), p))
            tbl.append(tr)
        return self._insert(tbl, after)

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

    def _cell(self, tbl: etree._Element, row: int, col: int) -> etree._Element:
        if tbl.tag != TBL:
            raise DocumentError("Cell references like 'b3[0,1]' only apply to tables.")
        rows = self._live_rows(tbl)
        if not 0 <= row < len(rows):
            raise DocumentError(f"Row {row} is out of range; the table has {len(rows)} rows (zero-based).")
        cells = rows[row].findall(TC)
        if not 0 <= col < len(cells):
            raise DocumentError(f"Column {col} is out of range; row {row} has {len(cells)} cells (zero-based).")
        return cells[col]

    def _set_cell(self, tc: etree._Element, text: str) -> None:
        paragraphs = [p for p in tc.findall(P) if not is_deleted_block(p)]
        if not paragraphs:
            paragraphs = [make_paragraph()]
            self.edits.insert_block(tc, len(tc), paragraphs[0])
        self._set_paragraph_text(paragraphs[0], text)
        for extra in paragraphs[1:]:
            self.edits.delete_block(extra)

    def set_cell_text(self, table_id: str, row: int, col: int, text: str) -> None:
        self._set_cell(self._cell(self._table(table_id), row, col), text)

    def insert_table_row(self, table_id: str, cells: list[str], index: int | None = None) -> int:
        """Insert a row at ``index`` (default: append). The new row copies the
        layout of a neighbouring row. Returns the index of the new row."""
        tbl = self._table(table_id)
        rows = self._live_rows(tbl)
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
        new_cells = new_row.findall(TC)
        if len(cells) > len(new_cells):
            raise DocumentError(f"Got {len(cells)} cells but the table has {len(new_cells)} columns.")
        with self.edits.untracked():
            for i, tc in enumerate(new_cells):
                self._set_cell(tc, str(cells[i]) if i < len(cells) else "")
        if index == len(rows):
            anchor = rows[-1]
            self.edits.insert_block(tbl, tbl.index(anchor) + 1, new_row)
        else:
            self.edits.insert_block(tbl, tbl.index(rows[index]), new_row)
        return index

    def delete_table_row(self, table_id: str, row: int) -> None:
        tbl = self._table(table_id)
        rows = self._live_rows(tbl)
        if not 0 <= row < len(rows):
            raise DocumentError(f"Row {row} is out of range; the table has {len(rows)} rows (zero-based).")
        if len(rows) == 1:
            raise DocumentError("Cannot delete the only row; delete the table block instead.")
        self.edits.delete_block(rows[row])

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
            and (scope is None or (part is None and self._top(node) is scope))
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

        def is_mark(node: etree._Element) -> bool:  # on a paragraph mark or a table row
            return node.tag not in PROPERTY_CHANGES and node.getparent().tag in (RPR, TRPR)

        # Text first, then paragraph marks and rows: a paragraph whose inserted
        # text was rejected is empty by the time its mark is, and goes entirely.
        for node in sorted(nodes, key=is_mark):
            if any(ancestor is root for ancestor in node.iterancestors() for root in roots):
                self.edits.resolve(node, accept)
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
        for element in [e for e in self._by_element if e.getparent() is not self._body]:
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


def make_run_text(text: str) -> str:
    """The text a run built from ``text`` will actually contain."""
    return run_text(make_run(text))


def _truncate(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"
