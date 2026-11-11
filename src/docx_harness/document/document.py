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
import os
import re
from pathlib import Path

from lxml import etree

from . import styling
from .edits import EditContext
from .ooxml import (
    BODY,
    P,
    R,
    RPR,
    SDT,
    SECTPR,
    TBL,
    TC,
    TR,
    el,
    isolate_range,
    iter_runs,
    make_paragraph,
    make_run,
    paragraph_text,
    qn,
    run_text,
)
from .package import WordPackage

BLOCK_TAGS = (P, TBL, SDT)
_REF = re.compile(r"^(b\d+)(?:\[\s*(\d+)\s*,\s*(\d+)\s*\])?$")
PAGE_WIDTH_TWIPS = 9360  # 6.5in text width on Letter with 1in margins


class DocumentError(ValueError):
    """A document operation could not be performed (bad id, bad argument)."""


class WordDocument:
    def __init__(self, package: WordPackage, *, track_changes: bool = False, author: str = "docx-harness"):
        self.package = package
        self.edits = EditContext(track_changes=track_changes, author=author)
        self._counter = 0
        self._load()

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
        return [child for child in self._body if child.tag in BLOCK_TAGS]

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
        element, cell = self._parse_ref(ref)
        if cell is not None:
            return list(self._cell(element, *cell).iter(P))
        return list(element.iter(P))

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
        return "\n".join(paragraph_text(p) for p in tc.iter(P))

    def _rows(self, tbl: etree._Element) -> list[list[str]]:
        return [[self._cell_text(tc) for tc in tr.findall(TC)] for tr in tbl.findall(TR)]

    def _describe(self, element: etree._Element, include_formatting: bool = False) -> dict:
        info: dict = {"id": self._id_of(element), "type": self._type(element)}
        if element.tag == P:
            info["style"] = styling.style_name(self._styles, styling.paragraph_style_id(element))
            info["text"] = paragraph_text(element)
            if include_formatting:
                info["runs"] = [
                    {"text": run_text(r), **styling.run_format(r)}
                    for r in iter_runs(element) if run_text(r)
                ]
        elif element.tag == TBL:
            info["rows"] = self._rows(element)
        else:
            info["text"] = "\n".join(paragraph_text(p) for p in element.iter(P))
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

    def read(self, start: str | None = None, count: int = 20, include_formatting: bool = False) -> dict:
        """Full content of ``count`` blocks beginning at ``start`` (inclusive)."""
        elements = self._block_elements()
        index = elements.index(self._get(start)) if start else 0
        chunk = elements[index:index + max(count, 0)]
        following = elements[index + len(chunk):index + len(chunk) + 1]
        return {
            "blocks": [self._describe(e, include_formatting) for e in chunk],
            "next": self._id_of(following[0]) if following else None,
            "total_blocks": len(elements),
        }

    def text(self) -> str:
        """Plain text of the whole body, one paragraph per line."""
        return "\n".join(paragraph_text(p) for p in self._body.iter(P))

    def search(self, query: str, case_sensitive: bool = False, limit: int = 50) -> list[dict]:
        """Find blocks (or table cells) containing ``query``."""
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
                check(block_id, "\n".join(paragraph_text(p) for p in element.iter(P)))
        return hits

    # ----------------------------------------------------------------- write

    def _insert(self, new: etree._Element, after: str) -> str:
        after = str(after).strip()
        if after == "end":
            last = self._body[-1] if len(self._body) else None
            index = len(self._body) - 1 if last is not None and last.tag == SECTPR else len(self._body)
        elif after == "start":
            index = 0
        else:
            index = self._body.index(self._get(after)) + 1
        self.edits.insert_block(self._body, index, new)
        return self._id_of(new)

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

    def _set_paragraph_text(self, p: etree._Element, text: str) -> None:
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
            runs = isolate_range(p, match.start(), match.end())
            first = runs[0]
            if replacement:
                parent = first.getparent()
                self.edits.insert_run(parent, parent.index(first), make_run(replacement, first.find(RPR)))
            for run in runs:
                self.edits.delete_run(run)
            pos = match.start() + len(make_run_text(replacement))
            count += 1

    def replace_text(self, find: str, replace: str, ref: str | None = None, case_sensitive: bool = True) -> int:
        """Replace every occurrence of ``find``. Matches may span differently
        formatted runs; the replacement takes the formatting at the start of
        the match. Scope is the whole document, or one block / table cell when
        ``ref`` is given. Returns the number of replacements."""
        if not find:
            raise DocumentError("find must not be empty")
        pattern = re.compile(re.escape(find), 0 if case_sensitive else re.IGNORECASE)
        paragraphs = self._paragraphs_for(ref) if ref else list(self._body.iter(P))
        return sum(self._replace_in_paragraph(p, pattern, replace) for p in paragraphs)

    # ---------------------------------------------------------------- tables

    def _cell(self, tbl: etree._Element, row: int, col: int) -> etree._Element:
        if tbl.tag != TBL:
            raise DocumentError("Cell references like 'b3[0,1]' only apply to tables.")
        rows = tbl.findall(TR)
        if not 0 <= row < len(rows):
            raise DocumentError(f"Row {row} is out of range; the table has {len(rows)} rows (zero-based).")
        cells = rows[row].findall(TC)
        if not 0 <= col < len(cells):
            raise DocumentError(f"Column {col} is out of range; row {row} has {len(cells)} cells (zero-based).")
        return cells[col]

    def _set_cell(self, tc: etree._Element, text: str) -> None:
        paragraphs = tc.findall(P)
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
        rows = tbl.findall(TR)
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
        new_cells = new_row.findall(TC)
        if len(cells) > len(new_cells):
            raise DocumentError(f"Got {len(cells)} cells but the table has {len(new_cells)} columns.")
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
        rows = tbl.findall(TR)
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


def make_run_text(text: str) -> str:
    """The text a run built from ``text`` will actually contain."""
    return run_text(make_run(text))


def _truncate(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"
