"""Headers, footers, footnotes and endnotes (the stories outside the body), and the settings part.

Mixed into ``WordDocument`` (``document.py``), which provides the block-id
helpers it relies on (``_get``, ``_body_block``, ``_table``, ``_paragraphs_for``,
``_describe``...) and ``self.edits``, the single path for every mutation."""

from __future__ import annotations

import copy
import re

from lxml import etree

from . import styling
from .common import (
    _VAL,
    BLOCK_TAGS,
    DocumentError,
    _lines,
    _live_paragraphs,
)
from .ooxml import (
    FLD_SIMPLE,
    R_ID,
    SECTPR,
    SECTPR_ORDER,
    SETTINGS_BEFORE_EVEN_AND_ODD,
    W_NS,
    el,
    ensure_child,
    insert_after_any,
    is_deleted_block,
    iter_runs,
    make_paragraph,
    make_run,
    qn,
    run_text,
)
from .package import PackageError

# Text outside the body: relationship kind -> the element holding each note
# (None when the whole part is one story).
_STORY_KINDS = {"header": None, "footer": None, "footnotes": qn("w:footnote"), "endnotes": qn("w:endnote")}
_NOTE_SEPARATORS = ("separator", "continuationSeparator", "continuationNotice")
_PAGES = ("default", "first", "even")
_FIELD_TOKEN = re.compile(r"\{(PAGE|NUMPAGES)\}")
_HEADER_FOOTER_TYPES = {
    "header": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    "footer": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
}
_SETTINGS_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"


class StoriesMixin:
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
        (their text; a page number field written by :meth:`set_header` reads
        as ``{PAGE}`` or ``{NUMPAGES}``). A header or footer also has ``page``:
        which pages show it (default, first or even). The part name works as
        ``ref`` in :meth:`replace_text` and the formatting methods."""
        pages = {}
        for kind in ("header", "footer"):
            for reference in self._references(kind):
                relationship = self.package.relationship(self.package.main_document_part, reference.get(R_ID, ""))
                if relationship:
                    pages.setdefault(relationship["target"], reference.get(qn("w:type"), "default"))
        out = []
        for part, kind, root in self._story_roots():
            entry = {"part": part, "kind": kind,
                     "paragraphs": [self._text_with_fields(p) for p in self._story_paragraphs(kind, root)]}
            if part in pages:
                entry["page"] = pages[part]
            out.append(entry)
        return out

    @staticmethod
    def _text_with_fields(p: etree._Element) -> str:
        """Paragraph text with simple PAGE / NUMPAGES fields shown as tokens."""
        parts, seen = [], set()
        for run in iter_runs(p):
            # The run may sit in a revision mark inside the field.
            field = next((a for a in run.iterancestors() if a is p or a.tag == FLD_SIMPLE), None)
            field = field if field is not None and field.tag == FLD_SIMPLE else None
            name = (field.get(qn("w:instr")) or "").strip().split(" ")[0].upper() if field is not None else ""
            if name in ("PAGE", "NUMPAGES"):
                if field not in seen:
                    seen.add(field)
                    parts.append("{" + name + "}")
            else:
                parts.append(run_text(run))
        return "".join(parts)

    def _sections(self) -> list[etree._Element]:
        """The ``w:sectPr`` of every section, in order; the last one is the
        body's own and is created when the file has none."""
        if self._body.find(SECTPR) is None:
            self._body.append(el("w:sectPr"))  # layout, not content
        return [s for s in self._body.iter(SECTPR) if s.getparent().tag != qn("w:sectPrChange")]

    def _references(self, kind: str) -> list[etree._Element]:
        return [ref for sect in self._sections() for ref in sect.findall(qn(f"w:{kind}Reference"))]

    def _header_footer_paragraphs(self, kind: str, text: str, alignment: str | None) -> list[etree._Element]:
        style_id = self._style_id(kind.capitalize())
        paragraphs = []
        for line in _lines(text):
            p = make_paragraph("", style_id)
            position = 0
            for token in _FIELD_TOKEN.finditer(line):
                if token.start() > position:
                    p.append(make_run(line[position:token.start()]))
                p.append(el("w:fldSimple", {"w:instr": f" {token.group(1)} "}, make_run("1")))
                position = token.end()
            if position < len(line):
                p.append(make_run(line[position:]))
            if alignment is not None:
                styling.apply_paragraph_format(p, alignment=alignment)
            paragraphs.append(p)
        return paragraphs

    def _set_header_footer(self, kind: str, text: str, page: str, alignment: str | None) -> str:
        page = str(page).strip().lower()
        if page not in _PAGES:
            raise DocumentError(f"Invalid page {page!r}. Use 'default', 'first' (first page only) or 'even' (even pages).")
        new = self._header_footer_paragraphs(kind, str(text), alignment)
        main = self.package.main_document_part
        parts: list[str] = []
        used: list[etree._Element] = []
        for sect in self._sections():
            for reference in sect.findall(qn(f"w:{kind}Reference")):
                relationship = self.package.relationship(main, reference.get(R_ID, ""))
                if reference.get(qn("w:type"), "default") == page and relationship and relationship["target"] in self.package:
                    used.append(sect)
                    if relationship["target"] not in parts:
                        parts.append(relationship["target"])
        if parts:
            # Replace what the existing part says (a tracked change when tracking).
            for index, part in enumerate(parts):
                root = self.package.xml(part)
                old = [child for child in root if child.tag in BLOCK_TAGS and not is_deleted_block(child)]
                for offset, p in enumerate(new if index == 0 else [copy.deepcopy(p) for p in new]):
                    self.edits.insert_block(root, offset, p)
                for block in old:
                    self.edits.delete_block(block)
        else:
            tag = "w:hdr" if kind == "header" else "w:ftr"
            part = self.package.add_part(
                self.package.next_part_name(self._part_beside_main(kind), ".xml"),
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<{tag} xmlns:w="{W_NS}"/>',
                _HEADER_FOOTER_TYPES[kind],
            )
            root = self.package.xml(part)
            with self.edits.untracked():  # a new part is not a revision of anything
                for offset, p in enumerate(new):
                    self.edits.insert_block(root, offset, p)
            rel_id = self.package.relate(main, kind, part)
            first = self._sections()[0]
            reference = el(f"w:{kind}Reference", {"w:type": page, "r:id": rel_id})
            first.insert(len([c for c in first if c.tag in (qn("w:headerReference"), qn("w:footerReference"))]), reference)
            parts, used = [part], [first]
            self._bind()
        if page == "first":
            for sect in used:
                ensure_child(sect, "w:titlePg", SECTPR_ORDER).attrib.pop(_VAL, None)
        elif page == "even":
            settings = self._settings_root()
            flag = settings.find(qn("w:evenAndOddHeaders"))
            if flag is None:
                insert_after_any(settings, el("w:evenAndOddHeaders"), SETTINGS_BEFORE_EVEN_AND_ODD)
            else:
                flag.attrib.pop(_VAL, None)
        return parts[0]

    def _settings_root(self) -> etree._Element:
        if self.package.settings_part is None:
            name = self.package.add_part(
                self._part_beside_main("settings.xml"),
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings xmlns:w="{W_NS}"/>', _SETTINGS_TYPE,
            )
            self.package.relate(self.package.main_document_part, "settings", name)
        return self.package.xml(self.package.settings_part)

    def set_header(self, text: str, page: str = "default", alignment: str | None = None) -> str:
        """Set the text of the page header, creating it when the document has
        none. Each line of ``text`` is a paragraph; ``{PAGE}`` and
        ``{NUMPAGES}`` become the page number and page count, and a tab
        (``\\t``) moves to the centre, then to the right edge. ``page`` is
        which pages show it: "default" (all, unless the others are set),
        "first" or "even". Returns the part name, which works as ``ref`` in
        ``replace_text``, ``format_text`` and ``set_paragraph_format``.

        Replacing an existing header's text is tracked when tracking is on;
        creating the header is not."""
        return self._set_header_footer("header", text, page, alignment)

    def set_footer(self, text: str, page: str = "default", alignment: str | None = None) -> str:
        """Set the text of the page footer; see :meth:`set_header`."""
        return self._set_header_footer("footer", text, page, alignment)
