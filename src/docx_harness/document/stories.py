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
    PPR,
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
    P,
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
# Section start types: our names -> w:type values.
_START_TYPES = {"new_page": "nextPage", "continuous": "continuous", "even_page": "evenPage", "odd_page": "oddPage"}
_START_NAMES = {**{v: k for k, v in _START_TYPES.items()}, "nextColumn": "next_column"}


def _ends_section(block: etree._Element) -> bool:
    """True for a body paragraph that carries a section break."""
    ppr = block.find(PPR) if block.tag == P else None
    return ppr is not None and ppr.find(SECTPR) is not None


def _section_type(sect: etree._Element) -> str:
    kind = sect.find(qn("w:type"))
    return (kind.get(_VAL) if kind is not None else None) or "nextPage"


def _flag_on(element: etree._Element | None) -> bool:
    return element is not None and element.get(_VAL, "true").lower() not in ("0", "false", "off")


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
        shown_in: dict[str, list[int]] = {}
        if self.section_count() > 1:
            for section in self.sections():
                for kind in ("header", "footer"):
                    for shown in section[kind].values():
                        numbers = shown_in.setdefault(shown["part"], [])
                        if section["section"] not in numbers:
                            numbers.append(section["section"])
        out = []
        for part, kind, root in self._story_roots():
            entry = {"part": part, "kind": kind,
                     "paragraphs": [self._text_with_fields(p) for p in self._story_paragraphs(kind, root)]}
            if part in pages:
                entry["page"] = pages[part]
            if part in shown_in:
                entry["sections"] = shown_in[part]
            if _STORY_KINDS[kind] is not None:
                entry["notes"] = self._notes_in(root, kind)
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
        return [self._body.find(SECTPR) if p is None else p.find(PPR).find(SECTPR) for p in self._break_paragraphs()]

    def _break_paragraphs(self) -> list[etree._Element | None]:
        """The paragraph that ends each section, in order; None for the last
        section, which the body's own ``w:sectPr`` ends."""
        return [p for p in self._block_elements() if _ends_section(p)] + [None]

    def section_count(self) -> int:
        """Number of sections in the document (changes nothing)."""
        return len(self._break_paragraphs())

    def _section_index(self, section: int) -> int:
        count = self.section_count()
        try:
            number = int(section)
        except (TypeError, ValueError):
            number = 0
        if not 1 <= number <= count:
            raise DocumentError(
                f"Section {section} does not exist; the document has {count} section{'s' * (count != 1)} "
                "(numbered from 1; list_sections shows them)."
            )
        return number - 1

    def _section_spans(self) -> list[list[etree._Element]]:
        """The body blocks of each section, in order."""
        spans: list[list[etree._Element]] = [[]]
        for block in self._block_elements():
            spans[-1].append(block)
            if _ends_section(block):
                spans.append([])
        return spans

    def _own_reference(self, sect: etree._Element, kind: str, page: str) -> tuple[etree._Element, str] | None:
        """The section's own reference of this kind and page, and its part."""
        for reference in sect.findall(qn(f"w:{kind}Reference")):
            if reference.get(qn("w:type"), "default") != page:
                continue
            relationship = self.package.relationship(self.package.main_document_part, reference.get(R_ID, ""))
            if relationship and relationship["target"] in self.package:
                return reference, relationship["target"]
        return None

    def _effective_reference(self, sections: list[etree._Element], index: int, kind: str,
                             page: str) -> tuple[int, str] | None:
        """(index of the section that owns it, part) for the header or footer
        section ``index`` shows on ``page``: its own, else the nearest earlier
        section's, as Word links a section to the previous one."""
        for j in range(index, -1, -1):
            own = self._own_reference(sections[j], kind, page)
            if own is not None:
                return j, own[1]
        return None

    def sections(self) -> list[dict]:
        """The sections of the document, numbered from 1. Each has ``start``
        and ``end`` (its first and last block id), ``start_type`` (how it
        begins: new_page, continuous, even_page, odd_page or next_column),
        ``first_page_different``, and ``header`` / ``footer``: for each page
        type in use ("default", "first", "even"), ``{"part": name}``, with
        ``"from_section": N`` when the section shows section N's (linked to
        previous, in Word's terms)."""
        sections = self._sections()
        out = []
        for index, (sect, blocks) in enumerate(zip(sections, self._section_spans())):
            entry: dict = {
                "section": index + 1,
                "start": self._id_of(blocks[0]) if blocks else None,
                "end": self._id_of(blocks[-1]) if blocks else None,
                "start_type": _START_NAMES.get(_section_type(sect), _section_type(sect)),
                "first_page_different": _flag_on(sect.find(qn("w:titlePg"))),
            }
            for kind in ("header", "footer"):
                shown = {}
                for page in _PAGES:
                    found = self._effective_reference(sections, index, kind, page)
                    if found is not None:
                        shown[page] = {"part": found[1]}
                        if found[0] != index:
                            shown[page]["from_section"] = found[0] + 1
                entry[kind] = shown
            out.append(entry)
        return out

    def insert_section_break(self, after: str, start: str = "new_page") -> int:
        """Start a new section after block ``after``: an empty paragraph that
        ends the section is inserted there. ``start`` is how the new section
        begins: "new_page", "continuous", "even_page" or "odd_page". Both
        halves keep the page setup, headers and footers the section had,
        until :meth:`set_header` / :meth:`set_footer` is given a ``section``.
        Returns the new section's number (counted from 1).

        The paragraph is a tracked insertion when tracking is on; the section
        properties themselves are not recorded as a revision."""
        start = str(start).strip().lower().replace("-", "_").replace(" ", "_")
        if start not in _START_TYPES:
            raise DocumentError(f"Invalid start {start!r}. Use 'new_page', 'continuous', 'even_page' or 'odd_page'.")
        block = self._body_block(after)
        if _ends_section(block):
            raise DocumentError(f"Block {after} already ends a section. Choose the block before or after it.")
        sections = self._sections()
        index = next(i for i, span in enumerate(self._section_spans()) if block in span)
        sect = sections[index]
        first_half = copy.deepcopy(sect)  # keeps the way the original section started
        for change in first_half.findall(qn("w:sectPrChange")):
            first_half.remove(change)
        # The original w:sectPr now ends the second half, which starts at the break.
        ensure_child(sect, "w:type", SECTPR_ORDER).set(_VAL, _START_TYPES[start])
        p = el("w:p", None, el("w:pPr", None, first_half))
        self.edits.insert_block(self._body, self._body.index(block) + 1, p)
        self._id_of(p)
        return index + 2

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

    def _rewrite_story(self, part: str, new: list[etree._Element]) -> None:
        """Replace what a header or footer part says (a tracked change when tracking)."""
        root = self.package.xml(part)
        old = [child for child in root if child.tag in BLOCK_TAGS and not is_deleted_block(child)]
        for offset, p in enumerate(new):
            self.edits.insert_block(root, offset, p)
        for block in old:
            self.edits.delete_block(block)

    def _new_header_footer_part(self, kind: str, paragraphs: list[etree._Element]) -> str:
        tag = "w:hdr" if kind == "header" else "w:ftr"
        part = self.package.add_part(
            self.package.next_part_name(self._part_beside_main(kind), ".xml"),
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<{tag} xmlns:w="{W_NS}"/>',
            _HEADER_FOOTER_TYPES[kind],
        )
        root = self.package.xml(part)
        with self.edits.untracked():  # a new part is not a revision of anything
            for offset, p in enumerate(paragraphs or [make_paragraph("", self._style_id(kind.capitalize()))]):
                self.edits.insert_block(root, offset, p)
        return part

    def _add_reference(self, sect: etree._Element, kind: str, page: str, part: str) -> None:
        rel_id = self.package.relate(self.package.main_document_part, kind, part)
        reference = el(f"w:{kind}Reference", {"w:type": page, "r:id": rel_id})
        sect.insert(len([c for c in sect if c.tag in (qn("w:headerReference"), qn("w:footerReference"))]), reference)

    def _reference_count(self, part: str) -> int:
        main = self.package.main_document_part
        return sum(
            1 for kind in ("header", "footer") for reference in self._references(kind)
            if (self.package.relationship(main, reference.get(R_ID, "")) or {}).get("target") == part
        )

    def _set_section_header_footer(self, kind: str, new: list[etree._Element], page: str, section: int) -> tuple[str, list]:
        """Give one section its own header or footer, leaving the others as they look now."""
        sections = self._sections()
        index = self._section_index(section)
        sect = sections[index]
        own = self._own_reference(sect, kind, page)
        if own is not None:
            reference, part = own
            if self._reference_count(part) > 1:
                # Shared with another section: copy it first so only this one changes.
                part = self.package.duplicate_part(part)
                reference.set(R_ID, self.package.relate(self.package.main_document_part, kind, part))
                self._bind()
            self._rewrite_story(part, new)
            return part, [sect]
        # Linked to the previous section. The next section may be linked to
        # this one: pin it to what it shows now (a blank part if nothing).
        if index + 1 < len(sections) and self._own_reference(sections[index + 1], kind, page) is None:
            shown = self._effective_reference(sections, index, kind, page)
            self._add_reference(sections[index + 1], kind, page,
                                shown[1] if shown else self._new_header_footer_part(kind, []))
        part = self._new_header_footer_part(kind, new)
        self._add_reference(sect, kind, page, part)
        self._bind()
        return part, [sect]

    def _set_header_footer(self, kind: str, text: str, page: str, alignment: str | None,
                           section: int | None = None) -> str:
        page = str(page).strip().lower()
        if page not in _PAGES:
            raise DocumentError(f"Invalid page {page!r}. Use 'default', 'first' (first page only) or 'even' (even pages).")
        new = self._header_footer_paragraphs(kind, str(text), alignment)
        if section is not None:
            part, used = self._set_section_header_footer(kind, new, page, section)
            parts = [part]
        else:
            parts, used = [], []
            for sect in self._sections():
                own = self._own_reference(sect, kind, page)
                if own is not None:
                    used.append(sect)
                    if own[1] not in parts:
                        parts.append(own[1])
            for index, part in enumerate(parts):
                self._rewrite_story(part, new if index == 0 else [copy.deepcopy(p) for p in new])
            if not parts:
                first = self._sections()[0]
                parts, used = [self._new_header_footer_part(kind, new)], [first]
                self._add_reference(first, kind, page, parts[0])
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

    def set_header(self, text: str, page: str = "default", alignment: str | None = None,
                   section: int | None = None) -> str:
        """Set the text of the page header, creating it when the document has
        none. Each line of ``text`` is a paragraph; ``{PAGE}`` and
        ``{NUMPAGES}`` become the page number and page count, and a tab
        (``\\t``) moves to the centre, then to the right edge. ``page`` is
        which pages show it: "default" (all, unless the others are set),
        "first" or "even" (even pages, a setting for the whole document).
        Returns the part name, which works as ``ref`` in ``replace_text``,
        ``format_text`` and ``set_paragraph_format``.

        Without ``section`` this sets the header every section shows. With
        a section number (from 1, see :meth:`sections`) only that section
        changes: a header it shares with other sections is copied first, and
        when it showed the previous section's, the next section keeps what it
        showed.

        Replacing an existing header's text is tracked when tracking is on;
        creating the header is not."""
        return self._set_header_footer("header", text, page, alignment, section)

    def set_footer(self, text: str, page: str = "default", alignment: str | None = None,
                   section: int | None = None) -> str:
        """Set the text of the page footer; see :meth:`set_header`."""
        return self._set_header_footer("footer", text, page, alignment, section)
