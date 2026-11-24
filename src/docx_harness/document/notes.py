"""Creating and deleting footnotes and endnotes.

Mixed into ``WordDocument`` (``document.py``). Reading note text and editing
it by part name is in ``stories.py``; this adds the notes themselves: the
note in ``word/footnotes.xml`` (or ``endnotes.xml``) and the reference mark
in the body. Every content change goes through ``self.edits``."""

from __future__ import annotations

import re

from lxml import etree

from .common import DocumentError, _lines
from .ooxml import (
    HYPERLINK,
    SETTINGS_BEFORE_FOOTNOTE_PR,
    W_NS,
    el,
    insert_after_any,
    isolate_range,
    iter_runs,
    make_run,
    paragraph_text,
    qn,
)

NOTE_KINDS = ("footnote", "endnote")
_NOTES_TYPES = {
    "footnote": "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
    "endnote": "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
}
_SEPARATORS = (("-1", "separator"), ("0", "continuationSeparator"))


def _kind(kind: str) -> str:
    kind = str(kind).strip().lower().rstrip("s")
    if kind not in NOTE_KINDS:
        raise DocumentError(f"Invalid kind {kind!r}. Use 'footnote' or 'endnote'.")
    return kind


class NotesMixin:
    # ---------------------------------------------------------------- reading

    def _notes_root(self, kind: str) -> etree._Element | None:
        parts = self.package.related_parts(f"{kind}s")
        return self.package.xml(parts[0]) if parts else None

    @staticmethod
    def _note_elements(root: etree._Element, kind: str) -> list[etree._Element]:
        """The notes of a notes part, without the separators."""
        return [n for n in root.iter(qn(f"w:{kind}")) if n.get(qn("w:type")) in (None, "normal")]

    @staticmethod
    def _note_text(note: etree._Element) -> str:
        return "\n".join(paragraph_text(p) for p in note.iter(qn("w:p"))).strip()

    def _notes_in(self, root: etree._Element, story_kind: str) -> list[dict]:
        """``[{"id", "text"}]`` for each note of a footnotes or endnotes part."""
        kind = story_kind.rstrip("s")
        return [{"id": n.get(qn("w:id")), "text": self._note_text(n)} for n in self._note_elements(root, kind)]

    def _note_references(self, p: etree._Element) -> list[dict]:
        """The footnote and endnote marks in a paragraph, with their text."""
        found = []
        for run in iter_runs(p):
            for kind in NOTE_KINDS:
                for reference in run.iter(qn(f"w:{kind}Reference")):
                    note_id = reference.get(qn("w:id"))
                    note = self._find_note(kind, note_id)
                    found.append({"kind": kind, "id": note_id, "text": self._note_text(note) if note is not None else ""})
        return found

    def _find_note(self, kind: str, note_id: str) -> etree._Element | None:
        root = self._notes_root(kind)
        if root is None:
            return None
        return next((n for n in self._note_elements(root, kind) if n.get(qn("w:id")) == str(note_id)), None)

    # --------------------------------------------------------------- creating

    def _ensure_notes_root(self, kind: str) -> etree._Element:
        """The footnotes or endnotes part, created with Word's separators."""
        root = self._notes_root(kind)
        if root is not None:
            return root
        separators = "".join(
            f'<w:{kind} w:type="{name}" w:id="{note_id}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" '
            f'w:lineRule="auto"/></w:pPr><w:r><w:{name}/></w:r></w:p></w:{kind}>'
            for note_id, name in _SEPARATORS
        )
        name = self.package.add_part(
            self._part_beside_main(f"{kind}s.xml"),
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:{kind}s xmlns:w="{W_NS}">{separators}</w:{kind}s>',
            _NOTES_TYPES[kind],
        )
        self.package.relate(self.package.main_document_part, f"{kind}s", name)
        settings = self._settings_root()
        if settings.find(qn(f"w:{kind}Pr")) is None:
            properties = el(f"w:{kind}Pr", None, *(el(f"w:{kind}", {"w:id": note_id}) for note_id, _ in _SEPARATORS))
            earlier = SETTINGS_BEFORE_FOOTNOTE_PR + (["w:footnotePr"] if kind == "endnote" else [])
            insert_after_any(settings, properties, earlier)
        self._bind()
        return self.package.xml(name)

    def _next_note_id(self, root: etree._Element, kind: str) -> str:
        ids = [int(n.get(qn("w:id"))) for n in root.iter(qn(f"w:{kind}"))
               if re.fullmatch(r"-?\d+", n.get(qn("w:id"), ""))]
        return str(max([0, *ids]) + 1)

    def add_note(self, ref: str, text: str, match: str | None = None, kind: str = "footnote") -> str:
        """Add a footnote (or endnote, with ``kind="endnote"``) whose text is
        ``text`` (each line a paragraph). Its number goes right after the
        first occurrence of ``match`` in block ``ref``, or at the end of the
        block's (last) paragraph. Returns the note id, which
        :meth:`delete_note` takes; the note's text can be changed like any
        other through the notes part name (see :meth:`stories`).

        The mark and the note text are tracked insertions when tracking is on."""
        kind = _kind(kind)
        if not str(text).strip():
            raise DocumentError("A note needs text.")
        paragraphs = [p for p in self._paragraphs_for(ref) if self._attached(p)]
        if not paragraphs:
            raise DocumentError(f"{ref} is not a block of the body. Notes can only be added to body text.")
        if match:
            p = next((p for p in paragraphs if match in paragraph_text(p)), None)
            if p is None:
                raise DocumentError(f"Text {match!r} was not found in {ref}. Search for the exact text first.")
            start = paragraph_text(p).index(match)
            anchor = isolate_range(p, start, start + len(match))[-1]
            if anchor.getparent().tag == HYPERLINK:
                anchor = anchor.getparent()  # the mark goes after the link, not inside it
            parent, index = anchor.getparent(), anchor.getparent().index(anchor) + 1
        else:
            parent = paragraphs[-1]
            index = len(parent)

        root = self._ensure_notes_root(kind)
        note_id = self._next_note_id(root, kind)
        text_style = self._style_id(f"{kind} text")
        mark_style = self._style_id(f"{kind} reference", "character")

        note = el(f"w:{kind}", {"w:id": note_id})
        with self.edits.untracked():  # a new note is not a revision of anything
            root.append(note)
        for offset, line in enumerate(_lines(str(text).strip())):
            p = el("w:p", None, el("w:pPr", None, el("w:pStyle", {"w:val": text_style})))
            if offset == 0:
                p.append(el("w:r", None, el("w:rPr", None, el("w:rStyle", {"w:val": mark_style})), el(f"w:{kind}Ref")))
                line = " " + line
            if line:
                p.append(make_run(line))
            self.edits.insert_block(note, offset, p)

        mark = el("w:r", None, el("w:rPr", None, el("w:rStyle", {"w:val": mark_style})),
                  el(f"w:{kind}Reference", {"w:id": note_id}))
        self.edits.insert_run(parent, index, mark)
        return note_id

    def delete_note(self, note_id: str, kind: str = "footnote") -> None:
        """Delete a footnote or endnote: its mark in the body and, unless the
        deletion is tracked (then rejecting it brings the mark back), the note."""
        kind = _kind(kind)
        note_id = str(note_id).strip()
        tag = qn(f"w:{kind}Reference")
        live = [r for p in self._body.iter(qn("w:p")) for r in iter_runs(p)
                if any(ref.get(qn("w:id")) == note_id for ref in r.iter(tag))]
        if not live:
            raise DocumentError(f"No {kind} with id {note_id!r}. read_headers_footnotes lists the notes and their ids.")
        for run in live:
            self.edits.delete_run(run)
        still_referenced = any(ref.get(qn("w:id")) == note_id for ref in self._body.iter(tag))
        note = self._find_note(kind, note_id)
        if note is not None and not still_referenced:
            note.getparent().remove(note)  # the note's container, like a part: nothing refers to it now
