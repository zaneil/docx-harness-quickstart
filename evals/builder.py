"""Turn a task's ``input:`` into the starting document.

Each step is a mapping with one *kind* key (``paragraph``, ``table``...) and,
for some kinds, option keys beside it:

    input:
      - heading: Status report
      - paragraph: A quoted line.
        style: Quote
      - table: [[Area, Owner], [Docs, Tom]]

Everything goes through the public ``WordDocument`` methods, so an input can
hold whatever the document layer can write. To add a kind, write a function
decorated with ``@step("name", "option", ...)``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from docx_harness import WordDocument

EVALS_DIR = Path(__file__).parent

#: kind -> (function, option keys allowed beside the kind key)
STEPS: dict[str, tuple[Callable, frozenset[str]]] = {}


def step(name: str, *options: str):
    def register(fn: Callable) -> Callable:
        STEPS[name] = (fn, frozenset(options))
        return fn
    return register


def _kind(item: Any, known: dict | set, what: str = "input step") -> str:
    if not isinstance(item, dict):
        raise ValueError(f"each {what} must be a mapping like '- paragraph: text', got {item!r}")
    kinds = [key for key in item if key in known]
    if len(kinds) != 1:
        raise ValueError(f"{what} {item!r} must have exactly one of: {', '.join(sorted(known))}")
    return kinds[0]


def validate(spec: Any) -> None:
    """Raise ValueError when ``spec`` is not a usable ``input:`` value."""
    if spec is None:
        return
    if isinstance(spec, dict):
        if set(spec) != {"file"}:
            raise ValueError("'input' must be a list of steps or {file: path-to.docx}")
        if not _fixture(spec["file"]).is_file():
            raise ValueError(f"input file {spec['file']!r} was not found (paths are relative to evals/)")
        return
    if not isinstance(spec, list):
        raise ValueError("'input' must be a list of steps or {file: path-to.docx}")
    for item in spec:
        kind = _kind(item, STEPS)
        extra = set(item) - {kind} - STEPS[kind][1]
        if extra:
            allowed = ", ".join(sorted(STEPS[kind][1])) or "none"
            raise ValueError(f"input step {kind!r} does not take {', '.join(sorted(extra))} (options: {allowed})")
        if kind == "tracked":
            for inner in (item["tracked"] or {}).get("steps", []):
                _kind(inner, _TRACKED, "tracked step")


def _fixture(path: str) -> Path:
    candidate = Path(path)
    return candidate if candidate.is_absolute() else EVALS_DIR / candidate


def build(spec: Any) -> WordDocument:
    """The starting document for ``spec`` (a blank one for None)."""
    validate(spec)
    if isinstance(spec, dict):
        return WordDocument.open(_fixture(spec["file"]))
    doc = WordDocument.open()
    for item in spec or []:
        kind = _kind(item, STEPS)
        STEPS[kind][0](doc, item[kind], {k: v for k, v in item.items() if k != kind})
    return doc


def block_with(doc: WordDocument, text: str) -> str:
    """Id of the first body paragraph containing ``text``."""
    for hit in doc.search(text, case_sensitive=True):
        if "[" not in hit["ref"] and "/" not in hit["ref"]:
            return hit["ref"]
    raise ValueError(f"input refers to text {text!r}, which is not in a body paragraph built so far")


# ------------------------------------------------------------------- kinds

@step("title")
def _title(doc, value, options):
    doc.insert_paragraph(str(value), "Title")


@step("heading", "level")
def _heading(doc, value, options):
    doc.insert_paragraph(str(value), f"Heading {int(options.get('level', 1))}")


@step("paragraph", "style", "level")
def _paragraph(doc, value, options):
    doc.insert_paragraph(str(value), options.get("style"), level=options.get("level"))


@step("bullets")
def _bullets(doc, value, options):
    for text in value:
        doc.insert_paragraph(str(text), "List Bullet")


@step("numbered")
def _numbered(doc, value, options):
    for text in value:
        doc.insert_paragraph(str(text), "List Number")


@step("table", "header", "merge")
def _table(doc, value, options):
    """``merge`` is a list of ``[row, col, end_row, end_col]`` rectangles."""
    table = doc.insert_table([[str(cell) for cell in row] for row in value], header=bool(options.get("header", True)))
    for row, col, end_row, end_col in options.get("merge") or []:
        doc.merge_cells(table, row, col, end_row, end_col)


@step("page_break")
def _page_break(doc, value, options):
    doc.insert_page_break()


@step("header")
def _header(doc, value, options):
    doc.set_header(str(value))


@step("footer")
def _footer(doc, value, options):
    doc.set_footer(str(value))


@step("section_break")
def _section_break(doc, value, options):
    """A new section starts after the paragraph containing ``value``."""
    doc.insert_section_break(block_with(doc, str(value)))


@step("section_header")
def _section_header(doc, value, options):
    """``{section, text}``: that section's own header."""
    doc.set_header(str(value["text"]), section=int(value["section"]))


@step("footnote")
def _footnote(doc, value, options):
    """``{after, text}``: a footnote after existing text."""
    doc.add_note(block_with(doc, value["after"]), value["text"], match=value["after"])


@step("link")
def _link(doc, value, options):
    """``{text, url}``: link existing text."""
    doc.add_hyperlink(block_with(doc, value["text"]), value["url"], match=value["text"])


@step("format")
def _format(doc, value, options):
    """``{text, bold: true, color: ...}``: format existing text."""
    fmt = {k: v for k, v in value.items() if k != "text"}
    doc.format_text(block_with(doc, value["text"]), value["text"], **fmt)


# Edits another author left as pending revisions.
_TRACKED = {"replace", "insert", "delete"}


@step("tracked")
def _tracked(doc, value, options):
    """``{author, steps}`` where each step is ``replace: [old, new]``,
    ``insert: text`` (with ``after: text of the paragraph above``) or
    ``delete: text of the paragraph to delete``."""
    previous = doc.edits.author
    doc.edits.author = str(value.get("author") or "Reviewer")
    doc.track_changes = True
    try:
        for item in value.get("steps", []):
            kind = _kind(item, _TRACKED, "tracked step")
            if kind == "replace":
                old, new = item["replace"]
                if doc.replace_text(str(old), str(new)) == 0:
                    raise ValueError(f"tracked replace: {old!r} is not in the document built so far")
            elif kind == "insert":
                after = block_with(doc, item["after"]) if item.get("after") else "end"
                doc.insert_paragraph(str(item["insert"]), item.get("style"), after=after)
            else:
                doc.delete_block(block_with(doc, item["delete"]))
    finally:
        doc.track_changes = False
        doc.edits.author = previous
