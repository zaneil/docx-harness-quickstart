"""The checks a task can run on the document the model produced.

A check is a function ``(ctx, argument) -> None | str``: None when it passes,
otherwise a short reason. ``ctx`` gives the saved output document, reopened,
and the input document it started from. To add a check, write a function
decorated with ``@check("name")`` and describe it in ``evals/README.md``.

Text is compared after :func:`norm`: curly quotes and long dashes become plain
ones and runs of whitespace one space, so a model is not failed for typography.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from functools import cached_property
from pathlib import Path
from typing import Any, Callable

from docx_harness import WordDocument

CHECKS: dict[str, Callable[["Context", Any], str | None]] = {}

_PUNCTUATION = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"', "–": "-", "—": "-", " ": " "})
_HEADING = re.compile(r"heading (\d)$", re.IGNORECASE)
_CELL = re.compile(r"\[(\d+),(\d+)\]$")


def check(name: str):
    def register(fn):
        CHECKS[name] = fn
        return fn
    return register


def norm(text: Any) -> str:
    text = unicodedata.normalize("NFKC", str(text)).translate(_PUNCTUATION)
    return " ".join(text.split())


def fold(text: Any) -> str:
    """:func:`norm`, case ignored."""
    return norm(text).casefold()


@dataclass
class CheckResult:
    name: str
    argument: Any
    passed: bool
    message: str = ""


class Context:
    """What a check can look at. ``out`` and ``inp`` are shared between checks
    and must not be changed; a check that edits calls :meth:`reopen`."""

    def __init__(self, out_path: str | Path, in_path: str | Path):
        self.out_path, self.in_path = Path(out_path), Path(in_path)

    @cached_property
    def out(self) -> WordDocument:
        return WordDocument.open(self.out_path)

    @cached_property
    def inp(self) -> WordDocument:
        return WordDocument.open(self.in_path)

    def reopen(self) -> WordDocument:
        """A private copy of the output document that may be edited."""
        return WordDocument.open(self.out_path)

    @cached_property
    def blocks(self) -> list[dict]:
        """Every body block of the output, with formatting."""
        return self.out.read(count=10**9, include_formatting=True)["blocks"]

    @cached_property
    def paragraphs(self) -> list[dict]:
        """Body paragraphs of the output that are not in a table."""
        return [b for b in self.blocks if b["type"] == "paragraph"]

    @cached_property
    def tables(self) -> list[dict]:
        return [b for b in self.blocks if b["type"] == "table"]

    @cached_property
    def lines(self) -> list[str]:
        """Normalized text of every non-empty paragraph, table cells included."""
        return [norm(line) for line in self.out.text().split("\n") if line.strip()]

    @cached_property
    def text(self) -> str:
        return " ".join(self.lines)


def run_checks(task_checks: list[tuple[str, Any]], ctx: Context) -> list[CheckResult]:
    results = []
    for name, argument in task_checks:
        try:
            message = CHECKS[name](ctx, argument)
        except Exception as exc:  # noqa: BLE001 - a broken check is a failed check, not a crashed run
            message = f"check could not run: {type(exc).__name__}: {exc}"
        results.append(CheckResult(name, argument, not message, message or ""))
    return results


# ----------------------------------------------------------------- helpers

def _strings(argument: Any) -> list[str]:
    return [str(a) for a in argument] if isinstance(argument, list) else [str(argument)]


def _options(argument: Any, main: str) -> dict:
    """Accept a bare value as shorthand for ``{main: value}``."""
    return dict(argument) if isinstance(argument, dict) else {main: argument}


def _in_range(value: int, spec: Any, what: str) -> str | None:
    """``spec`` is an exact number or ``{min, max}``."""
    if isinstance(spec, dict):
        low, high = spec.get("min"), spec.get("max")
        if (low is not None and value < low) or (high is not None and value > high):
            bounds = " and ".join(f"{k} {v}" for k, v in (("at least", low), ("at most", high)) if v is not None)
            return f"{what}: found {value}, expected {bounds}"
        return None
    return None if value == int(spec) else f"{what}: found {value}, expected {spec}"


def _missing(wanted: list[str], haystack: str, case: bool) -> list[str]:
    haystack = norm(haystack) if case else fold(haystack)
    return [w for w in wanted if (norm(w) if case else fold(w)) not in haystack]


def _heading_level(block: dict) -> int | None:
    """1-9 for a heading style, 0 for Title, None for anything else."""
    style = block.get("style") or ""
    if style.lower() == "title":
        return 0
    match = _HEADING.match(style)
    return int(match.group(1)) if match else None


def _matching(paragraphs: list[dict], text: str) -> list[dict]:
    """Paragraphs whose text is ``text``; failing that, those containing it."""
    wanted = fold(text)
    exact = [p for p in paragraphs if fold(p["text"]) == wanted]
    return exact or [p for p in paragraphs if wanted in fold(p["text"])]


# -------------------------------------------------------------------- text

@check("text_contains")
def text_contains(ctx, argument):
    """Every string appears in the body. ``{text: [...], case: false}`` ignores case."""
    options = _options(argument, "text")
    missing = _missing(_strings(options["text"]), ctx.text, options.get("case", True))
    return f"missing from the text: {missing}" if missing else None


@check("text_contains_any")
def text_contains_any(ctx, argument):
    """At least one of the strings appears in the body (case ignored)."""
    wanted = _strings(argument)
    return f"none of {wanted} is in the text" if len(_missing(wanted, ctx.text, False)) == len(wanted) else None


@check("text_absent")
def text_absent(ctx, argument):
    """None of the strings appears in the body. ``{text: [...], case: false}`` ignores case."""
    options = _options(argument, "text")
    wanted = _strings(options["text"])
    missing = _missing(wanted, ctx.text, options.get("case", True))
    present = [w for w in wanted if w not in missing]
    return f"still in the text: {present}" if present else None


@check("regex")
def regex(ctx, argument):
    """The pattern matches somewhere in the body (searched paragraph by
    paragraph joined with newlines). ``{pattern, min, max}`` bounds the number of matches."""
    options = _options(argument, "pattern")
    found = len(re.findall(options["pattern"], "\n".join(ctx.lines), re.MULTILINE))
    if "min" in options or "max" in options:
        return _in_range(found, {"min": options.get("min"), "max": options.get("max")},
                         f"matches of {options['pattern']!r}")
    return None if found else f"no match for {options['pattern']!r}"


@check("word_count")
def word_count(ctx, argument):
    """Words in the body: a number or ``{min, max}``."""
    return _in_range(len(ctx.text.split()), argument, "word count")


@check("section_words")
def section_words(ctx, argument):
    """Words in the paragraphs under ``heading``, up to the next heading:
    ``{heading, min, max}``. The heading itself is not counted."""
    count, inside = 0, False
    for block in ctx.blocks:
        if block["type"] == "paragraph" and _heading_level(block) is not None:
            if inside:
                break
            inside = fold(block["text"]) == fold(argument["heading"])
        elif inside:
            cells = [c for row in block.get("rows", []) for c in row]
            count += len(" ".join([block.get("text", ""), *cells]).split())
    if not inside and not count:
        return f"no heading {argument['heading']!r}"
    return _in_range(count, {"min": argument.get("min"), "max": argument.get("max")},
                     f"words under {argument['heading']!r}")


@check("unchanged")
def unchanged(ctx, argument):
    """Each string is still the exact text of a paragraph or table cell."""
    lines = set(ctx.lines)
    changed = [s for s in _strings(argument) if norm(s) not in lines]
    return f"no longer present as written: {changed}" if changed else None


@check("unchanged_except")
def unchanged_except(ctx, argument):
    """Every paragraph of the input is still there word for word, except
    those containing one of the listed strings. Use ``[]`` for 'nothing may change'."""
    exempt = [norm(s) for s in _strings(argument or [])]
    lines = set(ctx.lines)
    changed = [
        line for line in (norm(l) for l in ctx.inp.text().split("\n") if l.strip())
        if line not in lines and not any(e in line for e in exempt)
    ]
    return f"{len(changed)} paragraph(s) that should be untouched changed, e.g. {changed[0]!r}" if changed else None


@check("order")
def order(ctx, argument):
    """The strings appear in the body in this order."""
    text, position = fold(ctx.text), 0
    for wanted in _strings(argument):
        found = text.find(fold(wanted), position)
        if found < 0:
            where = "is not in the text" if fold(wanted) not in text else "comes too early"
            return f"{wanted!r} {where}"
        position = found + len(fold(wanted))
    return None


# --------------------------------------------------------------- structure

@check("heading")
def heading(ctx, argument):
    """A heading containing ``text`` exists, at ``level`` if given (0 is the
    Title style). ``{level: 0}`` alone asks for any title."""
    options = _options(argument, "text")
    level, text = options.get("level"), options.get("text", "")
    for block in _matching([p for p in ctx.paragraphs if _heading_level(p) is not None], text):
        if level is None or _heading_level(block) == int(level):
            return None
    at = f" at level {level}" if level is not None else ""
    return f"no heading{at} containing {text!r}"


@check("heading_count")
def heading_count(ctx, argument):
    """Number of headings: a number or ``{min, max}``, counting only ``level`` if given."""
    options = argument if isinstance(argument, dict) else {}
    level = options.get("level")
    levels = [_heading_level(p) for p in ctx.paragraphs]
    found = sum(1 for l in levels if l is not None and (l == int(level) if level is not None else l > 0))
    spec = {k: options[k] for k in ("min", "max") if k in options} if isinstance(argument, dict) else argument
    return _in_range(found, spec, "headings" if level is None else f"level {level} headings")


@check("style_of")
def style_of(ctx, argument):
    """The paragraph with this ``text`` uses ``style`` (a name or a list of acceptable names)."""
    wanted = {s.lower() for s in _strings(argument["style"])}
    found = _matching(ctx.paragraphs, argument["text"])
    if not found:
        return f"no paragraph containing {argument['text']!r}"
    wrong = [p.get("style") for p in found if (p.get("style") or "").lower() not in wanted]
    return f"paragraph {argument['text']!r} has style {wrong[0]!r}, expected {argument['style']!r}" if wrong else None


@check("block_count")
def block_count(ctx, argument):
    """Body blocks (non-empty paragraphs and tables): a number or ``{min, max}``."""
    found = sum(1 for b in ctx.blocks if b["type"] != "paragraph" or b["text"].strip() or b.get("images"))
    return _in_range(found, argument, "blocks")


@check("list_items")
def list_items(ctx, argument):
    """List paragraphs. ``kind`` (bullet or number) and ``level`` narrow which
    ones count; ``items`` must each be found in one; ``min`` / ``max`` bound how many there are."""
    items = [
        p for p in ctx.paragraphs if "list" in p
        and argument.get("kind") in (None, p["list"]["kind"])
        and argument.get("level") in (None, p["list"]["level"])
    ]
    kind = f"{argument['kind']} " if argument.get("kind") else ""
    missing = [w for w in _strings(argument.get("items", [])) if not any(fold(w) in fold(p["text"]) for p in items)]
    if missing:
        return f"not found as {kind}list items: {missing}"
    if "min" in argument or "max" in argument:
        return _in_range(len(items), {"min": argument.get("min"), "max": argument.get("max")}, f"{kind}list items")
    return None


@check("paragraph_format")
def paragraph_format(ctx, argument):
    """The paragraph with this ``text`` has the given direct layout, e.g. ``alignment: center``."""
    found = _matching(ctx.paragraphs, argument["text"])
    if not found:
        return f"no paragraph containing {argument['text']!r}"
    for key, wanted in argument.items():
        if key != "text" and any(p.get("format", {}).get(key) != wanted for p in found):
            return f"paragraph {argument['text']!r}: {key} is {found[0].get('format', {}).get(key)!r}, expected {wanted!r}"
    return None


# ------------------------------------------------------------------ tables

@check("table_count")
def table_count(ctx, argument):
    """Tables in the body: a number or ``{min, max}``."""
    return _in_range(len(ctx.tables), argument, "tables")


@check("table")
def table(ctx, argument):
    """One table, picked by ``with_text`` (the first table containing it) or
    ``index`` (default 0). Then any of: ``rows`` / ``cols`` (a number or
    ``{min, max}``), ``cells`` (``[{at: [row, col], text: ...}]``), ``has_row``
    (rows that must exist, cell for cell), ``contains`` / ``absent`` (strings
    anywhere in it), ``shaded`` (``[row, col]`` cells with a background color)
    and ``shaded_row`` (a row whose every cell has one)."""
    tables = ctx.tables
    if "with_text" in argument:
        wanted = fold(argument["with_text"])
        tables = [t for t in tables if any(wanted in fold(c) for row in t["rows"] for c in row)]
        if not tables:
            return f"no table containing {argument['with_text']!r}"
        found = tables[0]
    else:
        index = int(argument.get("index", 0))
        if index >= len(tables):
            return f"the document has {len(tables)} table(s); table {index} does not exist"
        found = tables[index]
    rows = [[fold(c) for c in row] for row in found["rows"]]
    cols = max((len(r) for r in rows), default=0)
    for key, value in (("rows", len(rows)), ("cols", cols)):
        if key in argument:
            problem = _in_range(value, argument[key], f"table {key}")
            if problem:
                return problem
    for cell in argument.get("cells", []):
        row, col = cell["at"]
        actual = rows[row][col] if row < len(rows) and col < len(rows[row]) else None
        if actual != fold(cell["text"]):
            return f"cell [{row},{col}] is {found['rows'][row][col] if actual is not None else 'missing'!r}, expected {cell['text']!r}"
    for wanted in argument.get("has_row", []):
        if [fold(c) for c in wanted] not in rows:
            return f"no row reads {wanted}; rows are {found['rows']}"
    flat = " | ".join(c for row in rows for c in row)
    missing = [w for w in _strings(argument.get("contains", [])) if fold(w) not in flat]
    if missing:
        return f"not in the table: {missing}"
    present = [w for w in _strings(argument.get("absent", [])) if fold(w) in flat]
    if present:
        return f"still in the table: {present}"
    shaded = set()
    for cell in found.get("format", {}).get("cells", []):
        match = _CELL.search(cell["cell"])
        if match and cell.get("shading"):
            shaded.add((int(match.group(1)), int(match.group(2))))
    wanted_cells = [tuple(c) for c in argument.get("shaded", [])]
    if "shaded_row" in argument:
        wanted_cells += [(int(argument["shaded_row"]), c) for c in range(cols)]
    plain = [list(c) for c in wanted_cells if c not in shaded]
    return f"cells without shading: {plain}" if plain else None


# ------------------------------------------------- formatting and content

@check("formatting")
def formatting(ctx, argument):
    """Every occurrence of ``text`` in body paragraphs carries the given
    direct character formatting (``bold``, ``italic``, ``underline``,
    ``strike``, ``color``, ``highlight``, ``size_pt``, ``font``). With
    ``only: true`` the rest of those paragraphs must not carry it."""
    wanted_text = argument["text"]
    wanted = {k: v for k, v in argument.items() if k not in ("text", "only")}

    def has(run: dict) -> bool:
        for key, value in wanted.items():
            actual = run.get(key)
            if isinstance(value, bool):
                if bool(actual) != value:
                    return False
            elif str(actual).lower() != str(value).lower():
                return False
        return True

    occurrences = 0
    for paragraph in ctx.paragraphs:
        runs = paragraph.get("runs", [])
        full = "".join(r["text"] for r in runs)
        spans = [(m.start(), m.end()) for m in re.finditer(re.escape(wanted_text), full)]
        if not spans:
            continue
        occurrences += len(spans)
        offset = 0
        for run in runs:
            start, end = offset, offset + len(run["text"])
            offset = end
            inside = any(start < hi and lo < end for lo, hi in spans)
            if inside and not has(run):
                return f"{wanted_text!r} is not formatted {wanted} (found {run})"
            if not inside and argument.get("only") and run["text"].strip() and has(run):
                return f"text other than {wanted_text!r} is also formatted {wanted}: {run['text']!r}"
    return None if occurrences else f"{wanted_text!r} is not in a body paragraph"


@check("link")
def link(ctx, argument):
    """The words ``text``, and only those words, are a hyperlink to ``url``.
    With ``exact: false`` a longer link that contains them also counts."""
    wanted_url = str(argument["url"]).rstrip("/").lower()
    links = [l for p in ctx.paragraphs for l in p.get("links", [])]
    for found in links:
        same = fold(argument["text"]) == fold(found["text"]) or (
            argument.get("exact") is False and fold(argument["text"]) in fold(found["text"]))
        if same and found["url"].rstrip("/").lower() == wanted_url:
            return None
    return f"no link on {argument['text']!r} to {argument['url']}; links are {links}"


def _story(ctx, kind: str, argument) -> str | None:
    text = "\n".join(line for story in ctx.out.stories() if story["kind"] == kind for line in story["paragraphs"])
    missing = _missing(_strings(argument), text, True)
    return f"missing from the {kind}: {missing} (it reads {text!r})" if missing else None


@check("header_contains")
def header_contains(ctx, argument):
    """Every string appears in a page header. Page fields read as {PAGE} and {NUMPAGES}."""
    return _story(ctx, "header", argument)


@check("footer_contains")
def footer_contains(ctx, argument):
    """Every string appears in a page footer. Page fields read as {PAGE} and {NUMPAGES}."""
    return _story(ctx, "footer", argument)


# --------------------------------------------------------------- redlines

@check("revisions")
def revisions(ctx, argument):
    """Pending tracked changes: a number or ``{min, max}``, counting only
    those by ``author`` or of ``type`` (insert, delete, format, move) if given."""
    options = argument if isinstance(argument, dict) else {}
    found = [
        r for r in ctx.out.revisions()
        if options.get("author") in (None, r["author"]) and options.get("type") in (None, r["type"])
    ]
    spec = {k: options[k] for k in ("min", "max") if k in options} if isinstance(argument, dict) else argument
    who = f" by {options['author']}" if options.get("author") else ""
    return _in_range(len(found), spec, f"revisions{who}")


@check("reject_all_restores_input")
def reject_all_restores_input(ctx, argument):
    """Rejecting the tracked changes gives back the input text, so nothing
    was changed without being recorded. ``{author: name}`` rejects only that author's."""
    author = argument.get("author") if isinstance(argument, dict) else None
    doc = ctx.reopen()
    doc.reject_revisions(author=author)
    before = [norm(l) for l in ctx.inp.text().split("\n") if l.strip()]
    after = [norm(l) for l in doc.text().split("\n") if l.strip()]
    if before == after:
        return None
    differing = next((b for b in before if b not in after), None) or next(a for a in after if a not in before)
    return f"after rejecting the revisions the text differs from the input, e.g. at {differing!r}"
