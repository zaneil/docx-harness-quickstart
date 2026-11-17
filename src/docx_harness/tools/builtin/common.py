"""Helpers shared by the built-in tool groups: tidying what models send, and
describing the blocks around an edit so the model can see where it landed."""

from __future__ import annotations

import re
from typing import Iterable

from ...document import WordDocument
from ...document.builtin_styles import ALL as BUILTIN_STYLES

_BULLET = re.compile(r"^\s*[-*•–]\s+")
_HEADING = re.compile(r"^\s*(#{1,3})\s+(.*)$")
_NUMBER = re.compile(r"^\s*\d+[.)]\s+")
_MD_LINK = re.compile(r"\[([^\[\]]+)\]\(((?:https?://|mailto:)[^\s()]+)\)")
#: Well-known styles the document layer adds to a file the first time they are used.
CREATABLE_STYLES = sorted({style.name.lower() for style in BUILTIN_STYLES if style.type == "paragraph"})


def as_list(value) -> list:
    """Models often send one id where a list is expected; accept both."""
    if isinstance(value, (str, int)):
        return [value]
    return list(value or [])


# ------------------------------------------------------------------ outline


def outline_line(entry: dict) -> str:
    """One line of the outline: id, style (and list level) and a preview."""
    label = entry.get("style") or entry["type"]
    if entry.get("list", {}).get("level"):
        label += f", level {entry['list']['level']}"
    return f"{entry['id']} [{label}] {entry['preview']}"


def around(doc: WordDocument, touched: Iterable[str], *, heading: str = "Around the change:",
           mark: bool = True, max_shown: int = 4, preview_chars: int = 60) -> str:
    """The outline from the block before the first ``touched`` block to the
    block after the last one, touched blocks marked with ``*`` (unless
    ``mark`` is False). Long runs of
    touched blocks are shortened so a long draft does not echo every paragraph."""
    order = doc.block_ids()
    touched = set(touched)
    positions = [i for i, block_id in enumerate(order) if block_id in touched]
    if not positions:
        return ""
    low, high = max(positions[0] - 1, 0), min(positions[-1] + 1, len(order) - 1)
    entries = doc.outline(order[low], high - low + 1, preview_chars)
    lines = [("* " if mark and e["id"] in touched else "  ") + outline_line(e) for e in entries]
    # Keep the context lines and the first and last touched ones; elide the middle.
    lead = 1 if order[low] not in touched else 0
    tail = 1 if order[high] not in touched else 0
    inner = len(lines) - lead - tail
    if inner > max_shown:
        keep_front = lead + max_shown - 1
        hidden = inner - max_shown
        lines = [*lines[:keep_front], f"  … {hidden} more", *lines[len(lines) - tail - 1:]]
    return heading + "\n" + "\n".join(lines)


def with_context(summary: str, doc: WordDocument, touched: Iterable[str]) -> str:
    context = around(doc, touched)
    return f"{summary} {context}" if context else summary


# --------------------------------------------------------------- paragraphs


def _levels(lines: list[str]) -> list[int]:
    """Nesting level of each line of a markdown-style list, from its indentation."""
    indents = [len(line.expandtabs(4)) - len(line.expandtabs(4).lstrip()) for line in lines]
    base = min(indents)
    step = min((i - base for i in indents if i > base), default=2)
    return [min((i - base) // max(step, 1), 8) for i in indents]


def split_links(text: str) -> tuple[str, list[tuple[str, str]]]:
    """Plain text and the (label, address) pairs of the markdown links in it."""
    found = [(m.group(1), m.group(2)) for m in _MD_LINK.finditer(text)]
    return _MD_LINK.sub(lambda m: m.group(1), text), found


def _item(text: str, style: str | None, level: int | None) -> dict:
    return {"text": text, "style": style, **({"level": level} if level else {})}


def normalize_paragraphs(items: list, available_styles: set[str]) -> list[dict]:
    """Tidy what models commonly send: markdown-style lists crammed into one
    paragraph, list markers typed in front of list-styled text, and blank-line
    separated paragraphs. ``available_styles`` are lower-cased style names."""
    out: list[dict] = []
    for item in items:
        if isinstance(item, str):
            item = {"text": item}
        text = str(item.get("text", ""))
        style = item.get("style") or None
        level = item.get("level")
        lines = [line for line in text.split("\n") if line.strip()]
        if len(lines) > 1 and all(line.strip().startswith("|") for line in lines):
            raise ValueError(
                "Paragraph text contains a markdown table. Nothing was inserted. "
                "Use insert_table with rows as a list of lists instead, and resend the other paragraphs."
            )
        heading = _HEADING.match(text) if len(lines) == 1 else None
        if heading and f"heading {len(heading.group(1))}" in available_styles:
            out.append({"text": heading.group(2).strip(), "style": style or f"Heading {len(heading.group(1))}"})
            continue
        is_list_style = bool(style) and "list" in style.lower()
        handled = False
        for marker, auto_style in ((_BULLET, "List Bullet"), (_NUMBER, "List Number")):
            if lines and all(marker.match(line) for line in lines) and (len(lines) > 1 or is_list_style):
                target = style if is_list_style else auto_style
                if is_list_style or auto_style.lower() in available_styles:
                    nesting = _levels(lines)
                    out.extend(_item(marker.sub("", line).strip(), target, (level or 0) + depth)
                               for line, depth in zip(lines, nesting))
                    handled = True
                break
        if handled:
            continue
        if "\n\n" in text:
            out.extend(_item(part.strip(), style, level) for part in re.split(r"\n\s*\n", text) if part.strip())
        else:
            out.append(_item(text, style, level))
    return out
