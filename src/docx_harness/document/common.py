"""Names shared by ``WordDocument`` and its mixins."""

from __future__ import annotations

from lxml import etree

from .ooxml import (
    SDT,
    TBL,
    P,
    is_deleted_block,
    make_run,
    qn,
    run_text,
)

BLOCK_TAGS = (P, TBL, SDT)
_VAL = qn("w:val")
PAGE_WIDTH_TWIPS = 9360  # 6.5in text width on Letter with 1in margins


class DocumentError(ValueError):
    """A document operation could not be performed (bad id, bad argument)."""


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
