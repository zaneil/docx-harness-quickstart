"""Low-level OOXML helpers: namespaces, element builders and run/text mapping.

Nothing in this module knows about blocks, tools or agents. It only knows how
WordprocessingML represents text.
"""

from __future__ import annotations

import copy
from typing import Iterator

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NSMAP = {"w": W_NS, "r": R_NS}


def qn(tag: str) -> str:
    """Turn ``"w:p"`` into the Clark-notation name lxml uses."""
    prefix, _, local = tag.partition(":")
    return f"{{{NSMAP[prefix]}}}{local}"


P = qn("w:p")
R = qn("w:r")
T = qn("w:t")
TAB = qn("w:tab")
BR = qn("w:br")
CR = qn("w:cr")
PPR = qn("w:pPr")
RPR = qn("w:rPr")
TBL = qn("w:tbl")
TR = qn("w:tr")
TC = qn("w:tc")
SDT = qn("w:sdt")
SECTPR = qn("w:sectPr")
BODY = qn("w:body")
DEL = qn("w:del")
INS = qn("w:ins")
MOVE_FROM = qn("w:moveFrom")
XML_SPACE = f"{{{XML_NS}}}space"

# Child order required by the schema (CT_RPr / CT_PPr). Word can refuse files
# whose property elements are out of order, so inserts respect these lists.
RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
    "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
    "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
    "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
    "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath",
]
PPR_ORDER = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
    "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs",
    "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
]


def el(tag: str, attrs: dict[str, str] | None = None, *children: etree._Element) -> etree._Element:
    """Build an element: ``el("w:jc", {"w:val": "center"})``."""
    element = etree.Element(qn(tag), nsmap={"w": W_NS})
    for key, value in (attrs or {}).items():
        element.set(qn(key), str(value))
    for child in children:
        element.append(child)
    return element


def ensure_child(parent: etree._Element, tag: str, order: list[str]) -> etree._Element:
    """Return ``parent``'s ``tag`` child, creating it at its schema position."""
    existing = parent.find(qn(tag))
    if existing is not None:
        return existing
    new = el(tag)
    later = {qn(t) for t in order[order.index(tag) + 1:]}
    for index, child in enumerate(parent):
        if child.tag in later:
            parent.insert(index, new)
            return new
    parent.append(new)
    return new


def remove_child(parent: etree._Element, tag: str) -> None:
    existing = parent.find(qn(tag))
    if existing is not None:
        parent.remove(existing)


def get_or_add_ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(PPR)
    if ppr is None:
        ppr = el("w:pPr")
        p.insert(0, ppr)
    return ppr


def get_or_add_rpr(r: etree._Element) -> etree._Element:
    rpr = r.find(RPR)
    if rpr is None:
        rpr = el("w:rPr")
        r.insert(0, rpr)
    return rpr


# --------------------------------------------------------------------------
# Text <-> runs
# --------------------------------------------------------------------------

def _clean(text: str) -> str:
    """Drop characters that are illegal in XML 1.0 (keeps tab and newline)."""
    return "".join(
        ch for ch in text.replace("\r\n", "\n").replace("\r", "\n")
        if ch in "\t\n" or ord(ch) >= 0x20
    )


def append_text(r: etree._Element, text: str) -> None:
    """Append text to a run. Newlines become line breaks, tabs become tabs."""
    buffer: list[str] = []

    def flush() -> None:
        if buffer:
            t = etree.SubElement(r, T)
            t.text = "".join(buffer)
            t.set(XML_SPACE, "preserve")
            buffer.clear()

    for ch in _clean(text):
        if ch == "\n":
            flush()
            etree.SubElement(r, BR)
        elif ch == "\t":
            flush()
            etree.SubElement(r, TAB)
        else:
            buffer.append(ch)
    flush()


def make_run(text: str, rpr: etree._Element | None = None) -> etree._Element:
    """Create a run holding ``text``, optionally copying run properties."""
    r = el("w:r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    append_text(r, text)
    return r


def make_paragraph(text: str = "", style_id: str | None = None) -> etree._Element:
    p = el("w:p")
    if style_id:
        p.append(el("w:pPr", None, el("w:pStyle", {"w:val": style_id})))
    if text:
        p.append(make_run(text))
    return p


def run_text(r: etree._Element) -> str:
    """Visible text of one run (tabs as ``\\t``, breaks as ``\\n``)."""
    parts: list[str] = []
    for child in r:
        if child.tag == T:
            parts.append(child.text or "")
        elif child.tag == TAB:
            parts.append("\t")
        elif child.tag in (BR, CR):
            parts.append("\n")
    return "".join(parts)


def iter_runs(p: etree._Element) -> Iterator[etree._Element]:
    """Yield the live runs of a paragraph in document order.

    Runs inside hyperlinks, content controls and tracked insertions are
    included. Runs inside tracked deletions (``w:del``, ``w:moveFrom``) and in
    nested paragraphs (text boxes) are not.
    """
    for r in p.iter(R):
        ancestor = r.getparent()
        live = True
        while ancestor is not None and ancestor is not p:
            if ancestor.tag in (DEL, MOVE_FROM, P):
                live = False
                break
            ancestor = ancestor.getparent()
        if live:
            yield r


def paragraph_text(p: etree._Element) -> str:
    return "".join(run_text(r) for r in iter_runs(p))


def _trim(r: etree._Element, start: int, end: int | None) -> None:
    """Keep only the text in ``[start, end)`` of a run, in place."""
    pos = 0
    for child in list(r):
        if child.tag == RPR:
            continue
        if child.tag == T:
            text = child.text or ""
            lo = max(start - pos, 0)
            hi = len(text) if end is None else min(end - pos, len(text))
            if hi <= lo:
                r.remove(child)
            else:
                child.text = text[lo:hi]
                child.set(XML_SPACE, "preserve")
            pos += len(text)
        elif child.tag in (TAB, BR, CR):
            if pos < start or (end is not None and pos >= end):
                r.remove(child)
            pos += 1
        elif start > 0:
            # Non-text content (drawings, field chars) stays with the left half.
            r.remove(child)


def split_run(r: etree._Element, offset: int) -> tuple[etree._Element, etree._Element]:
    """Split a run at a text offset. Both halves keep the run's formatting."""
    right = copy.deepcopy(r)
    right.tail = None
    _trim(r, 0, offset)
    _trim(right, offset, None)
    r.addnext(right)
    return r, right


def isolate_range(p: etree._Element, start: int, end: int) -> list[etree._Element]:
    """Split runs so that ``[start, end)`` of the paragraph text is covered by
    whole runs, and return those runs."""
    covered: list[etree._Element] = []
    pos = 0
    for r in list(iter_runs(p)):
        length = len(run_text(r))
        r_start, r_end = pos, pos + length
        pos = r_end
        if length == 0 or r_end <= start or r_start >= end:
            continue
        current = r
        if start > r_start:
            _, current = split_run(current, start - r_start)
            r_start = start
        if end < r_end:
            current, _ = split_run(current, end - r_start)
        covered.append(current)
    return covered
