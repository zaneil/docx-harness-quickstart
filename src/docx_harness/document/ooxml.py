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
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
NSMAP = {"w": W_NS, "r": R_NS, "wp": WP_NS, "a": A_NS, "pic": PIC_NS}


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
MOVE_TO = qn("w:moveTo")
MOVE_RANGE_STARTS = (qn("w:moveFromRangeStart"), qn("w:moveToRangeStart"))
MOVE_RANGE_ENDS = (qn("w:moveFromRangeEnd"), qn("w:moveToRangeEnd"))
DEL_TEXT = qn("w:delText")
INSTR_TEXT = qn("w:instrText")
DEL_INSTR_TEXT = qn("w:delInstrText")
TRPR = qn("w:trPr")
TCPR = qn("w:tcPr")
TBLPR = qn("w:tblPr")
TBLPREX = qn("w:tblPrEx")
TBL_GRID = qn("w:tblGrid")
GRID_COL = qn("w:gridCol")
GRID_SPAN = qn("w:gridSpan")
GRID_BEFORE = qn("w:gridBefore")
V_MERGE = qn("w:vMerge")
H_MERGE = qn("w:hMerge")
TBL_HEADER = qn("w:tblHeader")
CELL_INS = qn("w:cellIns")
CELL_DEL = qn("w:cellDel")
CELL_MERGE = qn("w:cellMerge")
SDT_CONTENT = qn("w:sdtContent")
CUSTOM_XML = qn("w:customXml")
RPR_CHANGE = qn("w:rPrChange")
PPR_CHANGE = qn("w:pPrChange")
TRPR_CHANGE = qn("w:trPrChange")
TCPR_CHANGE = qn("w:tcPrChange")
TBLPR_CHANGE = qn("w:tblPrChange")
PROPERTY_CHANGES = (RPR_CHANGE, PPR_CHANGE, TRPR_CHANGE, TCPR_CHANGE, TBLPR_CHANGE)
AUTHOR = qn("w:author")
HYPERLINK = qn("w:hyperlink")
DRAWING = qn("w:drawing")
FLD_SIMPLE = qn("w:fldSimple")
R_ID = qn("r:id")
XML_SPACE = f"{{{XML_NS}}}space"

# Child order required by the schema (CT_RPr / CT_PPr). Word can refuse files
# whose property elements are out of order, so inserts respect these lists.
RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps",
    "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint",
    "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden", "w:color", "w:spacing",
    "w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u", "w:effect",
    "w:bdr", "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
    "w:eastAsianLayout", "w:specVanish", "w:oMath", "w:rPrChange",
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
# Table properties. w:tcPr and w:tblPr are sequences; w:trPr is an unordered
# group followed by the revision marks, so only its tail is listed.
TCPR_ORDER = [
    "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd",
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
    "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
]
TBLPR_ORDER = [
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
    "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd",
    "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
    "w:tblDescription", "w:tblPrChange",
]
TRPR_ORDER = ["w:gridBefore", "w:tblHeader", "w:ins", "w:del", "w:trPrChange"]
# The children of w:settings (CT_Settings) that come before w:trackRevisions.
SETTINGS_BEFORE_TRACK_REVISIONS = [
    "w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation",
    "w:removeDateAndTime", "w:doNotDisplayPageBoundaries", "w:displayBackgroundShape",
    "w:printPostScriptOverText", "w:printFractionalCharacterWidth", "w:printFormsData",
    "w:embedTrueTypeFonts", "w:embedSystemFonts", "w:saveSubsetFonts", "w:saveFormsData",
    "w:mirrorMargins", "w:alignBordersAndEdges", "w:bordersDoNotSurroundHeader",
    "w:bordersDoNotSurroundFooter", "w:gutterAtTop", "w:hideSpellingErrors",
    "w:hideGrammaticalErrors", "w:activeWritingStyle", "w:proofState", "w:formsDesign",
    "w:attachedTemplate", "w:linkStyles", "w:stylePaneFormatFilter", "w:stylePaneSortMethod",
    "w:documentType", "w:mailMerge", "w:revisionView",
]
TRACK_REVISIONS = qn("w:trackRevisions")
# ...and those that come before w:evenAndOddHeaders.
SETTINGS_BEFORE_EVEN_AND_ODD = SETTINGS_BEFORE_TRACK_REVISIONS + [
    "w:trackRevisions", "w:doNotTrackMoves", "w:doNotTrackFormatting", "w:documentProtection",
    "w:autoFormatOverride", "w:styleLockTheme", "w:styleLockQFSet", "w:defaultTabStop",
    "w:autoHyphenation", "w:consecutiveHyphenLimit", "w:hyphenationZone", "w:doNotHyphenateCaps",
    "w:showEnvelope", "w:summaryLength", "w:clickAndTypeStyle", "w:defaultTableStyle",
]
# Children of w:sectPr in schema order (the header and footer references come first).
SECTPR_ORDER = [
    "w:headerReference", "w:footerReference", "w:footnotePr", "w:endnotePr", "w:type", "w:pgSz",
    "w:pgMar", "w:paperSrc", "w:pgBorders", "w:lnNumType", "w:pgNumType", "w:cols", "w:formProt",
    "w:vAlign", "w:noEndnote", "w:titlePg", "w:textDirection", "w:bidi", "w:rtlGutter", "w:docGrid",
    "w:printerSettings", "w:sectPrChange",
]


def el(tag: str, attrs: dict[str, str] | None = None, *children: etree._Element) -> etree._Element:
    """Build an element: ``el("w:jc", {"w:val": "center"})``."""
    prefixes = {"w", tag.partition(":")[0], *(key.partition(":")[0] for key in attrs or {})}
    element = etree.Element(qn(tag), nsmap={prefix: NSMAP[prefix] for prefix in prefixes})
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
    return insert_in_order(parent, el(tag), order)


def insert_in_order(parent: etree._Element, new: etree._Element, order: list[str]) -> etree._Element:
    """Insert ``new`` into ``parent`` before the first child the schema
    ``order`` puts after it."""
    names = [qn(t) for t in order]
    later = set(names[names.index(new.tag) + 1:])
    for index, child in enumerate(parent):
        if child.tag in later:
            parent.insert(index, new)
            return new
    parent.append(new)
    return new


def insert_after_any(parent: etree._Element, new: etree._Element, earlier: list[str]) -> None:
    """Insert ``new`` after the last child of ``parent`` whose tag is in
    ``earlier`` (the tags the schema puts before it), or first if there is none."""
    tags = {qn(t) for t in earlier}
    index = 0
    for position, child in enumerate(parent):
        if child.tag in tags:
            index = position + 1
    parent.insert(index, new)


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


def get_or_add_props(element: etree._Element, tag: str) -> etree._Element:
    """The property element (``w:tcPr``, ``w:trPr``, ``w:tblPr``) of a cell,
    row or table, created in its place at the front when missing."""
    props = element.find(qn(tag))
    if props is None:
        props = el(tag)
        element.insert(1 if len(element) and element[0].tag == TBLPREX else 0, props)
    return props


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
        rpr = copy.deepcopy(rpr)
        remove_child(rpr, "w:rPrChange")  # a new run has no formatting history
        r.append(rpr)
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


# --------------------------------------------------------------------------
# Tracked changes (revisions)
# --------------------------------------------------------------------------

def paragraph_mark(p: etree._Element, tag: str) -> etree._Element | None:
    """The ``w:ins`` or ``w:del`` (``tag``) recorded on a paragraph mark."""
    ppr = p.find(PPR)
    rpr = ppr.find(RPR) if ppr is not None else None
    return rpr.find(tag) if rpr is not None else None


def row_mark(tr: etree._Element, tag: str) -> etree._Element | None:
    """The ``w:ins`` or ``w:del`` (``tag``) recorded on a table row."""
    trpr = tr.find(TRPR)
    return trpr.find(tag) if trpr is not None else None


def is_deleted_block(element: etree._Element) -> bool:
    """True when a tracked deletion (or a tracked move to somewhere else) has
    removed this whole paragraph, table row, table or content control. It is
    still in the file until the change is accepted."""
    if element.tag == P:
        gone = paragraph_mark(element, DEL) is not None or paragraph_mark(element, MOVE_FROM) is not None
        return gone and next(iter_runs(element), None) is None
    if element.tag == TR:
        return row_mark(element, DEL) is not None
    if element.tag == TBL:
        rows = element.findall(TR)
        return bool(rows) and all(is_deleted_block(tr) for tr in rows)
    if element.tag == SDT:
        paragraphs = list(element.iter(P))
        return bool(paragraphs) and all(is_deleted_block(p) for p in paragraphs)
    return False


def strip_revisions(element: etree._Element) -> None:
    """Make a detached copy read as if every revision in it were accepted:
    deleted content is dropped, insertions are unwrapped, marks are removed."""
    for node in list(element.iter(DEL, MOVE_FROM, *PROPERTY_CHANGES, *MOVE_RANGE_STARTS, *MOVE_RANGE_ENDS)):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
    for node in list(element.iter(INS, MOVE_TO)):
        if node.getparent() is not None:
            unwrap(node)
    for mark in list(element.iter(CELL_INS, CELL_DEL)):
        props = mark.getparent()
        cell = props.getparent() if props is not None else None
        if mark.tag == CELL_DEL and cell is not None and cell.getparent() is not None:
            cell.getparent().remove(cell)
        elif props is not None:
            props.remove(mark)


def unwrap(node: etree._Element) -> None:
    """Replace ``node`` with its children."""
    parent = node.getparent()
    index = parent.index(node)
    for child in list(node):
        parent.insert(index, child)
        index += 1
    parent.remove(node)
