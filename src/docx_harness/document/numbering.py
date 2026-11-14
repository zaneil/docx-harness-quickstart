"""Lists: the numbering part (``word/numbering.xml``) and ``w:numPr``.

A list item is a paragraph with a numbering id (``w:numId``) and a level
(``w:ilvl``), set on the paragraph itself or supplied by its style. The id
names a ``w:num``, which points at a ``w:abstractNum`` holding the look of
each level. Paragraphs sharing a ``w:num`` share one counter, so a list that
starts again at 1 needs a ``w:num`` of its own.
"""

from __future__ import annotations

from lxml import etree

from .ooxml import PPR, PPR_ORDER, W_NS, el, ensure_child, get_or_add_ppr, qn
from .package import parse_xml

VAL = qn("w:val")
NUM = qn("w:num")
ABSTRACT_NUM = qn("w:abstractNum")
NUM_PR = qn("w:numPr")
_NUM_ID = qn("w:numId")
_ABSTRACT_ID = qn("w:abstractNumId")
_ILVL = qn("w:ilvl")
_LVL = qn("w:lvl")
MAX_LEVEL = 8
CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"
EMPTY_PART = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
)
# The definitions this library adds are found again by name.
_NAMES = {"bullet": "docx-harness bullet", "number": "docx-harness number"}
_BULLETS = "•○▪"
_NUMBER_FORMATS = ("decimal", "lowerLetter", "lowerRoman")
_NUMBERED = {"decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman", "decimalZero", "ordinal",
             "cardinalText", "ordinalText"}


def abstract_xml(kind: str, abstract_id: int = 0) -> str:
    """A nine-level list definition, as text without namespace declarations."""
    levels = []
    for level in range(MAX_LEVEL + 1):
        if kind == "bullet":
            fmt, text = "bullet", _BULLETS[level % 3]
        else:
            fmt, text = _NUMBER_FORMATS[level % 3], f"%{level + 1}."
        levels.append(
            f'<w:lvl w:ilvl="{level}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{720 * (level + 1)}" w:hanging="360"/></w:pPr></w:lvl>'
        )
    return (
        f'<w:abstractNum w:abstractNumId="{abstract_id}"><w:multiLevelType w:val="hybridMultilevel"/>'
        f'<w:name w:val="{_NAMES[kind]}"/>{"".join(levels)}</w:abstractNum>'
    )


def _int(value: str | None) -> int | None:
    try:
        return int(value) if value is not None else None
    except ValueError:
        return None


def _child_val(parent: etree._Element | None, tag: str) -> int | None:
    child = parent.find(tag) if parent is not None else None
    return _int(child.get(VAL)) if child is not None else None


# ------------------------------------------------------------------- reading

def num_pr(holder: etree._Element) -> tuple[int | None, int | None]:
    """``(numbering id, level)`` set directly on ``holder`` (a paragraph or a
    style); either is None when not given."""
    ppr = holder.find(PPR)
    pr = ppr.find(NUM_PR) if ppr is not None else None
    return (_child_val(pr, _NUM_ID), _child_val(pr, _ILVL)) if pr is not None else (None, None)


def abstract_of(numbering: etree._Element | None, num_id: int | None) -> etree._Element | None:
    """The ``w:abstractNum`` behind numbering id ``num_id``."""
    if numbering is None or not num_id:
        return None
    for num in numbering.iter(NUM):
        if _int(num.get(_NUM_ID)) == num_id:
            wanted = _child_val(num, _ABSTRACT_ID)
            return next((a for a in numbering.iter(ABSTRACT_NUM) if _int(a.get(_ABSTRACT_ID)) == wanted), None)
    return None


def level_count(abstract: etree._Element) -> int:
    return len(abstract.findall(_LVL))


def kind_of(numbering: etree._Element | None, num_id: int | None, level: int) -> str | None:
    """"bullet" or "number" for a level of a list; None when it has no numbering."""
    abstract = abstract_of(numbering, num_id)
    if abstract is None:
        return None
    lvl = next((l for l in abstract.findall(_LVL) if _int(l.get(_ILVL)) == level), None)
    fmt = lvl.find(qn("w:numFmt")) if lvl is not None else None
    value = fmt.get(VAL) if fmt is not None else "decimal"
    if value == "none":
        return None
    return "bullet" if value == "bullet" else "number"


def start_of(numbering: etree._Element | None, num_id: int, level: int) -> int:
    """The number the first item of a list level gets."""
    if numbering is None:
        return 1
    for num in numbering.iter(NUM):
        if _int(num.get(_NUM_ID)) == num_id:
            for override in num.findall(qn("w:lvlOverride")):
                if _int(override.get(_ILVL)) == level and override.find(qn("w:startOverride")) is not None:
                    return _child_val(override, qn("w:startOverride")) or 0
    abstract = abstract_of(numbering, num_id)
    lvl = next((l for l in abstract.findall(_LVL) if _int(l.get(_ILVL)) == level), None) if abstract is not None else None
    start = _child_val(lvl, qn("w:start"))
    return 1 if start is None else start


# ------------------------------------------------------------------- writing

def ensure_abstract(numbering: etree._Element, kind: str) -> int:
    """The id of this library's bullet or numbered list definition, added to
    the numbering part when it is not there yet."""
    for abstract in numbering.iter(ABSTRACT_NUM):
        name = abstract.find(qn("w:name"))
        if name is not None and name.get(VAL) == _NAMES[kind]:
            return _int(abstract.get(_ABSTRACT_ID)) or 0
    used = [_int(a.get(_ABSTRACT_ID)) or 0 for a in numbering.iter(ABSTRACT_NUM)]
    abstract_id = max(used, default=-1) + 1
    wrapper = parse_xml(f'<w:numbering xmlns:w="{W_NS}">{abstract_xml(kind, abstract_id)}</w:numbering>')
    abstract = wrapper[0]
    # Schema order: picture bullets, abstract definitions, then the w:num list.
    index = next((i for i, child in enumerate(numbering) if child.tag not in (qn("w:numPicBullet"), ABSTRACT_NUM)),
                 len(numbering))
    numbering.insert(index, abstract)
    return abstract_id


def new_num(numbering: etree._Element, abstract_id: int, start: int | None = None, level: int = 0) -> int:
    """Add a ``w:num`` on ``abstract_id`` and return its id. With ``start``
    its ``level`` begins counting there instead of carrying on."""
    num_id = max((_int(n.get(_NUM_ID)) or 0 for n in numbering.iter(NUM)), default=0) + 1
    num = el("w:num", {"w:numId": str(num_id)}, el("w:abstractNumId", {"w:val": str(abstract_id)}))
    if start is not None:
        num.append(el("w:lvlOverride", {"w:ilvl": str(level)}, el("w:startOverride", {"w:val": str(start)})))
    last = None
    for child in numbering:
        if child.tag in (NUM, ABSTRACT_NUM, qn("w:numPicBullet")):
            last = child
    if last is None:
        numbering.insert(0, num)
    else:
        last.addnext(num)
    return num_id


def base_num(numbering: etree._Element, kind: str) -> int:
    """A numbering id on this library's definition for ``kind`` that list
    styles can share: the first one without a restart, or a new one."""
    abstract_id = ensure_abstract(numbering, kind)
    for num in numbering.iter(NUM):
        if _child_val(num, _ABSTRACT_ID) == abstract_id and num.find(qn("w:lvlOverride")) is None:
            return _int(num.get(_NUM_ID)) or 0
    return new_num(numbering, abstract_id)


def abstract_id_of(numbering: etree._Element, num_id: int) -> int | None:
    abstract = abstract_of(numbering, num_id)
    return _int(abstract.get(_ABSTRACT_ID)) if abstract is not None else None


def set_num_pr(p: etree._Element, num_id: int, level: int) -> None:
    """Write the paragraph's own ``w:numPr``."""
    pr = ensure_child(get_or_add_ppr(p), "w:numPr", PPR_ORDER)
    for child in list(pr):
        if child.tag in (_ILVL, _NUM_ID):
            pr.remove(child)
    pr.insert(0, el("w:ilvl", {"w:val": str(level)}))
    pr.insert(1, el("w:numId", {"w:val": str(num_id)}))
