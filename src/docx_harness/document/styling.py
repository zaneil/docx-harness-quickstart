"""Formatting: run properties, paragraph properties and style lookup."""

from __future__ import annotations

import re

from lxml import etree

from . import builtin_styles
from .ooxml import (
    PPR,
    PPR_ORDER,
    RPR,
    RPR_ORDER,
    W_NS,
    el,
    ensure_child,
    get_or_add_ppr,
    get_or_add_rpr,
    qn,
    remove_child,
)
from .package import parse_xml

VAL = qn("w:val")

NAMED_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
    "blue": "0000FF", "yellow": "FFFF00", "orange": "FFA500", "purple": "800080",
    "gray": "808080", "grey": "808080", "navy": "000080", "teal": "008080",
}
HIGHLIGHTS = {
    "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue", "darkCyan",
    "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black",
}
ALIGNMENTS = {"left": "left", "center": "center", "centre": "center", "right": "right",
              "justify": "both", "justified": "both", "both": "both"}


class StyleError(ValueError):
    pass


# ---------------------------------------------------------------- run format

def _toggle(props: etree._Element, tag: str, order: list[str], value: bool) -> None:
    element = ensure_child(props, tag, order)
    if value:
        element.attrib.pop(VAL, None)
    else:
        element.set(VAL, "0")


def normalize_color(color: str) -> str:
    value = NAMED_COLORS.get(color.strip().lower(), color.strip().lstrip("#"))
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        raise StyleError(f"Invalid color {color!r}. Use a 6-digit hex value such as '1F4E79'.")
    return value.upper()


def apply_run_format(
    r: etree._Element,
    *,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    strike: bool | None = None,
    color: str | None = None,
    size_pt: float | None = None,
    font: str | None = None,
    highlight: str | None = None,
) -> None:
    """Apply direct formatting to a run. ``None`` leaves a property untouched."""
    rpr = get_or_add_rpr(r)
    if bold is not None:
        _toggle(rpr, "w:b", RPR_ORDER, bold)
    if italic is not None:
        _toggle(rpr, "w:i", RPR_ORDER, italic)
    if strike is not None:
        _toggle(rpr, "w:strike", RPR_ORDER, strike)
    if underline is not None:
        ensure_child(rpr, "w:u", RPR_ORDER).set(VAL, "single" if underline else "none")
    if color is not None:
        element = ensure_child(rpr, "w:color", RPR_ORDER)
        element.attrib.clear()
        element.set(VAL, normalize_color(color))
    if size_pt is not None:
        if not 1 <= size_pt <= 400:
            raise StyleError("size_pt must be between 1 and 400")
        half_points = str(round(size_pt * 2))
        ensure_child(rpr, "w:sz", RPR_ORDER).set(VAL, half_points)
        ensure_child(rpr, "w:szCs", RPR_ORDER).set(VAL, half_points)
    if font is not None:
        element = ensure_child(rpr, "w:rFonts", RPR_ORDER)
        element.attrib.clear()
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            element.set(qn(attr), font)
    if highlight is not None:
        if highlight.lower() in ("none", ""):
            remove_child(rpr, "w:highlight")
        else:
            match = next((h for h in HIGHLIGHTS if h.lower() == highlight.lower()), None)
            if match is None:
                raise StyleError(f"Invalid highlight {highlight!r}. Options: {', '.join(sorted(HIGHLIGHTS))}, none.")
            ensure_child(rpr, "w:highlight", RPR_ORDER).set(VAL, match)


def _is_on(props: etree._Element | None, tag: str) -> bool:
    if props is None:
        return False
    element = props.find(qn(tag))
    return element is not None and element.get(VAL, "1") not in ("0", "false", "none")


def run_format(r: etree._Element, explicit_off: bool = False) -> dict:
    """Direct formatting set on a run (style-inherited formatting is not
    resolved). ``r`` is anything with a ``w:rPr`` child: a run, or a style.
    With ``explicit_off`` a property switched off outright (``<w:b w:val="0"/>``,
    which overrides the style) is reported as ``False`` instead of left out."""
    rpr = r.find(RPR)
    out: dict = {}
    for key, tag in (("bold", "w:b"), ("italic", "w:i"), ("underline", "w:u"), ("strike", "w:strike")):
        if _is_on(rpr, tag):
            out[key] = True
        elif explicit_off and rpr is not None and rpr.find(qn(tag)) is not None:
            out[key] = False
    if rpr is not None:
        color = rpr.find(qn("w:color"))
        if color is not None and color.get(VAL) not in (None, "auto"):
            out["color"] = color.get(VAL)
        size = rpr.find(qn("w:sz"))
        if size is not None and size.get(VAL, "").isdigit():
            out["size_pt"] = int(size.get(VAL)) / 2
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is not None and fonts.get(qn("w:ascii")):
            out["font"] = fonts.get(qn("w:ascii"))
        highlight = rpr.find(qn("w:highlight"))
        if highlight is not None:
            out["highlight"] = highlight.get(VAL)
    return out


# ---------------------------------------------------------- paragraph format

def _number(element: etree._Element, attribute: str) -> float | None:
    try:
        return float(element.get(qn(attribute)))
    except (TypeError, ValueError):
        return None


def paragraph_format(p: etree._Element) -> dict:
    """Direct paragraph formatting, with the same keys as
    :func:`apply_paragraph_format` (style-inherited formatting is not
    resolved). ``p`` is anything with a ``w:pPr`` child: a paragraph, or a style."""
    ppr = p.find(PPR)
    out: dict = {}
    if ppr is None:
        return out
    jc = ppr.find(qn("w:jc"))
    if jc is not None:
        alignment = _ALIGNMENT_NAMES.get(jc.get(VAL, ""))
        if alignment:
            out["alignment"] = alignment
    spacing = ppr.find(qn("w:spacing"))
    if spacing is not None:
        for key, attribute in (("space_before_pt", "w:before"), ("space_after_pt", "w:after")):
            twips = _number(spacing, attribute)
            if twips is not None:
                out[key] = twips / 20
        line = _number(spacing, "w:line")
        if line is not None and spacing.get(qn("w:lineRule"), "auto") == "auto":
            out["line_spacing"] = round(line / 240, 2)
    ind = ppr.find(qn("w:ind"))
    if ind is not None:
        for key, attributes in (("indent_left_in", ("w:left", "w:start")), ("indent_right_in", ("w:right", "w:end"))):
            twips = next((v for v in (_number(ind, a) for a in attributes) if v is not None), None)
            if twips is not None:
                out[key] = round(twips / 1440, 3)
        hanging, first = _number(ind, "w:hanging"), _number(ind, "w:firstLine")
        if hanging is not None or first is not None:
            out["first_line_indent_in"] = round((-hanging if hanging is not None else first) / 1440, 3)
    if _is_on(ppr, "w:pageBreakBefore"):
        out["page_break_before"] = True
    if _is_on(ppr, "w:keepNext"):
        out["keep_with_next"] = True
    return out


_ALIGNMENT_NAMES = {"left": "left", "start": "left", "center": "center", "right": "right", "end": "right",
                    "both": "justify", "distribute": "justify"}


def apply_paragraph_format(
    p: etree._Element,
    *,
    alignment: str | None = None,
    space_before_pt: float | None = None,
    space_after_pt: float | None = None,
    line_spacing: float | None = None,
    indent_left_in: float | None = None,
    indent_right_in: float | None = None,
    first_line_indent_in: float | None = None,
    page_break_before: bool | None = None,
    keep_with_next: bool | None = None,
) -> None:
    """Apply direct paragraph formatting. ``None`` leaves a property untouched."""
    ppr = get_or_add_ppr(p)
    if alignment is not None:
        if alignment.lower() not in ALIGNMENTS:
            raise StyleError("alignment must be one of: left, center, right, justify")
        ensure_child(ppr, "w:jc", PPR_ORDER).set(VAL, ALIGNMENTS[alignment.lower()])
    if space_before_pt is not None or space_after_pt is not None or line_spacing is not None:
        spacing = ensure_child(ppr, "w:spacing", PPR_ORDER)
        if space_before_pt is not None:
            spacing.set(qn("w:before"), str(round(space_before_pt * 20)))
        if space_after_pt is not None:
            spacing.set(qn("w:after"), str(round(space_after_pt * 20)))
        if line_spacing is not None:
            if not 0.5 <= line_spacing <= 10:
                raise StyleError("line_spacing is a multiple such as 1.0, 1.5 or 2.0")
            spacing.set(qn("w:line"), str(round(line_spacing * 240)))
            spacing.set(qn("w:lineRule"), "auto")
    if indent_left_in is not None or indent_right_in is not None or first_line_indent_in is not None:
        ind = ensure_child(ppr, "w:ind", PPR_ORDER)
        if indent_left_in is not None:
            ind.set(qn("w:left"), str(round(indent_left_in * 1440)))
        if indent_right_in is not None:
            ind.set(qn("w:right"), str(round(indent_right_in * 1440)))
        if first_line_indent_in is not None:
            ind.attrib.pop(qn("w:firstLine"), None)
            ind.attrib.pop(qn("w:hanging"), None)
            twips = round(first_line_indent_in * 1440)
            ind.set(qn("w:firstLine") if twips >= 0 else qn("w:hanging"), str(abs(twips)))
    if page_break_before is not None:
        _toggle(ppr, "w:pageBreakBefore", PPR_ORDER, page_break_before)
    if keep_with_next is not None:
        _toggle(ppr, "w:keepNext", PPR_ORDER, keep_with_next)


def set_paragraph_style(p: etree._Element, style_id: str) -> None:
    ppr = get_or_add_ppr(p)
    ensure_child(ppr, "w:pStyle", PPR_ORDER).set(VAL, style_id)


def paragraph_style_id(p: etree._Element) -> str | None:
    ppr = p.find(PPR)
    if ppr is None:
        return None
    style = ppr.find(qn("w:pStyle"))
    return style.get(VAL) if style is not None else None


# -------------------------------------------------------------------- styles

def list_styles(styles_root: etree._Element | None) -> list[dict]:
    """All styles defined in styles.xml as ``{id, name, type}``."""
    if styles_root is None:
        return []
    out = []
    for style in styles_root.iter(qn("w:style")):
        name = style.find(qn("w:name"))
        out.append({
            "id": style.get(qn("w:styleId"), ""),
            "name": name.get(VAL, "") if name is not None else style.get(qn("w:styleId"), ""),
            "type": style.get(qn("w:type"), "paragraph"),
            "default": style.get(qn("w:default")) in ("1", "true"),
        })
    return out


def style_format(styles_root: etree._Element | None, style_id: str | None) -> dict:
    """What a paragraph style contributes before any direct formatting:
    ``{"paragraph": {...}, "run": {...}}`` with the keys of
    :func:`paragraph_format` and :func:`run_format`, resolved through the
    document defaults and the style's ``basedOn`` chain."""
    paragraph: dict = {}
    run: dict = {}
    if styles_root is None:
        return {"paragraph": paragraph, "run": run}
    layers: list[etree._Element] = []
    defaults = styles_root.find(qn("w:docDefaults"))
    if defaults is not None:
        layers += [d for d in (defaults.find(qn("w:pPrDefault")), defaults.find(qn("w:rPrDefault"))) if d is not None]
    styles = {s.get(qn("w:styleId")): s for s in styles_root.iter(qn("w:style")) if s.get(qn("w:type")) == "paragraph"}
    style = styles.get(style_id) if style_id else next(
        (s for s in styles.values() if s.get(qn("w:default")) in ("1", "true")), None)
    chain: list[etree._Element] = []
    while style is not None and style not in chain:  # most specific first; stop on a cycle
        chain.append(style)
        based_on = style.find(qn("w:basedOn"))
        style = styles.get(based_on.get(VAL)) if based_on is not None else None
    for layer in layers + chain[::-1]:
        paragraph.update(paragraph_format(layer))
        run.update(run_format(layer, explicit_off=True))
    return {"paragraph": paragraph, "run": {k: v for k, v in run.items() if v is not False}}


def _key(value: str) -> str:
    return re.sub(r"[\s_-]+", "", value).lower()


def _find_style(styles_root: etree._Element | None, name_or_id: str, style_type: str) -> str | None:
    wanted = _key(name_or_id)
    for style in list_styles(styles_root):
        if style["type"] == style_type and wanted in (_key(style["id"]), _key(style["name"])):
            return style["id"]
    return None


def _builtin(name_or_id: str, style_type: str) -> builtin_styles.BuiltinStyle | None:
    wanted = _key(name_or_id)
    return next((b for b in builtin_styles.ALL
                 if b.type == style_type and wanted in (_key(b.style_id), _key(b.name))), None)


def resolve_style(styles_root: etree._Element | None, name_or_id: str, style_type: str = "paragraph") -> str:
    """Map a style name or id ("Heading 1", "heading1", "Heading1") to its id."""
    found = _find_style(styles_root, name_or_id, style_type)
    if found is not None:
        return found
    candidates = [s for s in list_styles(styles_root) if s["type"] == style_type]
    available = ", ".join(sorted(s["name"] for s in candidates)) or "(none)"
    raise StyleError(
        f"No {style_type} style named {name_or_id!r} in this document. Available: {available}. "
        "Use one of these, or define a new style with create_style."
    )


def default_style_id(styles_root: etree._Element | None, style_type: str) -> str | None:
    return next((s["id"] for s in list_styles(styles_root) if s["type"] == style_type and s["default"]), None)


def unique_style_id(styles_root: etree._Element, wanted: str) -> str:
    taken = {s["id"] for s in list_styles(styles_root)}
    candidate, number = wanted, 1
    while candidate in taken:
        number += 1
        candidate = f"{wanted}{number}"
    return candidate


def add_style(styles_root: etree._Element, xml: str) -> etree._Element:
    """Append a ``w:style`` given as text (no namespace declarations). A
    ``basedOn`` naming a style the document lacks is pointed at the document's
    default style of that type instead, or dropped when it has none."""
    style = parse_xml(f'<w:styles xmlns:w="{W_NS}">{xml}</w:styles>')[0]
    based_on = style.find(qn("w:basedOn"))
    if based_on is not None and based_on.get(VAL) not in {s["id"] for s in list_styles(styles_root)}:
        fallback = default_style_id(styles_root, style.get(qn("w:type"), "paragraph"))
        if fallback:
            based_on.set(VAL, fallback)
        else:
            style.remove(based_on)
    next_style = style.find(qn("w:next"))
    if next_style is not None and next_style.get(VAL) not in {s["id"] for s in list_styles(styles_root)}:
        style.remove(next_style)
    styles_root.append(style)
    return style


def ensure_style(styles_root: etree._Element, name_or_id: str, style_type: str = "paragraph",
                 list_num_id=None) -> str:
    """Like :func:`resolve_style`, but a well-known style the document does
    not define (Word leaves unused built-in styles out of the file) is added
    from :mod:`builtin_styles`. ``list_num_id(kind)`` supplies the numbering
    id for a list style; without it list styles are not created."""
    found = _find_style(styles_root, name_or_id, style_type)
    if found is not None:
        return found
    builtin = _builtin(name_or_id, style_type)
    if builtin is None or (builtin.list_kind and list_num_id is None):
        return resolve_style(styles_root, name_or_id, style_type)  # raises, listing what there is
    num_id = list_num_id(builtin.list_kind) if builtin.list_kind else None
    style_id = unique_style_id(styles_root, builtin.style_id)
    add_style(styles_root, builtin.xml(num_id, style_id))
    return style_id


def style_num_pr(styles_root: etree._Element | None, style_id: str | None) -> tuple[int | None, int | None]:
    """``(numbering id, level)`` a paragraph style supplies, through its
    ``basedOn`` chain; ``style_id`` None means the default paragraph style."""
    if styles_root is None:
        return None, None
    styles = {s.get(qn("w:styleId")): s for s in styles_root.iter(qn("w:style")) if s.get(qn("w:type")) == "paragraph"}
    style = styles.get(style_id) if style_id else next(
        (s for s in styles.values() if s.get(qn("w:default")) in ("1", "true")), None)
    num_id = level = None
    seen: list[etree._Element] = []
    while style is not None and style not in seen:
        seen.append(style)
        ppr = style.find(PPR)
        pr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if pr is not None:
            for tag, current in (("w:numId", num_id), ("w:ilvl", level)):
                child = pr.find(qn(tag))
                if current is None and child is not None and child.get(VAL, "").lstrip("-").isdigit():
                    if tag == "w:numId":
                        num_id = int(child.get(VAL))
                    else:
                        level = int(child.get(VAL))
        based_on = style.find(qn("w:basedOn"))
        style = styles.get(based_on.get(VAL)) if based_on is not None else None
    return num_id, level


RUN_KEYS = ("bold", "italic", "underline", "strike", "color", "size_pt", "font", "highlight")
PARAGRAPH_KEYS = ("alignment", "space_before_pt", "space_after_pt", "line_spacing", "indent_left_in",
                  "indent_right_in", "first_line_indent_in", "page_break_before", "keep_with_next")


def create_style(styles_root: etree._Element, name: str, style_type: str = "paragraph",
                 based_on_id: str | None = None, **fmt) -> str:
    """Define a new paragraph or character style and return its id. ``fmt``
    takes the keywords of :func:`apply_run_format` and, for a paragraph
    style, :func:`apply_paragraph_format`."""
    name = " ".join(str(name).split())
    if not name:
        raise StyleError("A style needs a name.")
    if style_type not in ("paragraph", "character"):
        raise StyleError("style_type must be 'paragraph' or 'character'.")
    if _find_style(styles_root, name, style_type) is not None:
        raise StyleError(f"A {style_type} style named {name!r} already exists. Use it as it is, or pick another name.")
    fmt = {k: v for k, v in fmt.items() if v is not None}
    allowed = RUN_KEYS + (PARAGRAPH_KEYS if style_type == "paragraph" else ())
    unknown = sorted(set(fmt) - set(allowed))
    if unknown:
        raise StyleError(f"A {style_type} style cannot set {', '.join(unknown)}. Options: {', '.join(allowed)}.")
    style_id = unique_style_id(styles_root, re.sub(r"[^A-Za-z0-9]", "", name) or "Style")
    style = el("w:style", {"w:type": style_type, "w:customStyle": "1", "w:styleId": style_id},
               el("w:name", {"w:val": name}))
    if based_on_id:
        style.append(el("w:basedOn", {"w:val": based_on_id}))
    style.append(el("w:qFormat"))
    # Created here, in schema order, so the format helpers find them in place.
    ppr, rpr = el("w:pPr"), el("w:rPr")
    style.append(ppr)
    style.append(rpr)
    apply_paragraph_format(style, **{k: v for k, v in fmt.items() if k in PARAGRAPH_KEYS})
    apply_run_format(style, **{k: v for k, v in fmt.items() if k in RUN_KEYS})
    for props in (ppr, rpr):
        if len(props) == 0:
            style.remove(props)
    styles_root.append(style)
    return style_id


def style_name(styles_root: etree._Element | None, style_id: str | None, style_type: str = "paragraph") -> str:
    """Display name for a style id; the default style's name when id is None."""
    for style in list_styles(styles_root):
        if style["type"] != style_type:
            continue
        if (style_id is None and style["default"]) or style["id"] == style_id:
            return style["name"]
    return style_id or "Normal"
