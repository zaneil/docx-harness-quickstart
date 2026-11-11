"""Formatting: run properties, paragraph properties and style lookup."""

from __future__ import annotations

import re

from lxml import etree

from .ooxml import (
    PPR,
    PPR_ORDER,
    RPR,
    RPR_ORDER,
    ensure_child,
    get_or_add_ppr,
    get_or_add_rpr,
    qn,
    remove_child,
)

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


def run_format(r: etree._Element) -> dict:
    """Direct formatting set on a run (style-inherited formatting is not resolved)."""
    rpr = r.find(RPR)
    out: dict = {}
    for key, tag in (("bold", "w:b"), ("italic", "w:i"), ("underline", "w:u"), ("strike", "w:strike")):
        if _is_on(rpr, tag):
            out[key] = True
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


def _key(value: str) -> str:
    return re.sub(r"[\s_-]+", "", value).lower()


def resolve_style(styles_root: etree._Element | None, name_or_id: str, style_type: str = "paragraph") -> str:
    """Map a style name or id ("Heading 1", "heading1", "Heading1") to its id."""
    candidates = [s for s in list_styles(styles_root) if s["type"] == style_type]
    wanted = _key(name_or_id)
    for style in candidates:
        if _key(style["id"]) == wanted or _key(style["name"]) == wanted:
            return style["id"]
    available = ", ".join(sorted(s["name"] for s in candidates)) or "(none)"
    raise StyleError(
        f"No {style_type} style named {name_or_id!r} in this document. Available: {available}"
    )


def style_name(styles_root: etree._Element | None, style_id: str | None, style_type: str = "paragraph") -> str:
    """Display name for a style id; the default style's name when id is None."""
    for style in list_styles(styles_root):
        if style["type"] != style_type:
            continue
        if (style_id is None and style["default"]) or style["id"] == style_id:
            return style["name"]
    return style_id or "Normal"
