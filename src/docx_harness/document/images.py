"""Pictures: reading an image file's type and size, and the inline drawing
that places one in a paragraph.

Only PNG, JPEG and GIF are accepted, recognised by their first bytes rather
than by file name, so nothing else can be embedded by mistake.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass
from xml.sax.saxutils import quoteattr

from lxml import etree

from .ooxml import A_NS, DRAWING, PIC_NS, R_NS, W_NS, WP_NS, qn
from .package import parse_xml

EMU_PER_INCH = 914400
MAX_BYTES = 20 * 1024 * 1024
DEFAULT_DPI = 96.0
CONTENT_TYPES = {"png": "image/png", "jpeg": "image/jpeg", "gif": "image/gif"}
EXTENT = qn("wp:extent")
DOC_PR = qn("wp:docPr")
BLIP = qn("a:blip")
_EXT = qn("a:ext")
_EMBED = qn("r:embed")


class ImageError(ValueError):
    """The data is not an image this library can embed."""


@dataclass(frozen=True)
class ImageInfo:
    format: str  # png, jpeg or gif
    width_px: int
    height_px: int
    dpi: float = DEFAULT_DPI

    @property
    def size_in(self) -> tuple[float, float]:
        return self.width_px / self.dpi, self.height_px / self.dpi


def probe(data: bytes) -> ImageInfo:
    """Type and pixel size of an image, from its header."""
    if len(data) > MAX_BYTES:
        raise ImageError(f"The image is larger than {MAX_BYTES // (1024 * 1024)} MB.")
    try:
        if data[:8] == b"\x89PNG\r\n\x1a\n":
            info = _png(data)
        elif data[:3] == b"\xff\xd8\xff":
            info = _jpeg(data)
        elif data[:6] in (b"GIF87a", b"GIF89a"):
            width, height = struct.unpack("<HH", data[6:10])
            info = ImageInfo("gif", width, height)
        else:
            info = None
    except (struct.error, IndexError):
        raise ImageError("The image file is damaged or cut short.") from None
    if info is None:
        raise ImageError("Not a PNG, JPEG or GIF image. Convert the picture to one of these formats first.")
    if info.width_px <= 0 or info.height_px <= 0:
        raise ImageError("The image has no size.")
    return info


def _png(data: bytes) -> ImageInfo:
    width, height = struct.unpack(">II", data[16:24])
    dpi = DEFAULT_DPI
    position = 8
    while position + 8 <= len(data):
        length, kind = struct.unpack(">I4s", data[position:position + 8])
        if kind == b"pHYs":
            per_unit, _, unit = struct.unpack(">IIB", data[position + 8:position + 17])
            if unit == 1 and per_unit:  # pixels per metre
                dpi = per_unit * 0.0254
            break
        if kind in (b"IDAT", b"IEND"):
            break
        position += 12 + length
    return ImageInfo("png", width, height, dpi if 30 <= dpi <= 1200 else DEFAULT_DPI)


def _jpeg(data: bytes) -> ImageInfo | None:
    dpi = DEFAULT_DPI
    position = 2
    while position + 4 <= len(data):
        if data[position] != 0xFF:
            position += 1
            continue
        marker = data[position + 1]
        if marker in (0xFF, 0x00) or 0xD0 <= marker <= 0xD9:
            position += 1 if marker == 0xFF else 2
            continue
        length = struct.unpack(">H", data[position + 2:position + 4])[0]
        segment = data[position + 4:position + 2 + length]
        if marker == 0xE0 and segment[:5] == b"JFIF\x00" and len(segment) >= 12:
            units, x_density = segment[7], struct.unpack(">H", segment[8:10])[0]
            if x_density and units in (1, 2):
                dpi = x_density * (1.0 if units == 1 else 2.54)
        if 0xC0 <= marker <= 0xCF and marker not in (0xC4, 0xC8, 0xCC):
            height, width = struct.unpack(">HH", segment[1:5])
            return ImageInfo("jpeg", width, height, dpi if 30 <= dpi <= 1200 else DEFAULT_DPI)
        position += 2 + length
    return None


def drawing(rel_id: str, width_in: float, height_in: float, number: int, name: str, alt_text: str = "") -> etree._Element:
    """A ``w:drawing`` placing the related image in the line of text."""
    cx, cy = round(width_in * EMU_PER_INCH), round(height_in * EMU_PER_INCH)
    return parse_xml(
        f'<w:drawing xmlns:w="{W_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}" xmlns:r="{R_NS}">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{number}" name="Picture {number}" descr={quoteattr(alt_text)}/>'
        f'<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'<a:graphic><a:graphicData uri="{PIC_NS}"><pic:pic>'
        f'<pic:nvPicPr><pic:cNvPr id="0" name={quoteattr(name)}/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>'
    )


def next_number(*roots: etree._Element) -> int:
    """A drawing id not used in any of ``roots``."""
    used = [int(pr.get("id")) for root in roots for pr in root.iter(DOC_PR) if (pr.get("id") or "").isdigit()]
    return max(used, default=0) + 1


def describe(element: etree._Element) -> dict:
    """``{"alt", "width_in", "height_in", "rel_id"}`` of a ``w:drawing``."""
    extent = next(element.iter(EXTENT), None)
    doc_pr = next(element.iter(DOC_PR), None)
    blip = next(element.iter(BLIP), None)
    out: dict = {"alt": (doc_pr.get("descr") or "") if doc_pr is not None else ""}
    if extent is not None and (extent.get("cx") or "").isdigit() and (extent.get("cy") or "").isdigit():
        out["width_in"] = round(int(extent.get("cx")) / EMU_PER_INCH, 2)
        out["height_in"] = round(int(extent.get("cy")) / EMU_PER_INCH, 2)
    out["rel_id"] = blip.get(_EMBED) if blip is not None else None
    return out


def resize(element: etree._Element, width_in: float, height_in: float) -> None:
    cx, cy = str(round(width_in * EMU_PER_INCH)), str(round(height_in * EMU_PER_INCH))
    for node in list(element.iter(EXTENT)) + [ext for ext in element.iter(_EXT) if ext.getparent().tag == qn("a:xfrm")]:
        node.set("cx", cx)
        node.set("cy", cy)


def set_alt_text(element: etree._Element, alt_text: str) -> None:
    for doc_pr in element.iter(DOC_PR):
        doc_pr.set("descr", alt_text)


__all__ = ["ImageError", "ImageInfo", "probe", "drawing", "describe", "resize", "set_alt_text", "next_number", "DRAWING"]
