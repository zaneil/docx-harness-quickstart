"""Inline pictures and hyperlinks.

The pure image helpers it calls are in ``images.py``.

Mixed into ``WordDocument`` (``document.py``), which provides the block-id
helpers it relies on (``_get``, ``_body_block``, ``_table``, ``_paragraphs_for``,
``_describe``...) and ``self.edits``, the single path for every mutation."""

from __future__ import annotations

import os
import re
from pathlib import Path

from lxml import etree

from . import images, styling, tables
from .common import (
    _VAL,
    PAGE_WIDTH_TWIPS,
    DocumentError,
    _live_paragraphs,
)
from .ooxml import (
    DEL,
    DRAWING,
    HYPERLINK,
    MOVE_FROM,
    R_ID,
    RPR,
    RPR_ORDER,
    SECTPR,
    R,
    el,
    ensure_child,
    get_or_add_rpr,
    isolate_range,
    iter_runs,
    paragraph_text,
    qn,
    run_text,
)

_URL_SCHEMES = ("http", "https", "mailto")


class MediaMixin:
    # ---------------------------------------------------------------- images

    def _images(self, p: etree._Element) -> list[dict]:
        """The pictures in a paragraph as ``{"name", "alt", "width_in", "height_in"}``."""
        out = []
        for run in iter_runs(p):
            for node in run.iter(DRAWING):
                info = images.describe(node)
                rel_id = info.pop("rel_id")
                relationship = self.package.relationship(self._part_of(p), rel_id) if rel_id else None
                out.append({"name": relationship["target"].rpartition("/")[2] if relationship else "", **info})
        return out

    def _text_width_in(self) -> float:
        """Width between the page margins, from the section properties."""
        sect = self._body.find(SECTPR)
        size = sect.find(qn("w:pgSz")) if sect is not None else None
        margins = sect.find(qn("w:pgMar")) if sect is not None else None
        try:
            twips = int(size.get(qn("w:w"))) - int(margins.get(qn("w:left"))) - int(margins.get(qn("w:right")))
        except (AttributeError, TypeError, ValueError):
            twips = PAGE_WIDTH_TWIPS
        return (twips if twips > 0 else PAGE_WIDTH_TWIPS) / tables.TWIPS_PER_INCH

    @staticmethod
    def _image_size(natural: tuple[float, float], width_in: float | None, height_in: float | None,
                    fit_width: float | None = None) -> tuple[float, float]:
        try:
            width = float(width_in) if width_in is not None else None
            height = float(height_in) if height_in is not None else None
        except (TypeError, ValueError):
            raise DocumentError("width_in and height_in must be numbers, in inches.") from None
        ratio = natural[1] / natural[0]
        if width is None and height is None:
            width, height = natural
            if fit_width and width > fit_width:
                width, height = fit_width, fit_width * ratio
        elif height is None:
            height = width * ratio
        elif width is None:
            width = height / ratio
        if not (0 < width <= 22 and 0 < height <= 22):
            raise DocumentError("An image must be more than 0 and at most 22 inches wide and high.")
        return width, height

    def insert_image(
        self,
        source: str | os.PathLike | bytes,
        after: str = "end",
        width_in: float | None = None,
        height_in: float | None = None,
        alt_text: str = "",
        alignment: str | None = None,
    ) -> str:
        """Insert a picture (PNG, JPEG or GIF) in a paragraph of its own after
        block ``after``. ``source`` is a file path or the image bytes. With
        one of ``width_in`` / ``height_in`` the other keeps the proportions;
        with neither the picture has its natural size, reduced to fit between
        the margins. ``alt_text`` describes it for screen readers. Returns
        the new block id."""
        if isinstance(source, (bytes, bytearray)):
            data, filename = bytes(source), "image"
        else:
            try:
                data, filename = Path(source).read_bytes(), Path(source).name
            except OSError as exc:
                raise DocumentError(f"Cannot read the image file {str(source)!r}: {exc.strerror or exc}.") from None
        try:
            info = images.probe(data)
        except images.ImageError as exc:
            raise DocumentError(str(exc)) from None
        width, height = self._image_size(info.size_in, width_in, height_in, self._text_width_in())
        index = self._index_after(after)

        main = self.package.main_document_part
        media = self._part_beside_main("media/image")
        part = next(
            (name for name in self.package.list_parts()
             if name.startswith(media) and name.endswith("." + info.format) and self.package.part_bytes(name) == data),
            None,
        )
        if part is None:  # the same picture used twice is stored once
            part = self.package.add_part(self.package.next_part_name(media, "." + info.format), data,
                                         images.CONTENT_TYPES[info.format])
        rel_id = self.package.relate(main, "image", part)
        number = images.next_number(self.package.xml(main), *(root for _, _, root in self._story_roots()))
        p = el("w:p", None, el("w:r", None, images.drawing(rel_id, width, height, number, filename, str(alt_text or ""))))
        if alignment is not None:
            styling.apply_paragraph_format(p, alignment=alignment)
        self.edits.insert_block(self._body, index, p)
        return self._id_of(p)

    def set_image(self, block_id: str, width_in: float | None = None, height_in: float | None = None,
                  alt_text: str | None = None) -> None:
        """Resize the picture in a block (giving one dimension keeps the
        proportions) or change its alt text. Not recorded as a tracked change."""
        if width_in is None and height_in is None and alt_text is None:
            raise DocumentError("Give at least one of width_in, height_in, alt_text.")
        element = self._get(block_id)
        node = next((d for p in _live_paragraphs(element) for run in iter_runs(p) for d in run.iter(DRAWING)), None)
        if node is None:
            raise DocumentError(f"Block {block_id} has no picture.")
        current = images.describe(node)
        if width_in is not None or height_in is not None:
            if "width_in" not in current:
                raise DocumentError(f"The picture in {block_id} has no readable size, so it cannot be resized.")
            images.resize(node, *self._image_size((current["width_in"], current["height_in"]), width_in, height_in))
        if alt_text is not None:
            images.set_alt_text(node, str(alt_text))

    # ------------------------------------------------------------ hyperlinks

    def _part_of(self, element: etree._Element) -> str:
        """Name of the part ``element`` is in: the main document, or a header,
        footer or notes part. Relationships are kept per part."""
        root = element.getroottree().getroot()
        for part, _, story in self._story_roots():
            if story is root:
                return part
        return self.package.main_document_part

    @staticmethod
    def _link_of(run: etree._Element, p: etree._Element) -> etree._Element | None:
        """The hyperlink of paragraph ``p`` that ``run`` is inside."""
        for ancestor in run.iterancestors():
            if ancestor is p:
                return None
            if ancestor.tag == HYPERLINK:
                return ancestor
        return None

    def _links(self, p: etree._Element) -> list[dict]:
        """The hyperlinks in a paragraph as ``{"text", "url"}``; a link to a
        bookmark in the document reads as ``#name``."""
        found: dict[etree._Element, str] = {}
        for run in iter_runs(p):
            link = self._link_of(run, p)
            if link is not None:
                found[link] = found.get(link, "") + run_text(run)
        out = []
        for link, text in found.items():
            relationship = self.package.relationship(self._part_of(p), link.get(R_ID, "")) if link.get(R_ID) else None
            url = relationship["target"] if relationship else "#" + link.get(qn("w:anchor"), "")
            if text:
                out.append({"text": text, "url": url})
        return out

    @staticmethod
    def _check_url(url: str) -> str:
        url = str(url).strip()
        scheme, _, rest = url.partition(":")
        ok = scheme.lower() in _URL_SCHEMES and rest and not any(ch.isspace() or ord(ch) < 0x20 for ch in url)
        if ok and scheme.lower() != "mailto":
            ok = rest.startswith("//") and len(rest) > 2
        if not ok:
            raise DocumentError(
                f"Invalid link address {url!r}. Give a full address starting with https://, http:// or mailto:."
            )
        return url

    def add_hyperlink(self, ref: str, url: str, match: str | None = None, case_sensitive: bool = True) -> int:
        """Turn text into a link to ``url`` (https, http or mailto): the whole
        text of ``ref`` (a block, table cell or header/footer part), or only
        the occurrences of ``match`` in it. Text that is already a link is
        pointed at ``url``. Returns the number of links made.

        When changes are tracked, the text is recorded as replaced by its
        linked copy; re-pointing an existing link is not recorded."""
        url = self._check_url(url)
        style_id = self._style_id("Hyperlink", "character")
        count = 0
        for p in self._paragraphs_for(ref):
            text = paragraph_text(p)
            if match:
                pattern = re.compile(re.escape(match), 0 if case_sensitive else re.IGNORECASE)
                spans = [(m.start(), m.end()) for m in pattern.finditer(text)]
            else:
                spans = [(0, len(text))] if text else []
            if not spans:
                continue
            rel_id = self.package.relate(self._part_of(p), "hyperlink", url, external=True)
            for start, end in reversed(spans):
                runs = isolate_range(p, start, end)
                existing = {run: self._link_of(run, p) for run in runs}
                for link in {link for link in existing.values() if link is not None}:
                    link.set(R_ID, rel_id)
                    link.attrib.pop(qn("w:anchor"), None)
                free = [run for run in runs if existing[run] is None]
                for wrapper in self.edits.wrap_runs(free, el("w:hyperlink", {"r:id": rel_id, "w:history": "1"})):
                    for run in wrapper.iter(R):
                        if run.getparent().tag not in (DEL, MOVE_FROM):
                            self.edits.before_run_format(run)
                            ensure_child(get_or_add_rpr(run), "w:rStyle", RPR_ORDER).set(_VAL, style_id)
                count += 1
        if count == 0:
            raise DocumentError(f"Text {match!r} was not found in {ref}." if match else f"{ref} has no text to link.")
        return count

    def remove_hyperlink(self, ref: str, match: str | None = None) -> int:
        """Turn links back into plain text: every link in ``ref``, or only
        those whose text contains ``match``. Returns the number removed."""
        style_id = styling._find_style(self._styles, "Hyperlink", "character")
        count = 0
        for p in self._paragraphs_for(ref):
            texts: dict[etree._Element, str] = {}
            for run in iter_runs(p):
                link = self._link_of(run, p)
                if link is not None:
                    texts[link] = texts.get(link, "") + run_text(run)
            for link, text in texts.items():
                if match and match not in text:
                    continue
                for run in self.edits.unwrap_runs(link):
                    rpr = run.find(RPR)
                    style = rpr.find(qn("w:rStyle")) if rpr is not None else None
                    if style is not None and style.get(_VAL) == style_id:
                        self.edits.before_run_format(run)
                        rpr.remove(style)
                count += 1
        if count == 0:
            raise DocumentError(f"No link{f' containing {match!r}' if match else ''} was found in {ref}.")
        return count
