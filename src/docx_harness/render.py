"""Lay out a document as pages, with LibreOffice.

    from docx_harness import render

    pages = render.render(doc)          # a WordDocument, a path or .docx bytes
    pages.page_count                    # 2
    pages.page_of("Appendix")           # 2: the page where the text first appears
    pages.pages[0]                      # page 1's text, headers, footers and footnotes included

LibreOffice converts the file to PDF and pypdfium2 reads the pages back. It is
a close stand-in for Word, not Word: page and line breaks can differ.

Optional: needs LibreOffice (``brew install --cask libreoffice`` or
``apt install libreoffice-writer-nogui``; set ``DOCX_HARNESS_SOFFICE`` to its
``soffice`` program if it is not found) and ``pip install docx-harness[render]``.
Nothing in the library or the agent depends on this module.
"""

from __future__ import annotations

import atexit
import os
import re
import shutil
import subprocess
import tempfile
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING, Union

if TYPE_CHECKING:
    from .document import WordDocument

Source = Union["WordDocument", str, os.PathLike, bytes]

_MAC_APP = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
_INSTALL = (
    "LibreOffice was not found. Install it (macOS: brew install --cask libreoffice; "
    "Debian/Ubuntu: apt install libreoffice-writer-nogui) or set DOCX_HARNESS_SOFFICE to its soffice program."
)
_IGNORED = re.compile(r"[\s­​]+")  # whitespace, soft hyphens, zero-width spaces

# Two soffice processes cannot share a user profile, so conversions take turns.
_lock = threading.Lock()
_profile: str | None = None


class RenderError(RuntimeError):
    """The document could not be laid out; the message says why and what to do."""


def find_soffice() -> str | None:
    """The LibreOffice ``soffice`` program: ``$DOCX_HARNESS_SOFFICE``, else
    ``soffice`` or ``libreoffice`` on the PATH, else the macOS app."""
    configured = os.environ.get("DOCX_HARNESS_SOFFICE")
    if configured:
        return configured if Path(configured).is_file() else None
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return _MAC_APP if Path(_MAC_APP).is_file() else None


def _pdfium():
    try:
        import pypdfium2
    except ImportError:
        raise RenderError("Reading the pages needs pypdfium2: pip install 'docx-harness[render]'.") from None
    return pypdfium2


def available() -> bool:
    """True when LibreOffice and pypdfium2 are both installed."""
    if find_soffice() is None:
        return False
    try:
        _pdfium()
    except RenderError:
        return False
    return True


def _profile_dir() -> str:
    """A LibreOffice profile of our own, so the user's settings are never touched."""
    global _profile
    if _profile is None:
        _profile = tempfile.mkdtemp(prefix="docx-harness-lo-")
        atexit.register(shutil.rmtree, _profile, True)
    return _profile


def _docx_bytes(source: Source) -> bytes:
    if isinstance(source, (bytes, bytearray)):
        return bytes(source)
    if isinstance(source, (str, os.PathLike)):
        return Path(source).read_bytes()
    package = getattr(source, "package", None)
    if package is None:
        raise TypeError(f"Cannot render {type(source).__name__}: pass a WordDocument, a path or .docx bytes.")
    return package.to_bytes()


def to_pdf(source: Source, *, timeout: float = 120) -> bytes:
    """The document as PDF, converted by LibreOffice."""
    soffice = find_soffice()
    if soffice is None:
        if os.environ.get("DOCX_HARNESS_SOFFICE"):
            raise RenderError(f"DOCX_HARNESS_SOFFICE is set to {os.environ['DOCX_HARNESS_SOFFICE']!r}, which is not a file. "
                              + _INSTALL)
        raise RenderError(_INSTALL)
    data = _docx_bytes(source)
    with tempfile.TemporaryDirectory(prefix="docx-harness-render-") as folder:
        docx = Path(folder) / "document.docx"
        docx.write_bytes(data)
        command = [
            soffice, "--headless", "--norestore", "--nologo", "--nodefault", "--nolockcheck",
            f"-env:UserInstallation={Path(_profile_dir()).as_uri()}",
            "--convert-to", "pdf", "--outdir", folder, str(docx),
        ]
        with _lock:
            try:
                done = subprocess.run(command, capture_output=True, text=True, timeout=timeout)
            except subprocess.TimeoutExpired:
                raise RenderError(f"LibreOffice took longer than {timeout:g} s to convert the document.") from None
            except OSError as exc:
                raise RenderError(f"Could not start LibreOffice ({soffice}): {exc}") from None
        pdf = Path(folder) / "document.pdf"
        if done.returncode != 0 or not pdf.is_file():
            detail = (done.stderr or done.stdout or "").strip()[-500:]
            raise RenderError(f"LibreOffice could not convert the document (exit {done.returncode}). {detail}".strip())
        return pdf.read_bytes()


def _squash(text: str) -> str:
    return _IGNORED.sub("", text).replace("‑", "-").casefold()


@dataclass
class Rendering:
    """A document laid out as pages."""

    pdf: bytes
    #: The text of each page, in order: body, headers, footers and footnotes.
    pages: list[str] = field(default_factory=list)

    @property
    def page_count(self) -> int:
        return len(self.pages)

    def page_of(self, text: str) -> int | None:
        """The page (from 1) where ``text`` first appears, or None. Spaces,
        line breaks and soft hyphens are ignored, and so is case, since PDF
        text breaks lines and may hyphenate words."""
        wanted = _squash(text)
        if not wanted:
            return None
        for number, page in enumerate(self.pages, start=1):
            if wanted in _squash(page):
                return number
        # Text running over a page break.
        joined, starts = "", []
        for page in self.pages:
            starts.append(len(joined))
            joined += _squash(page)
        at = joined.find(wanted)
        if at < 0:
            return None
        return max(number for number, start in enumerate(starts, start=1) if start <= at)


def read_pdf(pdf: bytes) -> Rendering:
    """Page text of a PDF."""
    pdfium = _pdfium()
    document = pdfium.PdfDocument(pdf)
    try:
        pages = []
        for index in range(len(document)):
            page = document[index]
            textpage = page.get_textpage()
            try:
                pages.append(textpage.get_text_range())
            finally:
                textpage.close()
                page.close()
    finally:
        document.close()
    return Rendering(pdf, pages)


def render(source: Source, *, timeout: float = 120) -> Rendering:
    """Lay out a document (a ``WordDocument``, a path or .docx bytes) as pages."""
    _pdfium()  # fail before the slow conversion when pypdfium2 is missing
    return read_pdf(to_pdf(source, timeout=timeout))
