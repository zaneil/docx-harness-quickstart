"""WordPackage: the unzipped .docx.

A .docx file is a zip archive of "parts" (mostly XML). ``WordPackage`` holds
every part in memory, parses XML parts on demand, and zips everything back up
on save. Parts that were never modified are written back byte-for-byte.
"""

from __future__ import annotations

import io
import os
import zipfile
from importlib import resources
from pathlib import Path

from lxml import etree

from .ooxml import R_NS

CONTENT_TYPES = "[Content_Types].xml"
ROOT_RELS = "_rels/.rels"
DEFAULT_MAIN_PART = "word/document.xml"
PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_DOCUMENT_REL = R_NS + "/officeDocument"

MAX_ENTRIES = 5_000
MAX_TOTAL_BYTES = 300 * 1024 * 1024


class PackageError(ValueError):
    """The file is not a usable .docx package, or a part is invalid."""


def parse_xml(data: bytes | str) -> etree._Element:
    """Parse XML defensively: no entity expansion, no network, no DTDs."""
    if isinstance(data, str):
        data = data.encode("utf-8")
    parser = etree.XMLParser(resolve_entities=False, no_network=True, load_dtd=False, huge_tree=False)
    try:
        root = etree.fromstring(data, parser)
    except etree.XMLSyntaxError as exc:
        raise PackageError(f"XML is not well-formed: {exc}") from exc
    if root.getroottree().docinfo.doctype:
        raise PackageError("DTD declarations are not allowed in OOXML parts")
    return root


def _check_name(name: str) -> None:
    parts = name.replace("\\", "/").split("/")
    if name.startswith(("/", "\\")) or ".." in parts or (len(name) > 1 and name[1] == ":"):
        raise PackageError(f"Unsafe path in archive: {name!r}")


def _is_xml(name: str) -> bool:
    return name.endswith((".xml", ".rels"))


class WordPackage:
    """In-memory view of the parts inside a .docx file."""

    def __init__(self, parts: dict[str, bytes]):
        if CONTENT_TYPES not in parts:
            raise PackageError("Not a .docx package: [Content_Types].xml is missing")
        self._parts = dict(parts)
        # name -> (parsed root, serialisation at parse time)
        self._trees: dict[str, tuple[etree._Element, bytes]] = {}
        if self.main_document_part not in self._parts:
            raise PackageError("Not a Word document: main document part is missing")

    # ------------------------------------------------------------------ open

    @classmethod
    def open(
        cls,
        path: str | os.PathLike | None = None,
        *,
        max_total_bytes: int = MAX_TOTAL_BYTES,
        max_entries: int = MAX_ENTRIES,
    ) -> "WordPackage":
        """Open a .docx file, or the bundled blank template when ``path`` is None."""
        if path is None:
            data = resources.files("docx_harness.document").joinpath("templates/blank.docx").read_bytes()
        else:
            data = Path(path).read_bytes()
        return cls.from_bytes(data, max_total_bytes=max_total_bytes, max_entries=max_entries)

    @classmethod
    def from_bytes(
        cls,
        data: bytes,
        *,
        max_total_bytes: int = MAX_TOTAL_BYTES,
        max_entries: int = MAX_ENTRIES,
    ) -> "WordPackage":
        parts: dict[str, bytes] = {}
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                infos = zf.infolist()
                if len(infos) > max_entries:
                    raise PackageError(f"Archive has too many entries ({len(infos)})")
                total = 0
                for info in infos:
                    if info.is_dir():
                        continue
                    _check_name(info.filename)
                    with zf.open(info) as fh:
                        # Read at most one byte past the budget so a lying
                        # header cannot make us inflate a zip bomb.
                        blob = fh.read(max_total_bytes - total + 1)
                    total += len(blob)
                    if total > max_total_bytes:
                        raise PackageError("Archive is too large when uncompressed")
                    parts[info.filename] = blob
        except zipfile.BadZipFile as exc:
            raise PackageError(
                "Not a valid .docx file (legacy .doc files are not supported)"
            ) from exc
        return cls(parts)

    # ----------------------------------------------------------------- parts

    @property
    def main_document_part(self) -> str:
        """Name of the main document part, resolved through the package rels."""
        if ROOT_RELS in self._parts:
            try:
                rels = self.xml(ROOT_RELS)
            except PackageError:
                return DEFAULT_MAIN_PART
            for rel in rels.iter(f"{{{PKG_RELS_NS}}}Relationship"):
                if rel.get("Type") == OFFICE_DOCUMENT_REL:
                    return rel.get("Target", DEFAULT_MAIN_PART).lstrip("/")
        return DEFAULT_MAIN_PART

    @property
    def styles_part(self) -> str | None:
        folder = self.main_document_part.rsplit("/", 1)[0] if "/" in self.main_document_part else ""
        name = f"{folder}/styles.xml" if folder else "styles.xml"
        return name if name in self._parts else None

    def list_parts(self) -> list[str]:
        return list(self._parts)

    def __contains__(self, name: str) -> bool:
        return name in self._parts

    def _require(self, name: str) -> None:
        if name not in self._parts:
            raise PackageError(f"No part named {name!r}. Parts: {', '.join(self._parts)}")

    def xml(self, name: str) -> etree._Element:
        """Parsed root element of an XML part. Mutations are saved automatically."""
        self._require(name)
        if name not in self._trees:
            if not _is_xml(name):
                raise PackageError(f"{name!r} is not an XML part")
            root = parse_xml(self._parts[name])
            self._trees[name] = (root, etree.tostring(root))
        return self._trees[name][0]

    def part_bytes(self, name: str) -> bytes:
        """Current bytes of a part. Unmodified parts return their original bytes."""
        self._require(name)
        if name in self._trees:
            root, snapshot = self._trees[name]
            if etree.tostring(root) != snapshot:
                return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        return self._parts[name]

    def read_xml(self, name: str, pretty: bool = False) -> str:
        """Current XML source of a part as text."""
        root = self.xml(name)
        if pretty:
            return etree.tostring(copy_for_pretty(root), pretty_print=True, encoding="unicode")
        return etree.tostring(root, encoding="unicode")

    def write_xml(self, name: str, xml: str) -> None:
        """Replace an existing XML part. The XML must be well-formed."""
        self._require(name)
        if not _is_xml(name):
            raise PackageError(f"{name!r} is not an XML part")
        root = parse_xml(xml)
        self._parts[name] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        self._trees.pop(name, None)

    # ------------------------------------------------------------------ save

    def to_bytes(self) -> bytes:
        names = [CONTENT_TYPES] + [n for n in self._parts if n != CONTENT_TYPES]
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in names:
                zf.writestr(name, self.part_bytes(name))
        return buffer.getvalue()

    def save(self, path: str | os.PathLike) -> Path:
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_name(target.name + ".tmp")
        tmp.write_bytes(self.to_bytes())
        os.replace(tmp, target)
        return target


def copy_for_pretty(root: etree._Element) -> etree._Element:
    """Copy of a tree that is safe to indent (the original is left untouched)."""
    import copy

    return copy.deepcopy(root)
