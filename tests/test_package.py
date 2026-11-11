from __future__ import annotations

import io
import zipfile

import pytest

from word_agent import PackageError, WordDocument, WordPackage

from conftest import W, make_docx, template_parts, zip_parts


def test_template_opens_and_lists_parts():
    package = WordPackage.open()
    assert package.main_document_part == "word/document.xml"
    assert package.styles_part == "word/styles.xml"
    assert "[Content_Types].xml" in package.list_parts()


def test_round_trip_untouched_parts_are_byte_identical(tmp_path):
    original = template_parts()
    doc = WordDocument.open()
    doc.list_styles()  # parses styles.xml without changing it
    doc.insert_paragraph("Hello")
    path = doc.save(tmp_path / "out.docx")

    with zipfile.ZipFile(path) as zf:
        assert zf.namelist()[0] == "[Content_Types].xml"
        saved = {name: zf.read(name) for name in zf.namelist()}
    assert set(saved) == set(original)
    for name in original:
        if name == "word/document.xml":
            assert saved[name] != original[name]
        else:
            assert saved[name] == original[name], name

    assert WordDocument.open(path).text() == "Hello"


def test_unknown_parts_survive(tmp_path):
    data = make_docx("<w:p/>", {"word/media/image1.png": b"\x89PNG fake", "customXml/item1.xml": b"<a/>"})
    package = WordPackage.from_bytes(data)
    path = package.save(tmp_path / "out.docx")
    with zipfile.ZipFile(path) as zf:
        assert zf.read("word/media/image1.png") == b"\x89PNG fake"
        assert zf.read("customXml/item1.xml") == b"<a/>"


@pytest.mark.parametrize("name", ["../evil.xml", "/abs.xml", "word/../../x.xml", "C:\\x.xml"])
def test_rejects_path_traversal(name):
    parts = template_parts()
    parts[name] = b"<a/>"
    with pytest.raises(PackageError, match="Unsafe path"):
        WordPackage.from_bytes(zip_parts(parts))


def test_rejects_oversized_archive():
    parts = template_parts()
    parts["word/big.xml"] = b"<a>" + b"x" * 200_000 + b"</a>"
    with pytest.raises(PackageError, match="too large"):
        WordPackage.from_bytes(zip_parts(parts), max_total_bytes=100_000)


def test_rejects_too_many_entries():
    with pytest.raises(PackageError, match="too many entries"):
        WordPackage.from_bytes(zip_parts(template_parts()), max_entries=3)


def test_rejects_entity_payload():
    bomb = (
        '<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;">]>'
        f"<w:document {W}><w:body><w:p><w:r><w:t>&lol2;</w:t></w:r></w:p></w:body></w:document>"
    ).encode()
    parts = template_parts()
    parts["word/document.xml"] = bomb
    with pytest.raises(PackageError, match="DTD"):
        WordDocument(WordPackage.from_bytes(zip_parts(parts)))


def test_rejects_non_docx():
    with pytest.raises(PackageError, match="Not a valid .docx"):
        WordPackage.from_bytes(b"\xd0\xcf\x11\xe0 legacy doc")
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zf:
        zf.writestr("hello.txt", "hi")
    with pytest.raises(PackageError, match="Content_Types"):
        WordPackage.from_bytes(buffer.getvalue())


def test_write_xml_validates_and_only_existing_parts():
    package = WordPackage.open()
    with pytest.raises(PackageError, match="not well-formed"):
        package.write_xml("word/settings.xml", "<w:settings>")
    with pytest.raises(PackageError, match="No part named"):
        package.write_xml("word/new.xml", "<a/>")
    package.write_xml("word/settings.xml", f"<w:settings {W}/>")
    assert "w:settings" in package.read_xml("word/settings.xml")
