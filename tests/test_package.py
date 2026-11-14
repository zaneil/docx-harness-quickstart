from __future__ import annotations

import io
import zipfile

import pytest

from docx_harness import PackageError, WordDocument, WordPackage

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


def test_add_part_registers_content_type_and_relationship(tmp_path):
    package = WordPackage.open()
    main = package.main_document_part
    name = package.next_part_name("word/media/image", ".png")
    assert name == "word/media/image1.png"
    package.add_part(name, b"\x89PNG fake", "image/png")
    rel_id = package.relate(main, "image", name)
    assert package.relate(main, "image", name) == rel_id  # reused, not duplicated
    link = package.relate(main, "hyperlink", "https://example.com/a?b=1&c=2", external=True)
    assert link != rel_id
    assert package.next_part_name("word/media/image", ".png") == "word/media/image2.png"
    with pytest.raises(PackageError):
        package.add_part(name, b"again", "image/png")
    with pytest.raises(PackageError):
        package.add_part("../evil.xml", b"<a/>", "application/xml")

    again = WordPackage.from_bytes(package.to_bytes())
    assert again.part_bytes(name) == b"\x89PNG fake"
    assert again.related_parts("image") == [name]
    assert again.relationship(main, link) == {
        "id": link, "type": again.relationships()[-1]["type"], "target": "https://example.com/a?b=1&c=2", "external": True,
    }
    types = again.read_xml("[Content_Types].xml")
    assert 'Extension="png"' in types and types.count("image/png") == 1
    assert again.styles_part == "word/styles.xml" and again.numbering_part == "word/numbering.xml"


def test_relate_creates_the_rels_part_of_a_new_xml_part():
    package = WordPackage.open()
    header = package.add_part(
        "word/header1.xml", f"<w:hdr {W}><w:p/></w:hdr>",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    )
    package.relate(package.main_document_part, "header", header)
    package.relate(header, "hyperlink", "https://example.com", external=True)
    again = WordPackage.from_bytes(package.to_bytes())
    assert again.related_parts("header") == [header]
    assert "word/_rels/header1.xml.rels" in again.list_parts()
    assert again.relationships(header)[0]["external"]
    assert 'PartName="/word/header1.xml"' in again.read_xml("[Content_Types].xml")
