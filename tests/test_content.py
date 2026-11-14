"""Styles created on demand, lists, hyperlinks, images, headers and footers."""

from __future__ import annotations

import struct
import zlib

import pytest

from docx_harness import DocumentError, StyleError, WordDocument, WordPackage
from docx_harness.document.ooxml import qn

from conftest import W, doc_from_body, reopen_with_python_docx, template_parts, zip_parts

BARE_STYLES = (
    f'<w:styles {W}><w:style w:type="paragraph" w:default="1" w:styleId="Standard">'
    '<w:name w:val="Standard"/></w:style></w:styles>'
).encode()


def bare_doc(body_xml: str = "", drop: tuple[str, ...] = ()) -> WordDocument:
    """A document whose styles part defines one style under an unusual id,
    optionally without some parts (their relationships are left dangling,
    as in files written by other tools)."""
    parts = template_parts()
    parts["word/styles.xml"] = BARE_STYLES
    parts["word/document.xml"] = f"<w:document {W}><w:body>{body_xml}<w:sectPr/></w:body></w:document>".encode()
    for name in drop:
        del parts[name]
    return WordDocument(WordPackage.from_bytes(zip_parts(parts)))


def numbers(document, tmp_path):
    """(numId, ilvl) of each paragraph as python-docx reads the saved file."""
    out = []
    for p in reopen_with_python_docx(document, tmp_path).paragraphs:
        pr = p._p.pPr.numPr if p._p.pPr is not None else None
        out.append((pr.numId.val, pr.ilvl.val) if pr is not None else None)
    return out


# -------------------------------------------------------------------- styles

def test_well_known_style_is_added_when_the_file_lacks_it(tmp_path):
    document = bare_doc()
    assert [s["name"] for s in document.list_styles()] == ["Standard"]
    block = document.insert_paragraph("Deep heading", "Heading 4")
    document.insert_paragraph("Figure 1", "caption")
    again = document.insert_paragraph("Another", "heading4")

    names = {s["name"]: s["id"] for s in document.list_styles("paragraph")}
    assert names["heading 4"] == "Heading4" and "caption" in names
    assert len(document.list_styles("paragraph")) == 3  # asked for twice, added once
    style = next(s for s in document._styles if s.get(qn("w:styleId")) == "Heading4")
    assert style.find(qn("w:basedOn")).get(qn("w:val")) == "Standard"  # the file's own default style
    assert style.find(qn("w:next")) is None
    assert document.read(block, 1)["blocks"][0]["style"] == "heading 4"

    reopened = reopen_with_python_docx(document, tmp_path)
    assert [p.style.name for p in reopened.paragraphs] == ["Heading 4", "Caption", "Heading 4"]
    assert again != block


def test_unknown_style_still_fails_and_points_at_create_style(doc):
    with pytest.raises(StyleError, match="create_style"):
        doc.insert_paragraph("x", "Fancy Callout")


def test_styles_part_is_created_when_missing(tmp_path):
    document = bare_doc(drop=("word/styles.xml",))
    assert document.list_styles() == []
    document.insert_paragraph("Title here", "Title")
    reopened = reopen_with_python_docx(document, tmp_path)
    assert reopened.paragraphs[0].style.name == "Title"
    assert WordPackage.open(tmp_path / "out.docx").read_xml("[Content_Types].xml").count("/word/styles.xml") == 1


def test_missing_table_style_is_created(tmp_path):
    document = bare_doc()
    table = document.insert_table([["a", "b"]])
    assert "tblBorders" not in document.read_part_xml("word/document.xml")
    assert reopen_with_python_docx(document, tmp_path).tables[0].style.name == "Table Grid"
    assert document.read(table, 1)["blocks"][0]["rows"] == [["a", "b"]]


def test_create_style(doc, tmp_path):
    name = doc.create_style("Call Out", based_on="Quote", bold=True, color="C00000", size_pt=14,
                            alignment="left", space_after_pt=12)
    assert name == "Call Out"
    block = doc.insert_paragraph("Mind the gap", "call out")
    info = doc.read(block, 1, include_formatting=True, resolve_styles=True)["blocks"][0]
    assert info["style"] == "Call Out"
    assert info["style_format"]["run"] == {"bold": True, "italic": True, "color": "C00000", "size_pt": 14.0, "font": "Calibri"}
    assert info["style_format"]["paragraph"]["alignment"] == "left"
    style = reopen_with_python_docx(doc, tmp_path).styles["Call Out"]
    assert style.base_style.name == "Quote" and style.font.bold and style.paragraph_format.space_after.pt == 12

    with pytest.raises(StyleError, match="already exists"):
        doc.create_style("call-out")
    with pytest.raises(StyleError, match="cannot set"):
        doc.create_style("Shout", style_type="character", alignment="center")
    doc.create_style("Shout", based_on=None, style_type="character", bold=True)
    assert "Shout" in {s["name"] for s in doc.list_styles("character")}


# --------------------------------------------------------------------- lists

def test_a_new_numbered_list_starts_at_one_and_items_carry_it_on(doc, tmp_path):
    first = [doc.insert_paragraph(t, "List Number") for t in ("one", "two")]
    doc.insert_paragraph("Something else")
    second = [doc.insert_paragraph(t, "List Number") for t in ("uno", "dos")]
    middle = doc.insert_paragraph("one and a half", "List Number", after=first[0])

    found = numbers(doc, tmp_path)
    assert found[0] == found[1] == found[2] and found[3] is None
    assert found[4] == found[5] and found[4] != found[0]
    assert all(level == 0 for _, level in (f for f in found if f))
    assert doc.read(middle, 1)["blocks"][0]["list"] == {"kind": "number", "level": 0}
    xml = doc.read_part_xml("word/numbering.xml")
    assert xml.count("<w:startOverride") == 2
    assert doc.outline()[0]["list"] == {"kind": "number", "level": 0}
    assert second


def test_nested_bullets_do_not_break_the_numbered_list_around_them(doc, tmp_path):
    doc.insert_paragraph("Step one", "List Number")
    doc.insert_paragraph("detail", "List Bullet", level=1)
    doc.insert_paragraph("Step two", "List Number")
    doc.insert_paragraph("sub step", "List Number", level=1)
    found = numbers(doc, tmp_path)
    assert found[0] == found[2] and found[3] == (found[0][0], 1)
    assert found[1][1] == 1 and found[1][0] != found[0][0]
    blocks = doc.read()["blocks"]
    assert [b["list"] for b in blocks] == [
        {"kind": "number", "level": 0}, {"kind": "bullet", "level": 1},
        {"kind": "number", "level": 0}, {"kind": "number", "level": 1},
    ]


def test_an_item_put_in_front_of_a_list_joins_it(doc, tmp_path):
    intro = doc.insert_paragraph("Intro")
    doc.insert_paragraph("second", "List Number")
    doc.insert_paragraph("first", "List Number", after=intro)
    found = numbers(doc, tmp_path)
    assert found[1] == found[2]


def test_set_list_level_and_restart(doc, tmp_path):
    ids = [doc.insert_paragraph(t, "List Number") for t in "abcde"]
    assert doc.set_list_level(ids[1], 1) == 1
    assert doc.set_list_level(ids[1], 1) == 0
    assert doc.restart_numbering(ids[3], start=5) == 2
    found = numbers(doc, tmp_path)
    assert found[1] == (found[0][0], 1)
    assert found[3] == found[4] and found[3][0] != found[0][0]
    assert '<w:startOverride w:val="5"/>' in doc.read_part_xml("word/numbering.xml")

    plain = doc.insert_paragraph("plain")
    bullet = doc.insert_paragraph("dot", "List Bullet")
    with pytest.raises(DocumentError, match="not a list item"):
        doc.set_list_level(plain, 1)
    with pytest.raises(DocumentError, match="between 0 and 8"):
        doc.set_list_level(bullet, 9)
    with pytest.raises(DocumentError, match="not a numbered list item"):
        doc.restart_numbering(bullet)
    with pytest.raises(DocumentError, match="only applies to list items"):
        doc.insert_paragraph("x", level=2)
    assert doc.insert_paragraph("fine", level=0)


def test_lists_in_a_file_without_list_styles_or_a_numbering_part(tmp_path):
    document = bare_doc(drop=("word/numbering.xml",))
    document.insert_paragraph("a", "List Bullet")
    document.insert_paragraph("b", "List Bullet 2")
    document.insert_paragraph("c", "List Number")
    blocks = document.read()["blocks"]
    assert [b["list"] for b in blocks] == [
        {"kind": "bullet", "level": 0}, {"kind": "bullet", "level": 1}, {"kind": "number", "level": 0},
    ]
    reopened = reopen_with_python_docx(document, tmp_path)
    assert [p.style.name for p in reopened.paragraphs] == ["List Bullet", "List Bullet 2", "List Number"]
    package = WordPackage.open(tmp_path / "out.docx")
    assert package.numbering_part == "word/numbering.xml"
    assert package.read_xml("[Content_Types].xml").count("/word/numbering.xml") == 1
    assert package.read_xml("word/numbering.xml").count("<w:abstractNum ") == 2


def test_list_changes_are_tracked_and_can_be_rejected(tmp_path):
    document = doc_from_body("")
    ids = [document.insert_paragraph(t, "List Number") for t in "abc"]
    before = numbers(document, tmp_path)
    document.track_changes = True
    document.set_list_level(ids[1], 2)
    document.restart_numbering(ids[2])
    assert [r["type"] for r in document.revisions()] == ["format", "format"]
    assert numbers(document, tmp_path) != before
    document.reject_revisions()
    assert numbers(document, tmp_path) == before


# ---------------------------------------------------------------- hyperlinks

def links(document, tmp_path):
    reopened = reopen_with_python_docx(document, tmp_path)
    return [(h.text, h.address) for p in reopened.paragraphs for h in p.hyperlinks]


def test_add_retarget_and_remove_hyperlinks(doc, tmp_path):
    block = doc.insert_paragraph("See the docs and the docs mirror, or write to us.")
    assert doc.add_hyperlink(block, "https://example.com/docs?a=1&b=2", match="the docs") == 2
    assert doc.add_hyperlink(block, "mailto:team@example.com", match="write to us") == 1
    assert links(doc, tmp_path) == [
        ("the docs", "https://example.com/docs?a=1&b=2"), ("the docs", "https://example.com/docs?a=1&b=2"),
        ("write to us", "mailto:team@example.com"),
    ]
    info = doc.read(block, 1, include_formatting=True)["blocks"][0]
    assert info["text"] == "See the docs and the docs mirror, or write to us."
    assert info["links"][2] == {"text": "write to us", "url": "mailto:team@example.com"}
    assert "Hyperlink" in {s["name"] for s in doc.list_styles("character")}
    assert doc.read_part_xml("word/_rels/document.xml.rels").count("docs?a=1") == 1  # one relationship, reused

    assert doc.add_hyperlink(block, "https://example.org", match="write to us") == 1  # already a link: re-pointed
    assert links(doc, tmp_path)[2] == ("write to us", "https://example.org")
    assert doc.read_part_xml("word/document.xml").count("<w:hyperlink") == 3

    assert doc.remove_hyperlink(block, match="write") == 1
    assert len(links(doc, tmp_path)) == 2
    assert doc.remove_hyperlink(block) == 2
    assert links(doc, tmp_path) == []
    assert "rStyle" not in doc.read_part_xml("word/document.xml")
    assert doc.text() == "See the docs and the docs mirror, or write to us."


def test_hyperlink_errors_and_whole_block(doc, tmp_path):
    block = doc.insert_paragraph("Home page")
    for bad in ("example.com", "javascript:alert(1)", "file:///etc/passwd", "https://", "https://a b.com"):
        with pytest.raises(DocumentError, match="Invalid link address"):
            doc.add_hyperlink(block, bad)
    with pytest.raises(DocumentError, match="was not found"):
        doc.add_hyperlink(block, "https://example.com", match="missing")
    with pytest.raises(DocumentError, match="No link"):
        doc.remove_hyperlink(block)
    assert doc.add_hyperlink(block, "https://example.com") == 1
    assert links(doc, tmp_path) == [("Home page", "https://example.com")]
    doc.replace_text("Home", "Start")
    assert links(doc, tmp_path) == [("Start page", "https://example.com")]


def test_link_across_formatted_runs_keeps_their_formatting(tmp_path):
    document = doc_from_body(
        "<w:p><w:r><w:t xml:space=\"preserve\">Read </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>this</w:t></w:r>"
        "<w:r><w:t xml:space=\"preserve\"> guide now</w:t></w:r></w:p>"
    )
    block = document.outline()[0]["id"]
    document.add_hyperlink(block, "https://example.com", match="this guide")
    runs = document.read(block, 1, include_formatting=True)["blocks"][0]["runs"]
    assert [(r["text"], r.get("bold", False)) for r in runs] == [
        ("Read ", False), ("this", True), (" guide", False), (" now", False)]
    assert links(document, tmp_path) == [("this guide", "https://example.com")]


def test_tracked_hyperlink_can_be_rejected_or_accepted(tmp_path):
    for accept in (False, True):
        document = doc_from_body("<w:p><w:r><w:t>Visit our site today</w:t></w:r></w:p>")
        document.track_changes = True
        block = document.outline()[0]["id"]
        document.add_hyperlink(block, "https://example.com", match="our site")
        assert [r["type"] for r in document.revisions()] == ["delete", "insert"]
        assert document.text() == "Visit our site today"
        if accept:
            document.accept_revisions()
            assert links(document, tmp_path) == [("our site", "https://example.com")]
            document.track_changes = True
            document.remove_hyperlink(block)
            assert "links" not in document.read(block, 1)["blocks"][0]  # pending: reads as plain text
            document.reject_revisions()
            assert links(document, tmp_path) == [("our site", "https://example.com")]
            document.remove_hyperlink(block)
            document.accept_revisions()
            assert links(document, tmp_path) == []
            assert "<w:hyperlink" not in document.read_part_xml("word/document.xml")
        else:
            document.reject_revisions()
            assert links(document, tmp_path) == []
        assert document.text() == "Visit our site today"
        assert document.revisions() == []


# -------------------------------------------------------------------- images

def png(width: int, height: int, dpi: int | None = None) -> bytes:
    def chunk(kind: bytes, payload: bytes) -> bytes:
        return struct.pack(">I", len(payload)) + kind + payload + struct.pack(">I", zlib.crc32(kind + payload))

    physical = chunk(b"pHYs", struct.pack(">IIB", round(dpi / 0.0254), round(dpi / 0.0254), 1)) if dpi else b""
    rows = b"".join(b"\x00" + b"\x7f" * (3 * width) for _ in range(height))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + physical + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def jpeg_header(width: int, height: int) -> bytes:
    app0 = b"JFIF\x00\x01\x01\x01" + struct.pack(">HH", 72, 72) + b"\x00\x00"
    sof = b"\x08" + struct.pack(">HH", height, width) + b"\x03\x01\x11\x00\x02\x11\x01\x03\x11\x01"
    return (b"\xff\xd8\xff\xe0" + struct.pack(">H", len(app0) + 2) + app0
            + b"\xff\xc0" + struct.pack(">H", len(sof) + 2) + sof + b"\xff\xd9")


def test_insert_image_sizes_and_round_trip(doc, tmp_path):
    path = tmp_path / "logo.png"
    path.write_bytes(png(192, 96))
    intro = doc.insert_paragraph("Intro")
    natural = doc.insert_image(path, after=intro, alt_text="Company <logo> & \"mark\"", alignment="center")
    wide = doc.insert_image(png(1920, 960), width_in=None)             # wider than the page: fitted
    fixed = doc.insert_image(path, width_in=3)                          # height follows
    retina = doc.insert_image(png(288, 144, dpi=144))                   # 2x screenshot
    photo = doc.insert_image(jpeg_header(720, 360), height_in=1)

    blocks = {b["id"]: b for b in doc.read()["blocks"]}
    assert blocks[natural]["images"] == [
        {"name": "image1.png", "alt": 'Company <logo> & "mark"', "width_in": 2.0, "height_in": 1.0}]
    assert (blocks[wide]["images"][0]["width_in"], blocks[wide]["images"][0]["height_in"]) == (6.5, 3.25)
    assert (blocks[fixed]["images"][0]["width_in"], blocks[fixed]["images"][0]["height_in"]) == (3.0, 1.5)
    assert blocks[fixed]["images"][0]["name"] == "image1.png"  # same bytes: stored once
    assert (blocks[retina]["images"][0]["width_in"], blocks[retina]["images"][0]["height_in"]) == (2.0, 1.0)
    assert (blocks[photo]["images"][0]["width_in"], blocks[photo]["images"][0]["height_in"]) == (2.0, 1.0)
    assert blocks[photo]["images"][0]["name"].endswith(".jpeg")
    assert doc.outline()[1]["preview"] == '[image: Company <logo> & "mark"]'
    assert [b for b in doc.outline()][0]["preview"] == "Intro"

    reopened = reopen_with_python_docx(doc, tmp_path)
    shapes = reopened.inline_shapes
    assert len(shapes) == 5
    assert (shapes[0].width.inches, shapes[0].height.inches) == (2.0, 1.0)
    assert reopened.paragraphs[1].alignment is not None
    package = WordPackage.open(tmp_path / "out.docx")
    media = [name for name in package.list_parts() if name.startswith("word/media/")]
    assert len(media) == 4 and package.part_bytes("word/media/image1.png") == path.read_bytes()
    ids = [pr.get("id") for pr in package.xml("word/document.xml").iter(qn("wp:docPr"))]
    assert len(set(ids)) == 5

    doc.set_image(natural, width_in=4, alt_text="Logo")
    assert doc.read(natural, 1)["blocks"][0]["images"][0] == {
        "name": "image1.png", "alt": "Logo", "width_in": 4.0, "height_in": 2.0}
    doc.delete_block(wide)
    assert len(reopen_with_python_docx(doc, tmp_path).inline_shapes) == 4


def test_image_errors(doc, tmp_path):
    secret = tmp_path / "notes.png"
    secret.write_text("just text pretending to be a picture")
    with pytest.raises(DocumentError, match="Not a PNG, JPEG or GIF"):
        doc.insert_image(secret)
    with pytest.raises(DocumentError, match="Cannot read the image file"):
        doc.insert_image(tmp_path / "missing.png")
    with pytest.raises(DocumentError, match="damaged"):
        doc.insert_image(b"\x89PNG\r\n\x1a\n\x00")
    with pytest.raises(DocumentError, match="at most 22 inches"):
        doc.insert_image(png(10, 10), width_in=40)
    plain = doc.insert_paragraph("text")
    with pytest.raises(DocumentError, match="has no picture"):
        doc.set_image(plain, width_in=1)
    assert len(doc) == 1 and not [n for n in doc.list_parts() if "media" in n]


def test_rewriting_a_paragraph_keeps_its_picture(doc, tmp_path):
    block = doc.insert_image(png(96, 96))
    doc.set_paragraph_text(block, "Figure 1: a square")
    info = doc.read(block, 1)["blocks"][0]
    assert info["text"] == "Figure 1: a square" and len(info["images"]) == 1
    doc.set_paragraph_text(block, "Figure 1")
    assert doc.read(block, 1)["blocks"][0]["text"] == "Figure 1"
    assert len(reopen_with_python_docx(doc, tmp_path).inline_shapes) == 1


def test_tracked_image_insert_is_a_revision(tmp_path):
    for accept in (True, False):
        document = doc_from_body("<w:p><w:r><w:t>Before</w:t></w:r></w:p>")
        document.track_changes = True
        block = document.insert_image(png(96, 96), alt_text="dot")
        assert [r["type"] for r in document.revisions()] == ["insert", "insert"]
        assert document.read(block, 1)["blocks"][0]["images"][0]["alt"] == "dot"
        # python-docx only looks for pictures outside revision marks.
        assert len(reopen_with_python_docx(document, tmp_path).inline_shapes) == 0
        if accept:
            document.accept_revisions()
            assert len(reopen_with_python_docx(document, tmp_path).inline_shapes) == 1
        else:
            document.reject_revisions()
            assert len(document) == 1 and "drawing" not in document.read_part_xml("word/document.xml")


# -------------------------------------------------------- headers and footers

def test_create_header_and_footer_with_page_numbers(doc, tmp_path):
    doc.insert_paragraph("Body")
    header = doc.set_header("Quarterly report\tConfidential")
    footer = doc.set_footer("Page {PAGE} of {NUMPAGES}", alignment="center")
    cover = doc.set_header("", page="first")
    assert (header, footer, cover) == ("word/header1.xml", "word/footer1.xml", "word/header2.xml")
    assert doc.stories() == [
        {"part": "word/header1.xml", "kind": "header", "paragraphs": ["Quarterly report\tConfidential"], "page": "default"},
        {"part": "word/header2.xml", "kind": "header", "paragraphs": [""], "page": "first"},
        {"part": "word/footer1.xml", "kind": "footer", "paragraphs": ["Page {PAGE} of {NUMPAGES}"], "page": "default"},
    ]
    assert doc.search("confidential")[0]["ref"] == header

    assert doc.format_text(header, match="Confidential", bold=True, color="C00000") == 1
    assert doc.replace_text("Quarterly", "Annual", ref=header) == 1
    assert doc.set_header("Annual report 2026") == header  # same part, new text
    assert len(doc.stories()) == 3

    reopened = reopen_with_python_docx(doc, tmp_path)
    section = reopened.sections[0]
    assert section.header.paragraphs[0].text == "Annual report 2026"
    assert section.header.paragraphs[0].style.name == "Header"
    assert section.different_first_page_header_footer
    assert section.first_page_header.paragraphs[0].text == ""
    assert section.footer.paragraphs[0].alignment is not None
    package = WordPackage.open(tmp_path / "out.docx")
    xml = package.read_xml(footer)
    assert 'w:instr=" PAGE "' in xml and 'w:instr=" NUMPAGES "' in xml
    sect = package.xml("word/document.xml").find(qn("w:body")).find(qn("w:sectPr"))
    assert [child.tag for child in sect][:3] == [qn("w:headerReference"), qn("w:footerReference"), qn("w:headerReference")]
    assert [child.tag for child in sect][-1] == qn("w:titlePg")
    types = package.read_xml("[Content_Types].xml")
    assert types.count("header+xml") == 2 and types.count("footer+xml") == 1


def test_even_page_footer_switches_the_setting_on(doc, tmp_path):
    doc.set_footer("{PAGE}", page="even")
    settings = doc.read_part_xml("word/settings.xml")
    assert "<w:evenAndOddHeaders/>" in settings
    assert settings.index("defaultTabStop") < settings.index("evenAndOddHeaders") < settings.index("w:compat")
    assert reopen_with_python_docx(doc, tmp_path).settings.odd_and_even_pages_header_footer
    with pytest.raises(DocumentError, match="Invalid page"):
        doc.set_footer("x", page="odd")


def test_existing_header_is_rewritten_and_tracked(tmp_path):
    header = f'<w:hdr {W}><w:p><w:r><w:t>Draft v1</w:t></w:r></w:p><w:p><w:r><w:t>Internal</w:t></w:r></w:p></w:hdr>'
    R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    parts = template_parts()
    parts["word/header7.xml"] = header.encode()
    parts["word/document.xml"] = (
        f'<w:document {W} {R}><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p>'
        '<w:sectPr><w:headerReference w:type="default" r:id="rId9"/></w:sectPr></w:body></w:document>'
    ).encode()
    parts["word/_rels/document.xml.rels"] = parts["word/_rels/document.xml.rels"].replace(
        b"</Relationships>",
        b'<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        b'Target="header7.xml"/></Relationships>',
    )
    parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace(
        b"</Types>",
        b'<Override PartName="/word/header7.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        b'wordprocessingml.header+xml"/></Types>',
    )
    document = WordDocument(WordPackage.from_bytes(zip_parts(parts)), track_changes=True, author="Ed")
    assert document.stories()[0]["page"] == "default"
    assert document.set_header("Final") == "word/header7.xml"
    assert document.stories()[0]["paragraphs"] == ["Final"]
    assert {r["part"] for r in document.revisions()} == {"word/header7.xml"}
    assert len(document.list_parts()) == len(parts)

    document.reject_revisions()
    assert document.stories()[0]["paragraphs"] == ["Draft v1", "Internal"]
    document.set_header("Final")
    document.accept_revisions()
    assert reopen_with_python_docx(document, tmp_path).sections[0].header.paragraphs[0].text == "Final"
    assert document.stories()[0]["paragraphs"] == ["Final"]


def test_preview_numbers_follow_restarts_and_nesting(doc):
    for text, style, level in (("a", "List Number", 0), ("b", "List Number", 0), ("b1", "List Number", 1),
                               ("b2", "List Number", 1), ("c", "List Number", 0), ("c1", "List Number", 1),
                               ("dot", "List Bullet", 0)):
        doc.insert_paragraph(text, style, level=level)
    doc.insert_paragraph("break")
    later = doc.insert_paragraph("x", "List Number")
    doc.insert_paragraph("y", "List Number")
    doc.restart_numbering(later, start=7)
    blocks = doc.read(count=50, include_formatting=True, resolve_styles=True)["blocks"]
    assert [b.get("list", {}).get("number") for b in blocks] == [1, 2, 1, 2, 3, 1, None, None, 7, 8]
    assert "number" not in doc.read()["blocks"][0]["list"]


def test_accepting_a_deleted_paragraph_with_fields_or_links_leaves_nothing_behind(tmp_path):
    document = WordDocument.open()
    document.insert_paragraph("Body")
    linked = document.insert_paragraph("Old line with a link")
    document.add_hyperlink(linked, "https://example.com", match="a link")
    footer = document.set_footer("Page {PAGE} of {NUMPAGES}")

    document.track_changes = True
    document.set_footer("Confidential, page {PAGE}")
    document.delete_block(linked)
    assert document.stories()[0]["paragraphs"] == ["Confidential, page {PAGE}"]
    document.accept_revisions()
    assert document.stories()[0]["paragraphs"] == ["Confidential, page {PAGE}"]
    assert document.read_part_xml(footer).count("<w:p>") == 1
    assert document.text() == "Body" and "<w:hyperlink" not in document.read_part_xml("word/document.xml")
    assert reopen_with_python_docx(document, tmp_path).sections[0].footer.paragraphs[0].text.startswith("Confidential")
