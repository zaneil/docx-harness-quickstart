"""The LibreOffice renderer. Most tests need LibreOffice and pypdfium2 and
are skipped without them; install LibreOffice to run them."""

from __future__ import annotations

import pytest

from docx_harness import WordDocument, render
from docx_harness.render import Rendering, RenderError

from test_content import png

needs_renderer = pytest.mark.skipif(not render.available(), reason="LibreOffice or pypdfium2 is not installed")


# ------------------------------------------------------------ always run

def test_page_of_ignores_line_breaks_hyphens_and_case():
    pages = Rendering(b"", ["Annual report\r\nSales rose in every re-\ngion", "Appen­dix\nRaw  figures"])
    assert pages.page_count == 2
    assert pages.page_of("annual REPORT") == 1
    assert pages.page_of("Appendix") == 2
    assert pages.page_of("Raw figures") == 2
    assert pages.page_of("every region") is None  # a real hyphen is kept
    assert pages.page_of("in every re-gionAppendix") == 1  # runs over the page break: where it starts
    assert pages.page_of("missing") is None and pages.page_of("  ") is None


def test_a_missing_soffice_is_a_clear_error(monkeypatch, tmp_path):
    monkeypatch.setenv("DOCX_HARNESS_SOFFICE", str(tmp_path / "nothing"))
    assert render.find_soffice() is None and not render.available()
    with pytest.raises(RenderError, match="DOCX_HARNESS_SOFFICE is set to .* which is not a file. LibreOffice was not found"):
        render.to_pdf(WordDocument.open())


def test_render_rejects_other_objects():
    if render.find_soffice() is None:
        pytest.skip("LibreOffice is not installed")
    with pytest.raises(TypeError, match="pass a WordDocument, a path or .docx bytes"):
        render.to_pdf(42)


# ------------------------------------------------------ need LibreOffice

@needs_renderer
def test_page_breaks_and_sources(tmp_path):
    doc = WordDocument.open()
    doc.insert_paragraph("First page.")
    one = render.render(doc)
    assert one.page_count == 1 and one.pdf.startswith(b"%PDF")
    doc.insert_page_break()
    doc.insert_paragraph("Second page.")
    path = doc.save(tmp_path / "two.docx")
    for source in (doc, path, str(path), path.read_bytes()):
        pages = render.render(source)
        assert pages.page_count == 2
        assert (pages.page_of("First page."), pages.page_of("Second page.")) == (1, 2)


@needs_renderer
def test_sections_headers_and_footnotes_land_on_their_pages():
    doc = WordDocument.open()
    doc.set_header("Acme report")
    intro = doc.insert_paragraph("Introduction text.")
    doc.add_note(intro, "Source: annual report.")
    doc.insert_paragraph("Appendix", "Heading 1")
    doc.insert_paragraph("Raw figures.")
    doc.insert_section_break(intro)
    doc.set_header("Appendix header", section=2)
    pages = render.render(doc)
    assert pages.page_count == 2
    assert "Acme report" in pages.pages[0] and "Acme report" not in pages.pages[1]
    assert "Appendix header" in pages.pages[1] and "Appendix header" not in pages.pages[0]
    assert pages.page_of("Source: annual report.") == 1  # at the foot of its mark's page
    assert pages.page_of("Raw figures.") == 2


@needs_renderer
def test_tables_pictures_and_lists_render(tmp_path):
    picture = tmp_path / "dot.png"
    picture.write_bytes(png(40, 20))
    doc = WordDocument.open()
    doc.insert_table([["Region", "Sales"], ["North", "120"]])
    doc.insert_image(picture, width_in=2, alt_text="dot")
    doc.insert_paragraph("Step one", "List Number")
    text = render.render(doc).pages[0]
    for expected in ("Region", "Sales", "North", "120", "Step one"):
        assert expected in text
