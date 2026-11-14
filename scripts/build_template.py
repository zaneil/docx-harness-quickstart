"""Regenerate src/docx_harness/document/templates/blank.docx.

The template is a minimal, hand-written Word package with the styles the
default tools rely on: Normal, Title, Subtitle, Heading 1-3, List Bullet,
List Number, Quote and Table Grid. The style and list definitions come from
``docx_harness.document.builtin_styles`` and ``numbering``, which also add
them to documents that lack them.

Run: uv run python scripts/build_template.py
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from docx_harness.document.builtin_styles import BASE, DOC_DEFAULTS, TEMPLATE
from docx_harness.document.numbering import abstract_xml

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

CONTENT_TYPES = DECL + """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

ROOT_RELS = DECL + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOCUMENT_RELS = DECL + """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

DOCUMENT = DECL + f"""<w:document {W}>
<w:body>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body>
</w:document>"""

SETTINGS = DECL + f"""<w:settings {W}>
<w:defaultTabStop w:val="720"/>
<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
</w:settings>"""


STYLES = (
    DECL + f"<w:styles {W}>\n" + DOC_DEFAULTS + "\n"
    + "\n".join(style.xml({"bullet": 1, "number": 2}.get(style.list_kind)) for style in BASE + TEMPLATE)
    + "\n</w:styles>"
)

NUMBERING = DECL + f"""<w:numbering {W}>
{abstract_xml("bullet", 0)}
{abstract_xml("number", 1)}
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>"""

PARTS = {
    "[Content_Types].xml": CONTENT_TYPES,
    "_rels/.rels": ROOT_RELS,
    "word/document.xml": DOCUMENT,
    "word/_rels/document.xml.rels": DOCUMENT_RELS,
    "word/styles.xml": STYLES,
    "word/numbering.xml": NUMBERING,
    "word/settings.xml": SETTINGS,
}


def main() -> None:
    target = Path(__file__).resolve().parents[1] / "src/docx_harness/document/templates/blank.docx"
    target.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, xml in PARTS.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))  # reproducible output
            info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(info, xml.encode("utf-8"))
    print(f"wrote {target}")


if __name__ == "__main__":
    main()
