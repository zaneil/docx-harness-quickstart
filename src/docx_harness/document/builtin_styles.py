"""Definitions of Word's well-known styles.

One source for two uses: ``scripts/build_template.py`` writes some of them
into the bundled blank template, and ``styling.ensure_style`` adds one to a
document that is asked for a well-known style its file does not define (Word
keeps unused built-in styles out of ``styles.xml``).

The XML is written without namespace declarations; ``BuiltinStyle.xml``
returns a ``w:style`` element as text for embedding in a ``w:styles`` root.
"""

from __future__ import annotations

from dataclasses import dataclass

DOC_DEFAULTS = """<w:docDefaults>
<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>"""


@dataclass(frozen=True)
class BuiltinStyle:
    style_id: str
    name: str
    type: str  # paragraph, character, table or numbering
    body: str  # everything after w:name; "{num_id}" stands for the list's numbering id
    list_kind: str | None = None  # "bullet" or "number" for a list style
    default: bool = False

    def xml(self, num_id: int | None = None, style_id: str | None = None) -> str:
        default = ' w:default="1"' if self.default else ""
        body = self.body.replace("{num_id}", str(num_id if num_id is not None else 0))
        return (
            f'<w:style w:type="{self.type}"{default} w:styleId="{style_id or self.style_id}">'
            f'<w:name w:val="{self.name}"/>{body}</w:style>'
        )


def _heading(level: int, size: int, color: str = "2F5496", italic: bool = False) -> BuiltinStyle:
    emphasis = "<w:i/><w:iCs/>" if italic else ""
    return BuiltinStyle(
        f"Heading{level}", f"heading {level}", "paragraph",
        f"""<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>
<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="{240 if level == 1 else 160}" w:after="80"/><w:outlineLvl w:val="{level - 1}"/></w:pPr>
<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/>{emphasis}<w:color w:val="{color}"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>
""",
    )


def _list(style_id: str, name: str, kind: str, level: int = 0) -> BuiltinStyle:
    return BuiltinStyle(
        style_id, name, "paragraph",
        f"""<w:basedOn w:val="Normal"/><w:uiPriority w:val="36"/><w:qFormat/>
<w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="{{num_id}}"/></w:numPr><w:spacing w:after="60"/><w:contextualSpacing/></w:pPr>
""",
        list_kind=kind,
    )


#: The default style of each type, present in every file Word writes.
BASE = [
    BuiltinStyle("Normal", "Normal", "paragraph", "<w:qFormat/>", default=True),
    BuiltinStyle("DefaultParagraphFont", "Default Paragraph Font", "character",
                 '<w:uiPriority w:val="1"/><w:semiHidden/>', default=True),
    BuiltinStyle("TableNormal", "Normal Table", "table", """<w:uiPriority w:val="99"/><w:semiHidden/>
<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
""", default=True),
    BuiltinStyle("NoList", "No List", "numbering", '<w:uiPriority w:val="99"/><w:semiHidden/>', default=True),
]

#: The styles written into the blank template, in file order.
TEMPLATE = [
    BuiltinStyle("Title", "Title", "paragraph", """<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>
<w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>
<w:rPr><w:rFonts w:ascii="Calibri Light" w:hAnsi="Calibri Light" w:cs="Calibri Light"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>
"""),
    BuiltinStyle("Subtitle", "Subtitle", "paragraph", """<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/>
<w:rPr><w:color w:val="5A5A5A"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
"""),
    _heading(1, 32),
    _heading(2, 26),
    _heading(3, 24, "1F3763"),
    _list("ListBullet", "List Bullet", "bullet"),
    _list("ListNumber", "List Number", "number"),
    BuiltinStyle("Quote", "Quote", "paragraph", """<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/>
<w:pPr><w:spacing w:before="200"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr>
<w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>
"""),
    BuiltinStyle("TableGrid", "Table Grid", "table", """<w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/>
<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:tblPr><w:tblBorders>
<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
</w:tblBorders></w:tblPr>
"""),
]

#: Created on demand only.
EXTRA = [
    _heading(4, 22, italic=True),
    _heading(5, 22),
    _heading(6, 22, "1F3763"),
    _heading(7, 22, "1F3763", italic=True),
    _heading(8, 21, "272727"),
    _heading(9, 21, "272727", italic=True),
    _list("ListBullet2", "List Bullet 2", "bullet", 1),
    _list("ListBullet3", "List Bullet 3", "bullet", 2),
    _list("ListNumber2", "List Number 2", "number", 1),
    _list("ListNumber3", "List Number 3", "number", 2),
    BuiltinStyle("Caption", "caption", "paragraph", """<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/>
<w:pPr><w:spacing w:after="200" w:line="240" w:lineRule="auto"/></w:pPr>
<w:rPr><w:i/><w:iCs/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
"""),
    BuiltinStyle("Header", "header", "paragraph", """<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
<w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
"""),
    BuiltinStyle("Footer", "footer", "paragraph", """<w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
<w:pPr><w:tabs><w:tab w:val="center" w:pos="4680"/><w:tab w:val="right" w:pos="9360"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
"""),
    BuiltinStyle("Hyperlink", "Hyperlink", "character", """<w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>
<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>
"""),
]

ALL = BASE + TEMPLATE + EXTRA
