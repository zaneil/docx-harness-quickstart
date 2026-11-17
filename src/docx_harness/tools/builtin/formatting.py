"""The "format" group: styles, lists, character and paragraph formatting."""

from __future__ import annotations

from typing import Literal, Optional

from ...document import WordDocument
from ..base import Tool, tool
from .common import as_list


def format_tools(doc: WordDocument) -> list[Tool]:
    @tool
    def set_paragraph_style(blocks: list[str], style: str) -> str:
        """Change which named style blocks use (e.g. body text to 'Heading 2' or 'List Bullet'). For color, size or font use format_text.

        Args:
            blocks: Block ids, or table refs: cell 'b3[0,1]', row 'b3[0,*]', column 'b3[*,1]'.
            style: Style name; list_styles shows what is available.
        """
        count = sum(doc.set_paragraph_style(ref, style) for ref in as_list(blocks))
        if count == 0:
            return (
                f"Nothing changed: these blocks already use style {style!r}. "
                "To change color, size, font or bold, call format_text instead."
            )
        return f"Applied style {style!r} to {count} paragraph(s)."

    @tool
    def create_style(
        name: str,
        based_on: str = "Normal",
        style_type: Literal["paragraph", "character"] = "paragraph",
        font: Optional[str] = None,
        size_pt: Optional[float] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None,
        alignment: Optional[Literal["left", "center", "right", "justify"]] = None,
        space_before_pt: Optional[float] = None,
        space_after_pt: Optional[float] = None,
        line_spacing: Optional[float] = None,
        indent_left_in: Optional[float] = None,
    ) -> str:
        """Define a new named style for a look used in several places (e.g. 'Callout'). Standard styles such as 'Heading 4' or 'Caption' need no definition. Apply it with set_paragraph_style or insert_paragraphs. Alignment, spacing and indent apply to paragraph styles only.

        Args:
            name: Name of the new style.
            based_on: Existing style it inherits from.
            style_type: 'character' styles are defined only; apply character formatting with format_text.
            font: Font family.
            size_pt: Font size.
            bold: Bold.
            italic: Italic.
            color: 6-digit hex or a color name.
            alignment: Paragraph alignment.
            space_before_pt: Space above.
            space_after_pt: Space below.
            line_spacing: Line spacing multiple.
            indent_left_in: Left indent.
        """
        created = doc.create_style(
            name, based_on, style_type, font=font, size_pt=size_pt, bold=bold, italic=italic, color=color,
            alignment=alignment, space_before_pt=space_before_pt, space_after_pt=space_after_pt,
            line_spacing=line_spacing, indent_left_in=indent_left_in,
        )
        return f"Created {style_type} style {created!r}."

    @tool
    def set_list_format(blocks: list[str], level: Optional[int] = None, start_at: Optional[int] = None) -> str:
        """Change list items: nest them under the item above, or restart the numbering. The blocks must already be list items. Give level, start_at or both.

        Args:
            blocks: Ids of the list items.
            level: Nesting level: 0 outer list, 1 sub-items, up to 8.
            start_at: Restart numbering at the first block with this number (usually 1); later items follow on.
        """
        blocks = as_list(blocks)
        if level is None and start_at is None:
            raise ValueError("Give level, start_at or both. Nothing was changed.")
        done = []
        if level is not None:
            count = sum(doc.set_list_level(ref, level) for ref in blocks)
            done.append(f"set {count} item(s) to level {level}")
        if start_at is not None:
            count = doc.restart_numbering(blocks[0], start_at)
            done.append(f"restarted numbering at {start_at} from {blocks[0]} ({count} item(s) renumbered)")
        return (", ".join(done)).capitalize() + "."

    @tool
    def format_text(
        blocks: list[str],
        match: Optional[str] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        underline: Optional[bool] = None,
        strike: Optional[bool] = None,
        color: Optional[str] = None,
        size_pt: Optional[float] = None,
        font: Optional[str] = None,
        highlight: Optional[str] = None,
    ) -> str:
        """Apply character formatting to whole blocks, or only to some text inside them. true turns an effect on, false off.

        Args:
            blocks: Block ids, or table refs: cell 'b3[0,1]', row 'b3[0,*]', column 'b3[*,1]', or a table id for the whole table.
            match: Only format this exact text. Omit to format all text.
            bold: Bold.
            italic: Italic.
            underline: Underline.
            strike: Strikethrough.
            color: 6-digit hex (e.g. '1F4E79') or a color name.
            size_pt: Font size.
            font: Font family.
            highlight: Color name such as 'yellow', or 'none' to remove.
        """
        count = sum(
            doc.format_text(
                ref, match, bold=bold, italic=italic, underline=underline, strike=strike,
                color=color, size_pt=size_pt, font=font, highlight=highlight,
            )
            for ref in as_list(blocks)
        )
        return f"Formatted {count} run(s) in {len(as_list(blocks))} block(s)."

    @tool
    def set_paragraph_format(
        blocks: list[str],
        alignment: Optional[Literal["left", "center", "right", "justify"]] = None,
        space_before_pt: Optional[float] = None,
        space_after_pt: Optional[float] = None,
        line_spacing: Optional[float] = None,
        indent_left_in: Optional[float] = None,
        indent_right_in: Optional[float] = None,
        first_line_indent_in: Optional[float] = None,
        page_break_before: Optional[bool] = None,
        keep_with_next: Optional[bool] = None,
    ) -> str:
        """Set paragraph layout: alignment, spacing, indentation and page breaks.

        Args:
            blocks: Block ids, or table refs: cell 'b3[0,1]', row 'b3[0,*]', column 'b3[*,1]'.
            alignment: Horizontal alignment.
            space_before_pt: Space above.
            space_after_pt: Space below.
            line_spacing: Line spacing multiple, e.g. 1.0, 1.5 or 2.0.
            indent_left_in: Left indent.
            indent_right_in: Right indent.
            first_line_indent_in: Negative for a hanging indent.
            page_break_before: Start the paragraph on a new page.
            keep_with_next: Keep on the same page as the next paragraph.
        """
        count = sum(
            doc.set_paragraph_format(
                ref, alignment=alignment, space_before_pt=space_before_pt, space_after_pt=space_after_pt,
                line_spacing=line_spacing, indent_left_in=indent_left_in, indent_right_in=indent_right_in,
                first_line_indent_in=first_line_indent_in, page_break_before=page_break_before,
                keep_with_next=keep_with_next,
            )
            for ref in as_list(blocks)
        )
        return f"Formatted {count} paragraph(s)."

    return [set_paragraph_style, create_style, set_list_format, format_text, set_paragraph_format]
