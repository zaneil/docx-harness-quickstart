"""The "tables" group: edit tables that are already in the document.

``insert_table`` is in the write group; these tools are offered only while the
document has a table.
"""

from __future__ import annotations

from typing import Literal, Optional

from ...document import WordDocument
from ..base import Tool, tool
from .common import as_list


def table_tools(doc: WordDocument) -> list[Tool]:
    @tool
    def set_table_cell(table_id: str, row: int, col: int, text: str) -> str:
        """Set the text of one table cell; each line is a paragraph and keeps the cell's formatting. For a small wording change prefer replace_text with block='b3[row,col]'.

        Args:
            table_id: Id of the table block.
            row: Zero-based row.
            col: Zero-based column on the table's grid; any position inside a merged cell edits that cell.
            text: The new cell text.
        """
        doc.set_cell_text(table_id, row, col, text)
        return f"Updated {table_id}[{row},{col}]"

    @tool
    def insert_table_row(table_id: str, cells: list[str], index: Optional[int] = None) -> str:
        """Add a row to a table, copying the layout of the row next to it.

        Args:
            table_id: Id of the table block.
            cells: Cell texts for the new row, left to right.
            index: Zero-based position for the new row. Omit to append.
        """
        return f"Inserted row {doc.insert_table_row(table_id, cells, index)} in {table_id}"

    @tool
    def insert_table_column(
        table_id: str,
        cells: Optional[list[str]] = None,
        index: Optional[int] = None,
        width_in: Optional[float] = None,
    ) -> str:
        """Add a column to a table, copying the look of the column beside it.

        Args:
            table_id: Id of the table block.
            cells: Cell texts top to bottom, one per row from row 0; '' leaves a row empty (e.g. a title row). Omit for an empty column.
            index: Zero-based position for the new column. Omit to add it on the right.
            width_in: Width in inches. Omit to match its neighbour and keep the table's width.
        """
        position = doc.insert_table_column(table_id, index, cells, width_in)
        return f"Inserted column {position} in {table_id}; columns to its right moved up by one."

    @tool
    def merge_table_cells(table_id: str, row: int, col: int, end_row: int, end_col: int) -> str:
        """Merge a rectangle of cells into one, e.g. a heading across columns. All text is kept; the merged cell is addressed by its top-left position.

        Args:
            table_id: Id of the table block.
            row: Zero-based row of the top-left cell.
            col: Zero-based column of the top-left cell.
            end_row: Zero-based row of the bottom-right cell.
            end_col: Zero-based column of the bottom-right cell.
        """
        doc.merge_cells(table_id, row, col, end_row, end_col)
        top, left = min(row, end_row), min(col, end_col)
        return f"Merged into {table_id}[{top},{left}] ({abs(end_row - row) + 1} row(s) by {abs(end_col - col) + 1} column(s))."

    @tool
    def unmerge_table_cells(table_id: str, row: int, col: int) -> str:
        """Split a merged cell back into single cells. Its text stays in the top-left one.

        Args:
            table_id: Id of the table block.
            row: Zero-based row of any position inside the merged cell.
            col: Zero-based column of any position inside the merged cell.
        """
        doc.unmerge_cells(table_id, row, col)
        return f"Split the merged cell at {table_id}[{row},{col}] into single cells."

    @tool
    def delete_table_row(table_id: str, row: int) -> str:
        """Delete one row of a table.

        Args:
            table_id: Id of the table block.
            row: Zero-based row.
        """
        doc.delete_table_row(table_id, row)
        return f"Deleted row {row} of {table_id}"

    @tool
    def delete_table_column(table_id: str, col: int) -> str:
        """Delete one column of a table.

        Args:
            table_id: Id of the table block.
            col: Zero-based column on the table's grid.
        """
        doc.delete_table_column(table_id, col)
        return f"Deleted column {col} of {table_id}; columns to its right moved down by one."

    @tool
    def format_table_cells(
        cells: list[str],
        shading: Optional[str] = None,
        vertical_align: Optional[Literal["top", "center", "bottom"]] = None,
        border: Optional[Literal["single", "double", "dashed", "dotted", "thick", "none"]] = None,
        border_sides: Optional[list[Literal["all", "outer", "inner", "top", "bottom", "left", "right"]]] = None,
        border_color: Optional[str] = None,
        border_width_pt: Optional[float] = None,
        margins_in: Optional[float] = None,
        text_direction: Optional[Literal["horizontal", "vertical", "vertical_up"]] = None,
    ) -> str:
        """Set cell background, vertical alignment, borders, padding and text direction. For the text in cells (bold, color, alignment) use format_text or set_paragraph_format with the same refs.

        Args:
            cells: Cells 'b3[0,1]', rows 'b3[0,*]', columns 'b3[*,1]', or a table id for every cell.
            shading: Background as 6-digit hex (e.g. 'D9E2F3') or a color name; 'none' removes it.
            vertical_align: Where the text sits in the cell.
            border: Line style for border_sides; 'none' removes those lines.
            border_sides: 'all' (default), 'outer' (a box around the cells together), 'inner' (lines between them), or a side of each cell.
            border_color: Line color as 6-digit hex or a color name.
            border_width_pt: Line width in points, 0.25 to 6 (default 0.5).
            margins_in: Padding inside the cell, in inches.
            text_direction: 'vertical' reads top to bottom, 'vertical_up' bottom to top.
        """
        count = sum(
            doc.format_cells(ref, shading, vertical_align, border, border_sides, border_color, border_width_pt,
                             margins_in, text_direction)
            for ref in as_list(cells)
        )
        return f"Formatted {count} cell(s)."

    @tool
    def set_table_format(
        table_id: str,
        column_widths_in: Optional[list[float]] = None,
        header_rows: Optional[int] = None,
        alignment: Optional[Literal["left", "center", "right"]] = None,
        borders: Optional[Literal["all", "none", "outer", "horizontal"]] = None,
        cell_margins_in: Optional[float] = None,
    ) -> str:
        """Set a table's layout: column widths, repeating header rows, position, lines and cell padding. Give at least one.

        Args:
            table_id: Id of the table block.
            column_widths_in: Width of every column in inches, left to right.
            header_rows: Rows at the top that repeat on each page; 0 for none.
            alignment: Position between the margins.
            borders: 'all' (full grid), 'none', 'outer' (a box) or 'horizontal' (rules between rows). Lines on particular cells: format_table_cells.
            cell_margins_in: Padding inside every cell, in inches.
        """
        doc.set_table_format(table_id, column_widths_in, header_rows, alignment, borders, cell_margins_in)
        done = []
        if column_widths_in is not None:
            done.append(f"{len(column_widths_in)} column width(s)")
        if header_rows is not None:
            done.append(f"{header_rows} header row(s)")
        if alignment is not None:
            done.append(f"alignment {alignment}")
        if borders is not None:
            done.append(f"borders {borders}")
        if cell_margins_in is not None:
            done.append(f"cell padding {cell_margins_in} in")
        return f"Set {', '.join(done)} on {table_id}."

    return [
        set_table_cell, insert_table_row, insert_table_column, merge_table_cells, unmerge_table_cells,
        delete_table_row, delete_table_column, format_table_cells, set_table_format,
    ]
