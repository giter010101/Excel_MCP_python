"""Tool: excel_delete_range — Delete a range and shift remaining cells."""

from fastmcp import FastMCP
from openpyxl.utils import column_index_from_string
from excel_engine import open_workbook, save_workbook, parse_range, _escape


def register_delete_range(mcp: FastMCP):

    @mcp.tool(
        name="excel_delete_range",
        description="Delete a range of cells and shift remaining cells up or left.",
    )
    def excel_delete_range(
        fileAbsolutePath: str,
        sheetName: str,
        range: str,
        shiftDirection: str = "up",
    ) -> str:
        """
        Delete a range of cells and shift remaining cells.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            range: Range to delete (e.g. "B2:D5")
            shiftDirection: Direction to shift remaining cells after deletion.
                            "up" — delete rows covered by the range and shift cells up.
                            "left" — delete columns covered by the range and shift cells left.
        """
        if shiftDirection not in ("up", "left"):
            raise ValueError(
                f"Invalid shiftDirection: {shiftDirection!r}. Must be 'up' or 'left'."
            )

        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        min_col, min_row, max_col, max_row = parse_range(range)

        if shiftDirection == "up":
            row_count = max_row - min_row + 1
            ws.delete_rows(min_row, amount=row_count)
            detail = f"{row_count} row(s) deleted, cells shifted up"
        else:  # left
            col_count = max_col - min_col + 1
            ws.delete_cols(min_col, amount=col_count)
            detail = f"{col_count} column(s) deleted, cells shifted left"

        save_workbook(wb, fileAbsolutePath)

        html = "<h2>Delete Range</h2>\n"
        html += f"<p>Range {range} deleted in sheet '{_escape(sheetName)}'.</p>\n"
        html += f"<p>{detail}</p>\n"
        html += "<h2>Metadata</h2>\n<ul>\n"
        html += "<li>backend: openpyxl</li>\n"
        html += f"<li>sheet name: {_escape(sheetName)}</li>\n"
        html += f"<li>shift direction: {shiftDirection}</li>\n"
        html += "</ul>\n"
        return html
