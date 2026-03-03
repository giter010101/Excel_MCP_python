"""Tool: excel_move_range — Move a range of cells within a worksheet."""

from typing import Optional
from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook, _escape


def register_move_range(mcp: FastMCP):

    @mcp.tool(
        name="excel_move_range",
        description="Move a range of cells by a given number of rows and/or columns. The source range is cleared and the destination is overwritten. Optionally translates formulas.",
    )
    def excel_move_range(
        fileAbsolutePath: str,
        sheetName: str,
        range: str,
        rows: int = 0,
        cols: int = 0,
        translateFormulas: bool = True,
    ) -> str:
        """
        Move a range of cells within the same worksheet.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            range: Range to move (e.g. "D4:F10")
            rows: Number of rows to shift (negative = up, positive = down)
            cols: Number of columns to shift (negative = left, positive = right)
            translateFormulas: If True, formulas are adjusted to the new positions.
        """
        if rows == 0 and cols == 0:
            return "No movement requested (rows=0, cols=0)."

        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        ws.move_range(range, rows=rows, cols=cols, translate=translateFormulas)

        save_workbook(wb, fileAbsolutePath)

        direction_parts = []
        if rows > 0:
            direction_parts.append(f"{rows} row(s) down")
        elif rows < 0:
            direction_parts.append(f"{abs(rows)} row(s) up")
        if cols > 0:
            direction_parts.append(f"{cols} column(s) right")
        elif cols < 0:
            direction_parts.append(f"{abs(cols)} column(s) left")
        direction = ", ".join(direction_parts)

        html = "<h2>Move Range</h2>\n"
        html += f"<p>Moved range {range} in sheet '{_escape(sheetName)}': {direction}.</p>\n"
        html += "<h2>Metadata</h2>\n<ul>\n"
        html += "<li>backend: openpyxl</li>\n"
        html += f"<li>sheet name: {_escape(sheetName)}</li>\n"
        html += f"<li>rows shifted: {rows}</li>\n"
        html += f"<li>cols shifted: {cols}</li>\n"
        html += f"<li>formulas translated: {translateFormulas}</li>\n"
        html += "</ul>\n"
        return html
