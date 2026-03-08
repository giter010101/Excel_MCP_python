"""Tool: excel_move_range — Move a range of cells within a worksheet."""

from typing import Optional
from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook, _escape, format_result


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
        format: str = "json",
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
            format: Output format — "json" (default) or "html"
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

        return format_result(
            action="Move Range",
            message=f"Moved range {range} in sheet '{sheetName}': {direction}.",
            metadata={
                "backend": "openpyxl",
                "sheetName": sheetName,
                "rowsShifted": rows,
                "colsShifted": cols,
                "formulasTranslated": translateFormulas,
            },
            fmt=format,
        )
