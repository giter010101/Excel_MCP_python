"""Tool: excel_set_dimensions — Set row heights and column widths."""

from typing import Optional
from fastmcp import FastMCP
from openpyxl.utils import get_column_letter
from excel_engine import open_workbook, save_workbook, _escape, format_result


def register_set_dimensions(mcp: FastMCP):

    @mcp.tool(
        name="excel_set_dimensions",
        description="Set row heights and/or column widths in a worksheet. Can also hide/unhide rows and columns.",
    )
    def excel_set_dimensions(
        fileAbsolutePath: str,
        sheetName: str,
        rowHeights: Optional[dict[int, float]] = None,
        columnWidths: Optional[dict[str, float]] = None,
        hiddenRows: Optional[list[int]] = None,
        hiddenColumns: Optional[list[str]] = None,
        unhideRows: Optional[list[int]] = None,
        unhideColumns: Optional[list[str]] = None,
        format: str = "json",
    ) -> str:
        """
        Set row heights and column widths. Can also hide/unhide rows and columns.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            rowHeights: Dict mapping row numbers (1-based) to heights in points.
                        Example: {1: 30, 2: 25, 5: 40}
            columnWidths: Dict mapping column letters to widths in character units.
                          Example: {"A": 20, "B": 15, "D": 30}
            hiddenRows: List of row numbers to hide. Example: [3, 4]
            hiddenColumns: List of column letters to hide. Example: ["C", "E"]
            unhideRows: List of row numbers to unhide. Example: [3]
            unhideColumns: List of column letters to unhide. Example: ["C"]
            format: Output format — "json" (default) or "html"
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        changes: list[str] = []

        # Row heights
        if rowHeights:
            for row_num, height in rowHeights.items():
                ws.row_dimensions[int(row_num)].height = float(height)
                changes.append(f"Row {row_num} height → {height}pt")

        # Column widths
        if columnWidths:
            for col_letter, width in columnWidths.items():
                ws.column_dimensions[col_letter.upper()].width = float(width)
                changes.append(f"Column {col_letter.upper()} width → {width}")

        # Hide rows
        if hiddenRows:
            for row_num in hiddenRows:
                ws.row_dimensions[int(row_num)].hidden = True
                changes.append(f"Row {row_num} hidden")

        # Hide columns
        if hiddenColumns:
            for col_letter in hiddenColumns:
                ws.column_dimensions[col_letter.upper()].hidden = True
                changes.append(f"Column {col_letter.upper()} hidden")

        # Unhide rows
        if unhideRows:
            for row_num in unhideRows:
                ws.row_dimensions[int(row_num)].hidden = False
                changes.append(f"Row {row_num} unhidden")

        # Unhide columns
        if unhideColumns:
            for col_letter in unhideColumns:
                ws.column_dimensions[col_letter.upper()].hidden = False
                changes.append(f"Column {col_letter.upper()} unhidden")

        if not changes:
            return "No dimension changes requested."

        save_workbook(wb, fileAbsolutePath)

        return format_result(
            action="Set Dimensions",
            message=f"Applied {len(changes)} dimension change(s) in sheet '{sheetName}'.",
            metadata={
                "backend": "openpyxl",
                "sheetName": sheetName,
                "changes": changes,
            },
            fmt=format,
        )
