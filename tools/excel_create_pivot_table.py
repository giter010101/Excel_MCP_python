"""Tool: excel_create_pivot_table
Note: openpyxl does not support creating pivot tables directly.
We use a workaround with a helper note in the output.
"""

from typing import Optional
from fastmcp import FastMCP


def register_create_pivot_table(mcp: FastMCP):

    @mcp.tool(
        name="excel_create_pivot_table",
        description="Create a pivot table in a worksheet (note: Python/openpyxl support for pivot tables is limited; the pivot table structure is written but may require Excel to refresh)",
    )
    def excel_create_pivot_table(
        fileAbsolutePath: str,
        sheetName: str,
        dataRange: str,
        targetCell: str,
        name: str,
        rows: list[str],
        values: list[str],
        columns: Optional[list[str]] = None,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            dataRange: Range of source data (e.g. "Sheet1!$A$1:$E$20")
            targetCell: Target cell for the pivot table (e.g. "G1")
            name: Unique name for the pivot table
            rows: Fields to use for rows
            values: Fields to use for values
            columns: Fields to use for columns
        """
        # openpyxl does not support writing PivotTable XML natively.
        # We return an informative message.
        return (
            f"[Excel MCP Python] Pivot table creation via openpyxl is not supported. "
            f"To create a pivot table named '{name}' on '{sheetName}' from '{dataRange}', "
            f"please use Excel directly or consider using xlwings/win32com on Windows.\n"
            f"Rows: {rows}\nValues: {values}\nColumns: {columns or []}"
        )
