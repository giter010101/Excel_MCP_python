"""Tool: excel_read_sheet"""

from typing import Optional
from fastmcp import FastMCP
from excel_engine import open_workbook, read_sheet_html


def register_read_sheet(mcp: FastMCP):

    @mcp.tool(
        name="excel_read_sheet",
        description="Read values from Excel sheet with pagination.",
    )
    def excel_read_sheet(
        fileAbsolutePath: str,
        sheetName: str,
        range: Optional[str] = None,
        showFormula: bool = False,
        showStyle: bool = False,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            range: Range of cells to read (e.g. "A1:C10"). Defaults to first paging range.
            showFormula: Show formula instead of value
            showStyle: Show style information for cells (not yet implemented in Python version)
        """
        wb = open_workbook(fileAbsolutePath)
        result = read_sheet_html(wb, sheetName, range, showFormula, showStyle)
        wb.close()
        return result["html"]
