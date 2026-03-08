"""Tool: excel_read_sheet"""

import json
from typing import Optional
from fastmcp import FastMCP
from excel_engine import open_workbook, read_sheet_html, read_sheet_json


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
        format: str = "json",
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            range: Range of cells to read (e.g. "A1:C10"). Defaults to first paging range.
            showFormula: Show formula instead of value
            showStyle: Show style information for cells (not yet implemented in Python version)
            format: Output format — "json" (default, compact) or "html" (legacy verbose)
        """
        wb = open_workbook(fileAbsolutePath)
        if format == "html":
            result = read_sheet_html(wb, sheetName, range, showFormula, showStyle)
            wb.close()
            return result["html"]
        else:
            result = read_sheet_json(wb, sheetName, range, showFormula, showStyle)
            wb.close()
            return json.dumps(result, ensure_ascii=False, default=str)
