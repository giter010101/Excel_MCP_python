"""Tools: excel_merge_cells, excel_unmerge_cells"""

from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook


def register_merge_cells(mcp: FastMCP):

    @mcp.tool(
        name="excel_merge_cells",
        description="Merge a range of cells in a worksheet",
    )
    def excel_merge_cells(
        fileAbsolutePath: str,
        sheetName: str,
        range: str,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            range: Range of cells to merge (e.g. "A1:B2")
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]
        ws.merge_cells(range)
        save_workbook(wb, fileAbsolutePath)
        return f"Cells {range} merged in sheet '{sheetName}'"

    @mcp.tool(
        name="excel_unmerge_cells",
        description="Unmerge a range of cells in a worksheet",
    )
    def excel_unmerge_cells(
        fileAbsolutePath: str,
        sheetName: str,
        range: str,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            range: Range of cells to unmerge (e.g. "A1:B2")
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]
        ws.unmerge_cells(range)
        save_workbook(wb, fileAbsolutePath)
        return f"Cells {range} unmerged in sheet '{sheetName}'"
