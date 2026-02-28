"""Tools: excel_rename_sheet, excel_delete_sheet"""

from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook


def register_manage_sheets(mcp: FastMCP):

    @mcp.tool(
        name="excel_rename_sheet",
        description="Rename a worksheet in the Excel workbook",
    )
    def excel_rename_sheet(
        fileAbsolutePath: str,
        oldName: str,
        newName: str,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            oldName: Current name of the worksheet
            newName: New name for the worksheet
        """
        wb = open_workbook(fileAbsolutePath)
        if oldName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {oldName!r}")
        ws = wb[oldName]
        ws.title = newName
        save_workbook(wb, fileAbsolutePath)
        return f"Sheet '{oldName}' renamed to '{newName}' successfully"

    @mcp.tool(
        name="excel_delete_sheet",
        description="Delete a worksheet from the Excel workbook",
    )
    def excel_delete_sheet(
        fileAbsolutePath: str,
        sheetName: str,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet to delete
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        del wb[sheetName]
        save_workbook(wb, fileAbsolutePath)
        return f"Sheet '{sheetName}' deleted successfully"
