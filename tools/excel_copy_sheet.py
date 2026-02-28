"""Tool: excel_copy_sheet"""

from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook


def register_copy_sheet(mcp: FastMCP):

    @mcp.tool(
        name="excel_copy_sheet",
        description="Copy existing sheet to a new sheet",
    )
    def excel_copy_sheet(
        fileAbsolutePath: str,
        srcSheetName: str,
        dstSheetName: str,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            srcSheetName: Source sheet name in the Excel file
            dstSheetName: Sheet name to be copied (destination)
        """
        wb = open_workbook(fileAbsolutePath)
        if srcSheetName not in wb.sheetnames:
            raise ValueError(f"Source sheet not found: {srcSheetName!r}")
        if dstSheetName in wb.sheetnames:
            raise ValueError(f"Destination sheet already exists: {dstSheetName!r}")

        src = wb[srcSheetName]
        dst = wb.copy_worksheet(src)
        dst.title = dstSheetName

        save_workbook(wb, fileAbsolutePath)
        return f"Sheet '{srcSheetName}' copied to '{dstSheetName}' successfully"
