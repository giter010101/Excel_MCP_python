"""Tool: excel_create_workbook"""

from fastmcp import FastMCP
from excel_engine import create_workbook


def register_create_workbook(mcp: FastMCP):

    @mcp.tool(
        name="excel_create_workbook",
        description="Create a new Excel workbook",
    )
    def excel_create_workbook(fileAbsolutePath: str) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path where the Excel file should be created
        """
        create_workbook(fileAbsolutePath)
        return f"Workbook created successfully at {fileAbsolutePath}"
