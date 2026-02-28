"""Tool: excel_create_table"""

from typing import Optional
from fastmcp import FastMCP
from openpyxl.worksheet.table import Table, TableStyleInfo
from excel_engine import open_workbook, save_workbook


def register_create_table(mcp: FastMCP):

    @mcp.tool(
        name="excel_create_table",
        description="Create a table in the Excel sheet",
    )
    def excel_create_table(
        fileAbsolutePath: str,
        sheetName: str,
        tableName: str,
        range: Optional[str] = None,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name where the table is created
            tableName: Table name to be created
            range: Range to be a table (e.g. "A1:C10"). Defaults to used range.
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        ref = range or ws.dimensions
        if not ref:
            raise ValueError("No range specified and sheet has no data")

        table = Table(displayName=tableName, ref=ref)
        style = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws.add_table(table)

        save_workbook(wb, fileAbsolutePath)
        return f"Table '{tableName}' created on range {ref} in sheet '{sheetName}'"
