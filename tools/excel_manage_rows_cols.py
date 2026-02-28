"""Tools: excel_insert_rows, excel_delete_rows, excel_insert_columns, excel_delete_columns"""

from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook


def register_manage_rows_cols(mcp: FastMCP):

    @mcp.tool(
        name="excel_insert_rows",
        description="Insert rows into a worksheet",
    )
    def excel_insert_rows(
        fileAbsolutePath: str,
        sheetName: str,
        startRow: int,
        count: int = 1,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            startRow: Row number where to start inserting (1-based)
            count: Number of rows to insert
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]
        ws.insert_rows(startRow, amount=count)
        save_workbook(wb, fileAbsolutePath)
        return f"{count} rows inserted starting at row {startRow}"

    @mcp.tool(
        name="excel_delete_rows",
        description="Delete rows from a worksheet",
    )
    def excel_delete_rows(
        fileAbsolutePath: str,
        sheetName: str,
        startRow: int,
        count: int = 1,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            startRow: Row number where to start deleting (1-based)
            count: Number of rows to delete
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]
        ws.delete_rows(startRow, amount=count)
        save_workbook(wb, fileAbsolutePath)
        return f"{count} rows deleted starting at row {startRow}"

    @mcp.tool(
        name="excel_insert_columns",
        description="Insert columns into a worksheet",
    )
    def excel_insert_columns(
        fileAbsolutePath: str,
        sheetName: str,
        startCol: int,
        count: int = 1,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            startCol: Column number where to start inserting (1-based)
            count: Number of columns to insert
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]
        ws.insert_cols(startCol, amount=count)
        save_workbook(wb, fileAbsolutePath)
        return f"{count} columns inserted starting at column {startCol}"

    @mcp.tool(
        name="excel_delete_columns",
        description="Delete columns from a worksheet",
    )
    def excel_delete_columns(
        fileAbsolutePath: str,
        sheetName: str,
        startCol: int,
        count: int = 1,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            startCol: Column number where to start deleting (1-based)
            count: Number of columns to delete
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]
        ws.delete_cols(startCol, amount=count)
        save_workbook(wb, fileAbsolutePath)
        return f"{count} columns deleted starting at column {startCol}"
