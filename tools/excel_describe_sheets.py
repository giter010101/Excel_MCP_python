"""Tool: excel_describe_sheets"""

import json
from fastmcp import FastMCP
from excel_engine import open_workbook, get_used_range, get_paging_ranges


def register_describe_sheets(mcp: FastMCP):

    @mcp.tool(
        name="excel_describe_sheets",
        description="List all sheet information of specified Excel file",
    )
    def excel_describe_sheets(fileAbsolutePath: str) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
        """
        wb = open_workbook(fileAbsolutePath)

        sheets = []
        for name in wb.sheetnames:
            ws = wb[name]
            used_range = get_used_range(ws) or ""
            paging_ranges = get_paging_ranges(ws)

            # Tables
            tables = []
            if hasattr(ws, "tables"):
                for t in ws.tables.values():
                    tables.append({"name": t.displayName, "range": t.ref})

            sheets.append(
                {
                    "name": name,
                    "usedRange": used_range,
                    "tables": tables,
                    "pivotTables": [],  # openpyxl does not expose pivot tables
                    "pagingRanges": paging_ranges,
                }
            )

        response = {"backend": "openpyxl", "sheets": sheets}
        wb.close()
        return json.dumps(response, indent=2, ensure_ascii=False)
