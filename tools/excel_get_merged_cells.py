"""Tool: excel_get_merged_cells — List all merged cell ranges in a worksheet."""

import json
from fastmcp import FastMCP
from excel_engine import open_workbook


def register_get_merged_cells(mcp: FastMCP):

    @mcp.tool(
        name="excel_get_merged_cells",
        description="List all merged cell ranges in a worksheet.",
    )
    def excel_get_merged_cells(
        fileAbsolutePath: str,
        sheetName: str,
    ) -> str:
        """
        Get all merged cell ranges in a worksheet.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        # ws.merged_cells.ranges is the preferred API (merged_cell_ranges is deprecated)
        merged = []
        for cell_range in ws.merged_cells.ranges:
            merged.append(str(cell_range))

        if not merged:
            wb.close()
            return f"No merged cells found in sheet '{sheetName}'"

        result = {
            "sheetName": sheetName,
            "mergedRangeCount": len(merged),
            "mergedRanges": merged,
        }
        wb.close()
        return json.dumps(result, indent=2, ensure_ascii=False)
