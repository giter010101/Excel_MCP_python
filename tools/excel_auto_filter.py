"""Tool: excel_auto_filter (BETA)"""

from typing import Optional
from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook, get_used_range, _escape


def register_auto_filter(mcp: FastMCP):

    @mcp.tool(
        name="excel_auto_filter",
        description="Add or remove auto-filters (dropdown arrows on column headers) on a worksheet. (BETA)",
    )
    def excel_auto_filter(
        fileAbsolutePath: str,
        sheetName: str,
        range: Optional[str] = None,
        remove: bool = False,
    ) -> str:
        """
        Add auto-filter dropdown arrows to column headers in the given range,
        or remove existing auto-filters.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            range: Range to apply auto-filter to (e.g. "A1:F100").
                   Defaults to the used range of the sheet.
            remove: If True, removes any existing auto-filter from the sheet.
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        if remove:
            ws.auto_filter.ref = None
            save_workbook(wb, fileAbsolutePath)
            html = "<h2>Auto Filter</h2>\n"
            html += f"<p>Auto-filter removed from sheet '{_escape(sheetName)}'.</p>\n"
            return html

        filter_range = range
        if not filter_range:
            filter_range = get_used_range(ws)
        if not filter_range:
            raise ValueError("No range specified and sheet has no data")

        ws.auto_filter.ref = filter_range

        save_workbook(wb, fileAbsolutePath)

        html = "<h2>Auto Filter</h2>\n"
        html += f"<p>Auto-filter applied to range {filter_range} in sheet '{_escape(sheetName)}'.</p>\n"
        html += "<h2>Metadata</h2>\n<ul>\n"
        html += "<li>backend: openpyxl</li>\n"
        html += f"<li>sheet name: {_escape(sheetName)}</li>\n"
        html += f"<li>filter range: {filter_range}</li>\n"
        html += "</ul>\n"
        return html
