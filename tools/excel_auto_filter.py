"""Tool: excel_auto_filter (BETA)"""

from typing import Optional
from fastmcp import FastMCP
from excel_engine import open_workbook, save_workbook, get_used_range, parse_range, _escape, format_result


def _ranges_overlap(range1: str, range2: str) -> bool:
    """Check if two Excel ranges overlap."""
    try:
        c1_min, r1_min, c1_max, r1_max = parse_range(range1)
        c2_min, r2_min, c2_max, r2_max = parse_range(range2)
        # Two rectangles overlap if they overlap on both axes
        return (c1_min <= c2_max and c2_min <= c1_max and
                r1_min <= r2_max and r2_min <= r1_max)
    except Exception:
        return False

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
        format: str = "json",
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
            format: Output format — "json" (default) or "html"
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        if remove:
            ws.auto_filter.ref = None
            save_workbook(wb, fileAbsolutePath)
            return format_result(
                action="Auto Filter",
                message=f"Auto-filter removed from sheet '{sheetName}'.",
                metadata={},
                fmt=format,
            )

        filter_range = range
        if not filter_range:
            filter_range = get_used_range(ws)
        if not filter_range:
            raise ValueError("No range specified and sheet has no data")

        # Check if the range overlaps with an existing Table.
        # Excel Tables already include built-in auto-filter;
        # adding a separate ws.auto_filter on the same range corrupts the XML.
        for table in ws.tables.values():
            if table.ref and _ranges_overlap(filter_range, table.ref):
                raise ValueError(
                    f"Cannot add auto-filter: range '{filter_range}' overlaps with "
                    f"Table '{table.displayName}' ({table.ref}). "
                    f"Excel Tables already include built-in auto-filters."
                )

        ws.auto_filter.ref = filter_range

        save_workbook(wb, fileAbsolutePath)

        return format_result(
            action="Auto Filter",
            message=f"Auto-filter applied to range {filter_range} in sheet '{sheetName}'.",
            metadata={
                "backend": "openpyxl",
                "sheetName": sheetName,
                "filterRange": filter_range,
            },
            fmt=format,
        )
