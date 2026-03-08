"""Tool: excel_add_comment (BETA)"""

from typing import Optional
from fastmcp import FastMCP
from openpyxl.comments import Comment
from excel_engine import open_workbook, save_workbook, _escape, format_result


def register_add_comment(mcp: FastMCP):

    @mcp.tool(
        name="excel_add_comment",
        description="Add, update, or remove comments (notes) on cells in a worksheet. (BETA)",
    )
    def excel_add_comment(
        fileAbsolutePath: str,
        sheetName: str,
        comments: list[dict],
        format: str = "json",
    ) -> str:
        """
        Add, update, or remove comments on one or more cells.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            comments: List of comment objects. Each object has:
                - cell (str, required): Cell reference, e.g. "A1"
                - text (str|null): Comment text. Set to null to REMOVE the comment.
                - author (str, optional): Author name. Defaults to "MCP Server".
                - width (int, optional): Comment box width in pixels. Default 300.
                - height (int, optional): Comment box height in pixels. Default 100.

                Example:
                [
                    {"cell": "A1", "text": "This is a note", "author": "Nathan"},
                    {"cell": "B2", "text": "Check this value"},
                    {"cell": "C3", "text": null}
                ]
            format: Output format — "json" (default) or "html"
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        added = 0
        removed = 0

        for entry in comments:
            cell_ref = entry.get("cell")
            if not cell_ref:
                raise ValueError("Each comment entry must have a 'cell' key")

            text = entry.get("text")
            cell = ws[cell_ref]

            if text is None:
                # Remove comment
                cell.comment = None
                removed += 1
            else:
                author = entry.get("author", "MCP Server")
                comment = Comment(text, author)
                comment.width = entry.get("width", 300)
                comment.height = entry.get("height", 100)
                cell.comment = comment
                added += 1

        save_workbook(wb, fileAbsolutePath)

        return format_result(
            action="Add Comment",
            message=f"Added/updated {added} comment(s), removed {removed} comment(s) in sheet '{sheetName}'.",
            metadata={
                "backend": "openpyxl",
                "sheetName": sheetName,
                "commentsAdded": added,
                "commentsRemoved": removed,
            },
            fmt=format,
        )
