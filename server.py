"""
Excel MCP Server - Python version using FastMCP
Equivalent to the Go version: github.com/negokaz/excel-mcp-server
"""

from fastmcp import FastMCP
from tools.excel_describe_sheets import register_describe_sheets
from tools.excel_read_sheet import register_read_sheet
from tools.excel_write_to_sheet import register_write_to_sheet
from tools.excel_create_workbook import register_create_workbook
from tools.excel_manage_sheets import register_manage_sheets
from tools.excel_manage_rows_cols import register_manage_rows_cols
from tools.excel_format_range import register_format_range
from tools.excel_create_table import register_create_table
from tools.excel_create_chart import register_create_chart
from tools.excel_create_pivot_table import register_create_pivot_table
from tools.excel_copy_sheet import register_copy_sheet
from tools.excel_merge_cells import register_merge_cells
from tools.excel_manage_named_ranges import register_manage_named_ranges

mcp = FastMCP(
    name="excel-mcp-server",
    instructions=(
        "Use this server to read and write Excel files (.xlsx, .xls, .xlsm). "
        "Always provide absolute paths to Excel files. "
        "Use excel_describe_sheets to discover sheet names and ranges before reading/writing."
    ),
)

# Register all tools
register_describe_sheets(mcp)
register_read_sheet(mcp)
register_write_to_sheet(mcp)
register_create_workbook(mcp)
register_manage_sheets(mcp)
register_manage_rows_cols(mcp)
register_format_range(mcp)
register_create_table(mcp)
register_create_chart(mcp)
register_create_pivot_table(mcp)
register_copy_sheet(mcp)
register_merge_cells(mcp)
register_manage_named_ranges(mcp)

if __name__ == "__main__":
    mcp.run()
