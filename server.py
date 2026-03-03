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

from tools.excel_copy_sheet import register_copy_sheet
from tools.excel_merge_cells import register_merge_cells
from tools.excel_manage_named_ranges import register_manage_named_ranges

from tools.excel_auto_filter import register_auto_filter
from tools.excel_add_comment import register_add_comment
from tools.excel_data_validation import register_data_validation
from tools.excel_conditional_formatting import register_conditional_formatting
from tools.excel_validate_formula import register_validate_formula
from tools.excel_copy_range import register_copy_range
from tools.excel_delete_range import register_delete_range
from tools.excel_move_range import register_move_range
from tools.excel_get_validation_info import register_get_validation_info
from tools.excel_get_merged_cells import register_get_merged_cells
from tools.excel_set_dimensions import register_set_dimensions


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

register_copy_sheet(mcp)
register_merge_cells(mcp)
register_manage_named_ranges(mcp)

# Beta tools

register_auto_filter(mcp)
register_add_comment(mcp)
register_data_validation(mcp)
register_conditional_formatting(mcp)
register_validate_formula(mcp)
register_copy_range(mcp)
register_delete_range(mcp)
register_move_range(mcp)
register_get_validation_info(mcp)
register_get_merged_cells(mcp)
register_set_dimensions(mcp)


if __name__ == "__main__":
    mcp.run()
