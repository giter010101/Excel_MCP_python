"""Tool: excel_manage_named_ranges"""

import json
from typing import Optional
from fastmcp import FastMCP
from openpyxl.workbook.defined_name import DefinedName
from excel_engine import open_workbook, save_workbook


def register_manage_named_ranges(mcp: FastMCP):

    @mcp.tool(
        name="excel_manage_named_ranges",
        description="Manage named ranges (defined names) in the Excel file. Supports listing, creating, and deleting named ranges.",
    )
    def excel_manage_named_ranges(
        fileAbsolutePath: str,
        action: str,
        name: Optional[str] = None,
        refersTo: Optional[str] = None,
        scope: Optional[str] = None,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            action: Action to perform: "list", "create", or "delete"
            name: Name of the named range (required for create and delete)
            refersTo: Cell or range reference (e.g. "Sheet1!$A$1:$D$10"). Required for create.
            scope: Scope (sheet name for sheet-level, empty for workbook-level)
        """
        wb = open_workbook(fileAbsolutePath)

        if action == "list":
            names = []
            for dn in wb.defined_names.values():
                names.append({
                    "name": dn.name,
                    "refersTo": dn.attr_text,
                    "scope": dn.localSheetId,
                })
            return json.dumps(names, indent=2, ensure_ascii=False)

        elif action == "create":
            if not name:
                raise ValueError("'name' is required for create action")
            if not refersTo:
                raise ValueError("'refersTo' is required for create action")
            dn = DefinedName(name=name, attr_text=refersTo)
            wb.defined_names[name] = dn
            save_workbook(wb, fileAbsolutePath)
            return f"Named range '{name}' created pointing to '{refersTo}'"

        elif action == "delete":
            if not name:
                raise ValueError("'name' is required for delete action")
            if name not in wb.defined_names:
                raise ValueError(f"Named range not found: {name!r}")
            del wb.defined_names[name]
            save_workbook(wb, fileAbsolutePath)
            return f"Named range '{name}' deleted"

        else:
            raise ValueError(f"Unknown action: {action!r}. Must be 'list', 'create', or 'delete'")
