"""Tool: excel_data_validation (BETA)"""

from typing import Optional
from fastmcp import FastMCP
from openpyxl.worksheet.datavalidation import DataValidation
from excel_engine import open_workbook, save_workbook, _escape, format_result


def register_data_validation(mcp: FastMCP):

    @mcp.tool(
        name="excel_data_validation",
        description="Add data validation rules to cells (dropdown lists, number constraints, date ranges, etc.). (BETA)",
    )
    def excel_data_validation(
        fileAbsolutePath: str,
        sheetName: str,
        range: str,
        validationType: str,
        formula1: Optional[str] = None,
        formula2: Optional[str] = None,
        operator: Optional[str] = None,
        allowBlank: bool = True,
        showErrorMessage: bool = True,
        errorTitle: Optional[str] = None,
        errorMessage: Optional[str] = None,
        errorStyle: Optional[str] = None,
        showInputMessage: bool = False,
        promptTitle: Optional[str] = None,
        promptMessage: Optional[str] = None,
        format: str = "json",
    ) -> str:
        """
        Add a data validation rule to a range of cells.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
            range: Range to apply validation to (e.g. "B2:B100")
            validationType: Type of validation. One of:
                - "list"       : dropdown list. Set formula1 to comma-separated values
                                 e.g. '"Yes,No,Maybe"' (with quotes) or a range like "Sheet1!$A$1:$A$5"
                - "whole"      : whole number constraint
                - "decimal"    : decimal number constraint
                - "date"       : date constraint
                - "time"       : time constraint
                - "textLength" : text length constraint
                - "custom"     : custom formula validation. Set formula1 to the formula.
            formula1: First value/formula for the validation.
                For "list": comma-separated values in quotes like '"A,B,C"' or a cell range.
                For numeric/date: the first bound value.
                For "custom": an Excel formula like "=AND(A1>0,A1<100)".
            formula2: Second value (used with "between"/"notBetween" operators).
            operator: Comparison operator (not used for "list"/"custom"). One of:
                "between", "notBetween", "equal", "notEqual",
                "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual"
            allowBlank: Allow blank cells. Default True.
            showErrorMessage: Show error popup on invalid input. Default True.
            errorTitle: Title of the error popup.
            errorMessage: Text of the error popup.
            errorStyle: Error style: "stop" (block), "warning", or "information". Default "stop".
            showInputMessage: Show a hint message when the cell is selected.
            promptTitle: Title of the input hint.
            promptMessage: Text of the input hint.
            format: Output format — "json" (default) or "html"
        """
        VALID_TYPES = {"list", "whole", "decimal", "date", "time", "textLength", "custom"}
        if validationType not in VALID_TYPES:
            raise ValueError(
                f"Invalid validationType: {validationType!r}. Must be one of: {', '.join(sorted(VALID_TYPES))}"
            )

        VALID_OPERATORS = {
            "between", "notBetween", "equal", "notEqual",
            "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual",
        }
        if operator and operator not in VALID_OPERATORS:
            raise ValueError(
                f"Invalid operator: {operator!r}. Must be one of: {', '.join(sorted(VALID_OPERATORS))}"
            )

        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        # Build DataValidation
        dv_kwargs = {
            "type": validationType,
            "allow_blank": allowBlank,
        }
        if operator:
            dv_kwargs["operator"] = operator
        if formula1 is not None:
            dv_kwargs["formula1"] = formula1
        if formula2 is not None:
            dv_kwargs["formula2"] = formula2

        dv = DataValidation(**dv_kwargs)

        # Error message
        dv.showErrorMessage = showErrorMessage
        if errorTitle:
            dv.errorTitle = errorTitle
        if errorMessage:
            dv.error = errorMessage
        if errorStyle:
            dv.errorStyle = errorStyle

        # Input/prompt message
        dv.showInputMessage = showInputMessage
        if promptTitle:
            dv.promptTitle = promptTitle
        if promptMessage:
            dv.prompt = promptMessage

        ws.add_data_validation(dv)
        dv.add(range)

        save_workbook(wb, fileAbsolutePath)

        meta: dict = {
            "backend": "openpyxl",
            "sheetName": sheetName,
            "validatedRange": range,
            "type": validationType,
            "allowBlank": allowBlank,
        }
        if operator:
            meta["operator"] = operator
        if formula1 is not None:
            meta["formula1"] = str(formula1)
        if formula2 is not None:
            meta["formula2"] = str(formula2)

        return format_result(
            action="Data Validation",
            message=f"Validation rule ({validationType}) applied to range {range} in sheet '{sheetName}'.",
            metadata=meta,
            fmt=format,
        )
