"""Tool: excel_get_validation_info — Read existing data validation rules from a worksheet."""

import json
from fastmcp import FastMCP
from excel_engine import open_workbook, _escape


def register_get_validation_info(mcp: FastMCP):

    @mcp.tool(
        name="excel_get_validation_info",
        description="Read all data validation rules from a worksheet. Returns each rule's type, operator, formulas, target ranges, and messages.",
    )
    def excel_get_validation_info(
        fileAbsolutePath: str,
        sheetName: str,
    ) -> str:
        """
        Get all data validation rules configured in a worksheet.

        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Sheet name in the Excel file
        """
        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        # Access data validations
        validations = ws.data_validations
        if validations is None or len(validations.dataValidation) == 0:
            wb.close()
            return f"No data validation rules found in sheet '{sheetName}'"

        rules: list[dict] = []
        for dv in validations.dataValidation:
            rule: dict = {
                "type": dv.type,
                "ranges": str(dv.sqref) if dv.sqref else "",
            }
            if dv.operator:
                rule["operator"] = dv.operator
            if dv.formula1 is not None:
                rule["formula1"] = str(dv.formula1)
            if dv.formula2 is not None:
                rule["formula2"] = str(dv.formula2)
            rule["allowBlank"] = dv.allow_blank

            # Error message info
            if dv.showErrorMessage:
                err_info: dict = {}
                if dv.errorTitle:
                    err_info["title"] = dv.errorTitle
                if dv.error:
                    err_info["message"] = dv.error
                if dv.errorStyle:
                    err_info["style"] = dv.errorStyle
                if err_info:
                    rule["errorMessage"] = err_info

            # Prompt / input message info
            if dv.showInputMessage:
                prompt_info: dict = {}
                if dv.promptTitle:
                    prompt_info["title"] = dv.promptTitle
                if dv.prompt:
                    prompt_info["message"] = dv.prompt
                if prompt_info:
                    rule["inputMessage"] = prompt_info

            rules.append(rule)

        result = {
            "sheetName": sheetName,
            "validationCount": len(rules),
            "rules": rules,
        }
        wb.close()
        return json.dumps(result, indent=2, ensure_ascii=False, default=str)
