"""Tool: excel_create_chart"""

from typing import Optional
from fastmcp import FastMCP
from openpyxl.chart import (
    BarChart, LineChart, PieChart, ScatterChart, AreaChart,
    Reference,
)
from excel_engine import open_workbook, save_workbook, parse_range


CHART_TYPES = {
    "bar": BarChart,
    "line": LineChart,
    "pie": PieChart,
    "scatter": ScatterChart,
    "area": AreaChart,
}


def register_create_chart(mcp: FastMCP):

    @mcp.tool(
        name="excel_create_chart",
        description="Create a chart in a worksheet",
    )
    def excel_create_chart(
        fileAbsolutePath: str,
        sheetName: str,
        range: str,
        chartType: str,
        title: Optional[str] = None,
    ) -> str:
        """
        Args:
            fileAbsolutePath: Absolute path to the Excel file
            sheetName: Name of the worksheet
            range: Range of data for the chart (e.g. "Sheet1!A1:B10")
            chartType: Type of chart: line, bar, pie, scatter, area
            title: Title of the chart
        """
        chart_cls = CHART_TYPES.get(chartType.lower())
        if chart_cls is None:
            raise ValueError(
                f"Unknown chart type: {chartType!r}. "
                f"Supported: {', '.join(CHART_TYPES)}"
            )

        wb = open_workbook(fileAbsolutePath)
        if sheetName not in wb.sheetnames:
            raise ValueError(f"Sheet not found: {sheetName!r}")
        ws = wb[sheetName]

        # Parse data range (may include sheet name prefix like "Sheet1!A1:B10")
        if "!" in range:
            data_sheet_name, data_range = range.split("!", 1)
            data_sheet_name = data_sheet_name.strip("'\"")
        else:
            data_sheet_name = sheetName
            data_range = range

        if data_sheet_name not in wb.sheetnames:
            raise ValueError(f"Data sheet not found: {data_sheet_name!r}")
        data_ws = wb[data_sheet_name]

        min_col, min_row, max_col, max_row = parse_range(data_range)

        chart = chart_cls()
        if title:
            chart.title = title

        data_ref = Reference(
            data_ws,
            min_col=min_col,
            min_row=min_row,
            max_col=max_col,
            max_row=max_row,
        )

        if chartType.lower() == "scatter":
            from openpyxl.chart import Series
            chart.append(
                Series(
                    Reference(data_ws, min_col=max_col, min_row=min_row, max_row=max_row),
                    xvalues=Reference(data_ws, min_col=min_col, min_row=min_row, max_row=max_row),
                )
            )
        else:
            chart.add_data(data_ref, titles_from_data=True)

        # Anchor chart just below data
        anchor_cell = f"A{max_row + 2}"
        ws.add_chart(chart, anchor_cell)

        save_workbook(wb, fileAbsolutePath)
        return f"{chartType} chart created from range {range} in sheet '{sheetName}'"
