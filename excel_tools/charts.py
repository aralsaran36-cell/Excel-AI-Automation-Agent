from excel_tools.workbook_manager import ExcelToolError, get_sheet, safe_tool, validate_address

CHART_TYPES = {
    "column": "column_clustered",
    "bar": "bar_clustered",
    "line": "line",
    "pie": "pie",
    "area": "area",
    "scatter": "xy_scatter",
}


@safe_tool
def create_excel_chart(data_range: str, chart_type: str = "column", title: str | None = None,
                       position_cell: str = "H2", width: int = 480, height: int = 300,
                       workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """Create a chart from a range (include header row). chart_type: column, bar, line, pie, area, scatter."""
    if chart_type not in CHART_TYPES:
        raise ExcelToolError(f"Invalid chart_type. Use one of: {list(CHART_TYPES)}")
    if not (100 <= width <= 2000 and 100 <= height <= 2000):
        raise ExcelToolError("width/height 100 mudhal 2000 varaikkum.")
    source = validate_address(data_range)
    anchor = validate_address(position_cell)
    if ":" in anchor:
        raise ExcelToolError("position_cell oru single cell (example: H2).")
    sheet = get_sheet(workbook_name, sheet_name)
    anchor_cell = sheet.range(anchor)
    chart = sheet.charts.add(left=anchor_cell.left, top=anchor_cell.top, width=width, height=height)
    chart.set_source_data(sheet.range(source))
    chart.chart_type = CHART_TYPES[chart_type]
    if title:
        chart.name = title
    return {"success": True, "chart_type": chart_type, "data_range": source,
            "position": anchor, "name": chart.name}