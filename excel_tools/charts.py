import sys

from excel_tools.workbook_manager import ExcelToolError, get_sheet, safe_tool, validate_address

IS_WINDOWS = sys.platform.startswith("win")

# One shared type map (works on Windows and Mac via xlwings chart_type strings)
CHART_TYPES = {
    "column": "column_clustered",
    "bar": "bar_clustered",
    "line": "line",
    "pie": "pie",
    "area": "area",
    "scatter": "xy_scatter",
    "doughnut": "doughnut",
    "stacked": "column_stacked",
}
CHART_TYPE_ALIASES = {"donut": "doughnut", "xy_scatter": "scatter", "column_clustered": "column",
                      "column_stacked": "stacked"}


def _need_com(feature: str) -> None:
    """Raise a clear error when a COM-only feature is used on a non-Windows platform."""
    if not IS_WINDOWS:
        raise ExcelToolError(f"{feature} is supported only with Windows Excel (COM).")


def _com_chart(chart):
    """Return the real Excel Chart COM object (xlwings .api is a tuple on Windows)."""
    api = chart.api
    return api[1] if isinstance(api, tuple) else api


def _normalize_type(chart_type: str) -> str:
    key = str(chart_type).lower().strip().replace(" ", "_")
    key = CHART_TYPE_ALIASES.get(key, key)
    if key not in CHART_TYPES:
        raise ExcelToolError(f"Invalid chart_type '{chart_type}'. Use one of: {list(CHART_TYPES)}")
    return key


def _check_size(width, height) -> None:
    if width is not None and not (100 <= width <= 2000):
        raise ExcelToolError("width must be between 100 and 2000.")
    if height is not None and not (100 <= height <= 2000):
        raise ExcelToolError("height must be between 100 and 2000.")


def _single_cell(address: str, label: str) -> str:
    addr = validate_address(address)
    if ":" in addr:
        raise ExcelToolError(f"{label} must be a single cell (example: H2).")
    return addr


def _find_chart(sheet, chart_name: str):
    names = []
    for chart in sheet.charts:
        names.append(chart.name)
        if str(chart.name).lower().strip() == str(chart_name).lower().strip():
            return chart
    raise ExcelToolError(f"Chart '{chart_name}' not found. Available charts: {names}")


def _set_title(chart, title: str) -> None:
    com = _com_chart(chart)
    com.HasTitle = True
    com.ChartTitle.Text = title


@safe_tool
def create_excel_chart(data_range: str, chart_type: str = "column", title: str | None = None,
                       position_cell: str = "H2", width: int = 480, height: int = 300,
                       workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Create a new chart from a data range (include the header row).

    PARAMETERS:
        data_range: Source range with headers, example A1:B10.
        chart_type: column, bar, line, pie, area, scatter or doughnut.
        title: Chart title text (optional).
        position_cell: Top-left cell for the chart, example H2.
        width: Width in points, 100-2000.
        height: Height in points, 100-2000.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, chart_type, data_range, position, name, title_set.
    """
    ctype = _normalize_type(chart_type)
    _check_size(width, height)
    source = validate_address(data_range)
    anchor = _single_cell(position_cell, "position_cell")
    sheet = get_sheet(workbook_name, sheet_name)
    anchor_cell = sheet.range(anchor)
    chart = sheet.charts.add(left=anchor_cell.left, top=anchor_cell.top, width=width, height=height)
    chart.set_source_data(sheet.range(source))
    chart.chart_type = CHART_TYPES[ctype]

    title_set = False
    note = None
    if title:
        if IS_WINDOWS:
            _set_title(chart, title)
            title_set = True
        else:
            note = "Chart title could not be set on this platform (Windows Excel only)."

    result = {"success": True, "chart_type": ctype, "data_range": source,
              "position": anchor, "name": chart.name, "title_set": title_set}
    if note:
        result["note"] = note
    return result


@safe_tool
def modify_excel_chart(chart_name: str,
                       title: str | None = None,
                       x_axis_title: str | None = None,
                       y_axis_title: str | None = None,
                       show_legend: bool | None = None,
                       show_data_labels: bool | None = None,
                       chart_style: int | None = None,
                       chart_type: str | None = None,
                       left_cell: str | None = None,
                       width: int | None = None,
                       height: int | None = None,
                       y_axis_min: float | None = None,
                       y_axis_max: float | None = None,
                       workbook_name: str | None = None,
                       sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Modify an existing chart: title, axis titles, legend, labels, style, type, position, size, axis scale.

    PARAMETERS:
        chart_name: Name of the chart to change (see list_excel_charts).
        title: New chart title (optional).
        x_axis_title: X axis title (optional).
        y_axis_title: Y axis title (optional).
        show_legend: true or false (optional).
        show_data_labels: true or false (optional).
        chart_style: Style number 1-48 (optional).
        chart_type: column, bar, line, pie, area, scatter or doughnut (optional).
        left_cell: Move chart so its top-left is this cell, example J2 (optional).
        width: New width in points, 100-2000 (optional).
        height: New height in points, 100-2000 (optional).
        y_axis_min: Y axis minimum value (optional).
        y_axis_max: Y axis maximum value (optional).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, chart_name, modified_properties, message.
    """
    _check_size(width, height)
    if y_axis_min is not None and y_axis_max is not None and y_axis_min >= y_axis_max:
        raise ExcelToolError("y_axis_min must be less than y_axis_max.")
    if chart_style is not None and not (1 <= int(chart_style) <= 48):
        raise ExcelToolError("chart_style must be between 1 and 48.")

    sheet = get_sheet(workbook_name, sheet_name)
    chart = _find_chart(sheet, chart_name)
    changes = []

    # Cross-platform changes (xlwings only)
    if chart_type is not None:
        chart.chart_type = CHART_TYPES[_normalize_type(chart_type)]
        changes.append("chart_type")
    if left_cell is not None:
        anchor = sheet.range(_single_cell(left_cell, "left_cell"))
        chart.left = anchor.left
        chart.top = anchor.top
        changes.append("position")
    if width is not None:
        chart.width = width
        changes.append("width")
    if height is not None:
        chart.height = height
        changes.append("height")

    # COM-only changes (Windows Excel)
    com_needed = any(v is not None for v in (title, x_axis_title, y_axis_title, show_legend,
                                              show_data_labels, chart_style, y_axis_min, y_axis_max))
    if com_needed:
        _need_com("Title, axis, legend, data label, style and axis scale changes")
        com = _com_chart(chart)
        if title is not None:
            _set_title(chart, title)
            changes.append("title")
        if x_axis_title is not None:
            axis = com.Axes(1)
            axis.HasTitle = True
            axis.AxisTitle.Text = x_axis_title
            changes.append("x_axis_title")
        if y_axis_title is not None:
            axis = com.Axes(2)
            axis.HasTitle = True
            axis.AxisTitle.Text = y_axis_title
            changes.append("y_axis_title")
        if show_legend is not None:
            com.HasLegend = bool(show_legend)
            changes.append("legend")
        if show_data_labels is not None:
            for i in range(1, com.SeriesCollection().Count + 1):
                com.SeriesCollection(i).HasDataLabels = bool(show_data_labels)
            changes.append("data_labels")
        if chart_style is not None:
            com.ChartStyle = int(chart_style)
            changes.append("chart_style")
        if y_axis_min is not None:
            com.Axes(2).MinimumScale = float(y_axis_min)
            changes.append("y_axis_min")
        if y_axis_max is not None:
            com.Axes(2).MaximumScale = float(y_axis_max)
            changes.append("y_axis_max")

    if not changes:
        raise ExcelToolError("At least one chart modification parameter is required.")

    return {"success": True, "chart_name": chart.name, "modified_properties": changes,
            "message": "Chart updated."}


@safe_tool
def list_excel_charts(workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        List all charts on a sheet (read-only).

    PARAMETERS:
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sheet, count, charts (name, chart_type, left, top, width, height, source).
    """
    sheet = get_sheet(workbook_name, sheet_name)
    items = []
    for chart in sheet.charts:
        source = None
        if IS_WINDOWS:
            try:
                com = _com_chart(chart)
                source = com.SeriesCollection(1).Formula
            except Exception:
                source = None
        try:
            ctype = str(chart.chart_type)
        except Exception:
            ctype = None
        items.append({
            "name": chart.name,
            "chart_type": ctype,
            "left": round(chart.left, 1),
            "top": round(chart.top, 1),
            "width": round(chart.width, 1),
            "height": round(chart.height, 1),
            "source": source,
        })
    return {"success": True, "sheet": sheet.name, "count": len(items), "charts": items}


@safe_tool
def delete_excel_chart(chart_name: str, confirm: bool = False,
                       workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Delete a chart. First call shows a preview, second call with confirm=true deletes.

    PARAMETERS:
        chart_name: Name of the chart to delete (see list_excel_charts).
        confirm: Must be true to actually delete.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, deleted (false in preview), chart_name, message.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    chart = _find_chart(sheet, chart_name)
    name = chart.name
    if not confirm:
        return {"success": True, "deleted": False, "chart_name": name,
                "message": f"Chart '{name}' will be deleted. Please confirm."}
    chart.delete()
    return {"success": True, "deleted": True, "chart_name": name,
            "message": f"Chart '{name}' deleted."}


@safe_tool
def change_excel_chart_data(chart_name: str, data_range: str,
                            workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Change the source data range of an existing chart.

    PARAMETERS:
        chart_name: Name of the chart (see list_excel_charts).
        data_range: New source range with headers, example A1:C20.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, chart_name, data_range, message.
    """
    source = validate_address(data_range)
    sheet = get_sheet(workbook_name, sheet_name)
    chart = _find_chart(sheet, chart_name)
    chart.set_source_data(sheet.range(source))
    return {"success": True, "chart_name": chart.name, "data_range": source,
            "message": "Chart source data updated."}