from excel_tools.workbook_manager import (
    ExcelToolError, build_condition_mask, contiguous_blocks, get_sheet, hex_to_rgb,
    load_dataframe, safe_tool,
)


@safe_tool
def highlight_excel_values(column: str, condition: str, value: str | float | None = None,
                           color: str = "#FFFF00", workbook_name: str | None = None,
                           sheet_name: str | None = None) -> dict:
    """Color cells in a column that match a condition (equals, not_equals, greater_than,
    less_than, greater_or_equal, less_or_equal, contains, is_empty). Static fill color."""
    rgb = hex_to_rgb(color)
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if column not in df.columns:
        raise ExcelToolError(f"Column '{column}' illa. Available: {list(df.columns)}")
    mask = build_condition_mask(df[column], condition, value)
    matched = [used.row + 1 + int(i) for i in mask[mask].index]
    if not matched:
        return {"success": True, "highlighted_cells": 0, "message": "Match aana cells illa."}
    col_number = used.column + list(df.columns).index(column)
    for start, end in contiguous_blocks(matched):
        sheet.range((start, col_number), (end, col_number)).color = rgb
    return {"success": True, "column": column, "condition": condition,
            "highlighted_cells": len(matched), "color": color}