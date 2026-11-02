from excel_tools.workbook_manager import (
    ExcelToolError, build_condition_mask, get_sheet, get_workbook, load_dataframe,
    safe_tool, write_dataframe_to_sheet,
)

FILTER_PREFIX = "Filtered"


@safe_tool
def filter_excel_data(column: str, condition: str, value: str | float | None = None,
                      output_sheet: str = "Filtered_Data", workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """Filter rows and write matches to a separate sheet (original data untouched).
    Conditions: equals, not_equals, greater_than, less_than, greater_or_equal,
    less_or_equal, contains, is_empty."""
    sheet = get_sheet(workbook_name, sheet_name)
    if output_sheet == sheet.name:
        raise ExcelToolError("output_sheet source sheet-oda same-a irukka koodathu.")
    df, _ = load_dataframe(sheet)
    if column not in df.columns:
        raise ExcelToolError(f"Column '{column}' illa. Available: {list(df.columns)}")
    result = df[build_condition_mask(df[column], condition, value)]
    if result.empty:
        return {"success": True, "matched_rows": 0, "message": "Match aana rows illa. Sheet create aagala."}
    write_dataframe_to_sheet(sheet.book, output_sheet, result, FILTER_PREFIX)
    return {"success": True, "matched_rows": len(result), "total_rows": len(df),
            "output_sheet": output_sheet}


@safe_tool
def clear_excel_filter(output_sheet: str = "Filtered_Data",
                       workbook_name: str | None = None) -> dict:
    """Remove the filter result sheet created by filter_excel_data. Original data is untouched."""
    if not output_sheet.startswith(FILTER_PREFIX):
        raise ExcelToolError(f"Safety: '{FILTER_PREFIX}' la start aagura sheet-a mattum thaan delete pannuven.")
    book = get_workbook(workbook_name)
    if output_sheet not in [s.name for s in book.sheets]:
        return {"success": True, "removed": False, "message": f"'{output_sheet}' sheet illa."}
    if len(book.sheets) == 1:
        raise ExcelToolError("Idhu thaan oru-ey sheet. Delete panna mudiyaadhu.")
    book.app.display_alerts = False
    try:
        book.sheets[output_sheet].delete()
    finally:
        book.app.display_alerts = True
    return {"success": True, "removed": True, "sheet": output_sheet}