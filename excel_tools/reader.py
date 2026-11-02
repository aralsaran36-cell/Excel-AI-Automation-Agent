from excel_tools.workbook_manager import (
    ExcelToolError, df_to_records, get_sheet, get_workbook, load_dataframe, safe_tool,
)


@safe_tool
def read_excel_data(workbook_name: str | None = None, sheet_name: str | None = None,
                    max_rows: int = 50) -> dict:
    """Read data from a sheet (first row = headers). Returns columns and up to max_rows rows."""
    if not isinstance(max_rows, int) or not 1 <= max_rows <= 1000:
        raise ExcelToolError("max_rows 1 mudhal 1000 varaikkum irukkanum.")
    sheet = get_sheet(workbook_name, sheet_name)
    df, _ = load_dataframe(sheet)
    return {
        "success": True,
        "workbook": sheet.book.name,
        "sheet": sheet.name,
        "columns": list(df.columns),
        "total_rows": len(df),
        "returned_rows": min(len(df), max_rows),
        "rows": df_to_records(df, max_rows),
    }


@safe_tool
def get_excel_sheet_names(workbook_name: str | None = None) -> dict:
    """List sheet names of a workbook (active workbook by default)."""
    book = get_workbook(workbook_name)
    return {
        "success": True,
        "workbook": book.name,
        "sheets": [s.name for s in book.sheets],
        "active_sheet": book.sheets.active.name,
        "open_workbooks": [b.name for b in book.app.books],
    }