from excel_tools.workbook_manager import (
    ExcelToolError, backup_workbook, get_sheet, safe_tool, validate_address,
)


def _validate_rows(data) -> int:
    if not isinstance(data, list) or not data or not all(isinstance(r, list) for r in data):
        raise ExcelToolError("data oru list of rows aa irukkanum. Example: [['Name','Sales'],['A',100]]")
    width = len(data[0])
    if width == 0 or any(len(r) != width for r in data):
        raise ExcelToolError("Ellaa rows-kum same number of columns irukkanum.")
    return width


@safe_tool
def write_excel_data(data: list[list], start_cell: str = "A1", overwrite: bool = False,
                     workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """Write rows starting at start_cell. Refuses to overwrite existing data unless overwrite=True."""
    width = _validate_rows(data)
    cell = validate_address(start_cell)
    if ":" in cell:
        raise ExcelToolError("start_cell oru single cell aa irukkanum (example: A1).")
    sheet = get_sheet(workbook_name, sheet_name)
    target = sheet.range(cell).resize(len(data), width)
    existing = target.options(ndim=2).value
    has_data = any(v is not None for row in existing for v in row)
    if has_data and not overwrite:
        raise ExcelToolError(f"{target.address}-la already data irukku. Overwrite pannanum-na overwrite=True kudunga.")
    backup = backup_workbook(sheet.book) if has_data else None
    target.value = data
    return {"success": True, "sheet": sheet.name, "written_range": target.address,
            "rows_written": len(data), "backup": backup}


@safe_tool
def append_excel_data(rows: list[list], workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """Append rows below the last row that actually contains data."""
    width = _validate_rows(rows)
    sheet = get_sheet(workbook_name, sheet_name)
    used = sheet.used_range
    if used.count == 1 and used.value is None:
        start_row, start_col = 1, 1
    else:
        if width > used.columns.count:
            raise ExcelToolError(f"Rows-la {width} columns irukku, aana sheet-la {used.columns.count} columns thaan.")
        start_col = used.column
        first_row = used.row
        values = used.options(ndim=2).value
        last_data_row = first_row - 1
        for offset in range(len(values) - 1, -1, -1):
            if any(v is not None and v != "" for v in values[offset]):
                last_data_row = first_row + offset
                break
        start_row = last_data_row + 1
    target = sheet.range((start_row, start_col)).resize(len(rows), width)
    target.value = rows
    return {"success": True, "sheet": sheet.name, "appended_range": target.address,
            "rows_appended": len(rows)}