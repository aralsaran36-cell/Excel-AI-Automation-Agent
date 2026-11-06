"""Writer tools: write a block, append rows, update one cell, clear a range, fill a series.

All tools return {"success": bool, ...} and never raise (via @safe_tool).
Tools that replace existing data take a backup first (the last SAVED file on disk).
"""
from datetime import date, datetime

import pandas as pd

from excel_tools.workbook_manager import (
    ExcelToolError, backup_workbook, get_sheet, safe_tool, validate_address,
)

MAX_SERIES_COUNT = 10000
MAX_CLEAR_CELLS = 200000
SERIES_TYPES = ("number", "date")
DATE_UNITS = {"day": "days", "week": "weeks", "month": "months", "year": "years"}


# ---------------- helpers ----------------

def _validate_rows(data) -> int:
    """Check that data is a rectangular list of rows. Returns the width."""
    if not isinstance(data, list) or not data or not all(isinstance(r, list) for r in data):
        raise ExcelToolError(
            "data must be a list of rows. Example: [['Name','Sales'],['A',100]]")
    width = len(data[0])
    if width == 0 or any(len(r) != width for r in data):
        raise ExcelToolError("All rows must have the same number of columns.")
    return width


def _has_data(rng) -> bool:
    """True if any cell in the range holds a value."""
    values = rng.options(ndim=2).value
    return any(v not in (None, "") for row in values for v in row)


def _count_data(rng) -> int:
    """Number of non-empty cells in the range."""
    values = rng.options(ndim=2).value
    return sum(1 for row in values for v in row if v not in (None, ""))


def _json_value(value):
    """Make a cell value safe for JSON."""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, float) and value != value:  # NaN
        return None
    return value


def _single_cell(address: str, label: str) -> str:
    """Validate that an address is one cell, not a range."""
    cell = validate_address(address)
    if ":" in cell:
        raise ExcelToolError(f"{label} must be a single cell (example: A1).")
    return cell


# ---------------- 1. WRITE A BLOCK ----------------

@safe_tool
def write_excel_data(data: list[list], start_cell: str = "A1", overwrite: bool = False,
                     workbook_name: str | None = None,
                     sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Write rows of values starting at start_cell. Refuses to overwrite existing data unless overwrite=True.

    PARAMETERS:
        data: List of rows, for example [['Name','Sales'],['A',100]]. All rows need the same length.
        start_cell: Top-left cell, a single cell like 'A1'.
        overwrite: Set True only after the user confirmed replacing existing data.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Target sheet. Uses the active sheet if omitted.

    RETURNS:
        Sheet, written range, rows written and backup path (if data was replaced).
    """
    width = _validate_rows(data)
    cell = _single_cell(start_cell, "start_cell")
    sheet = get_sheet(workbook_name, sheet_name)
    target = sheet.range(cell).resize(len(data), width)
    has_data = _has_data(target)
    if has_data and not overwrite:
        raise ExcelToolError(
            f"{target.address} already contains data. Set overwrite=True to replace it.")
    backup = backup_workbook(sheet.book) if has_data else None
    target.value = data
    return {"success": True, "sheet": sheet.name, "written_range": target.address,
            "rows_written": len(data), "backup": backup}


# ---------------- 2. APPEND ROWS ----------------

@safe_tool
def append_excel_data(rows: list[list], workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Append rows below the last row that actually contains data.

    PARAMETERS:
        rows: List of rows to add, for example [['Ravi', 90], ['Anu', 85]].
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Target sheet. Uses the active sheet if omitted.

    RETURNS:
        Sheet, appended range and number of rows appended.
    """
    width = _validate_rows(rows)
    sheet = get_sheet(workbook_name, sheet_name)
    used = sheet.used_range
    if used.count == 1 and used.value is None:
        start_row, start_col = 1, 1
    else:
        if width > used.columns.count:
            raise ExcelToolError(
                f"The rows have {width} columns but the sheet has only {used.columns.count}.")
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


# ---------------- 3. UPDATE ONE CELL ----------------

@safe_tool
def update_excel_cell(cell: str, value: str | float | None,
                      overwrite: bool = False, workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Set the value of ONE cell (text, number or empty). Returns the old value.

    PARAMETERS:
        cell: A single cell like 'B5'.
        value: New text or number. Use null to empty the cell.
        overwrite: Needed only when the cell holds a formula. Set True after the user confirmed.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Target sheet. Uses the active sheet if omitted.

    RETURNS:
        Sheet, cell, old value, new value and backup path.

    NOTES:
        - Formulas are not written here. Use apply_excel_formula.
        - To change many cells at once use write_excel_data.
    """
    address = _single_cell(cell, "cell")
    if isinstance(value, str) and value.lstrip().startswith("="):
        raise ExcelToolError("This tool does not write formulas. Use apply_excel_formula.")
    sheet = get_sheet(workbook_name, sheet_name)
    target = sheet.range(address)
    formula = target.formula
    if isinstance(formula, str) and formula.startswith("=") and not overwrite:
        raise ExcelToolError(
            f"{address} contains a formula. Set overwrite=True to replace it.")
    old = target.value
    backup = backup_workbook(sheet.book) if old is not None else None
    target.value = value
    return {"success": True, "sheet": sheet.name, "cell": address,
            "old_value": _json_value(old), "new_value": value, "backup": backup}


# ---------------- 4. CLEAR A RANGE ----------------

@safe_tool
def clear_excel_range(cell_range: str, confirm: bool = False,
                      workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Clear the values of a range but keep its formatting. Without confirm=True it only previews.

    PARAMETERS:
        cell_range: Range like 'A2:D50' or a single cell.
        confirm: False = preview only. True = really clear. Set True only after the user confirmed.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Target sheet. Uses the active sheet if omitted.

    RETURNS:
        Preview (found = number of non-empty cells) or the cleared range and backup path.

    NOTES:
        - Formatting stays. The cells and rows are not deleted.
    """
    address = validate_address(cell_range)
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)
    if int(rng.count) > MAX_CLEAR_CELLS:
        raise ExcelToolError(
            f"Range is too large (limit {MAX_CLEAR_CELLS} cells). Use a smaller range.")
    found = _count_data(rng)
    if found == 0:
        return {"success": True, "cleared": False, "found": 0, "range": rng.address,
                "message": f"{rng.address} has no data. Nothing to clear."}
    if not confirm:
        return {"success": True, "cleared": False, "found": found, "range": rng.address,
                "message": f"{rng.address} has {found} cells with data. "
                           "Call again with confirm=True to clear them."}
    backup = backup_workbook(sheet.book)
    rng.clear_contents()
    return {"success": True, "cleared": True, "found": found, "range": rng.address,
            "backup": backup,
            "message": f"Cleared {found} cells in {rng.address}."}


# ---------------- 5. FILL A SERIES ----------------

@safe_tool
def fill_excel_series(start_cell: str, count: int, series_type: str = "number",
                      start_value: str | float | None = None, step: float = 1,
                      step_unit: str = "day", direction: str = "down",
                      overwrite: bool = False, workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Fill a column or row with a number series (1, 2, 3 ...) or a date series.

    PARAMETERS:
        start_cell: First cell of the series, a single cell like 'A2'.
        count: How many values to write, 1 to 10000.
        series_type: 'number' or 'date'.
        start_value: First value. Number default 1. For dates use 'YYYY-MM-DD' (default today).
        step: Difference between values. For dates it is a whole number of step_unit.
        step_unit: For dates only: 'day', 'week', 'month' or 'year'.
        direction: 'down' (a column) or 'right' (a row).
        overwrite: Set True only after the user confirmed replacing existing data.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Target sheet. Uses the active sheet if omitted.

    RETURNS:
        Written range, count, first and last value, backup path.
    """
    cell = _single_cell(start_cell, "start_cell")
    if isinstance(count, bool) or not isinstance(count, int) or not 1 <= count <= MAX_SERIES_COUNT:
        raise ExcelToolError(f"count must be a whole number from 1 to {MAX_SERIES_COUNT}.")
    kind = str(series_type).strip().lower()
    if kind not in SERIES_TYPES:
        raise ExcelToolError("series_type must be 'number' or 'date'.")
    way = str(direction).strip().lower()
    if way not in ("down", "right"):
        raise ExcelToolError("direction must be 'down' or 'right'.")
    if isinstance(step, bool) or not isinstance(step, (int, float)) or step == 0:
        raise ExcelToolError("step must be a non-zero number.")

    if kind == "number":
        try:
            first = 1 if start_value is None else float(start_value)
        except (TypeError, ValueError):
            raise ExcelToolError("start_value must be a number for a number series.")
        whole = float(first).is_integer() and float(step).is_integer()
        values = [int(first + i * step) if whole else round(first + i * step, 10)
                  for i in range(count)]
    else:
        unit = str(step_unit).strip().lower()
        if unit not in DATE_UNITS:
            raise ExcelToolError("step_unit must be 'day', 'week', 'month' or 'year'.")
        if not float(step).is_integer():
            raise ExcelToolError("step must be a whole number for a date series.")
        try:
            base = pd.Timestamp.today().normalize() if start_value is None \
                else pd.Timestamp(str(start_value)).normalize()
        except (ValueError, TypeError):
            raise ExcelToolError("start_value must be a date like '2026-11-01'.")
        key = DATE_UNITS[unit]
        values = [(base + pd.DateOffset(**{key: i * int(step)})).to_pydatetime()
                  for i in range(count)]

    sheet = get_sheet(workbook_name, sheet_name)
    anchor = sheet.range(cell)
    target = anchor.resize(count, 1) if way == "down" else anchor.resize(1, count)
    has_data = _has_data(target)
    if has_data and not overwrite:
        raise ExcelToolError(
            f"{target.address} already contains data. Set overwrite=True to replace it.")
    backup = backup_workbook(sheet.book) if has_data else None

    target.value = [[v] for v in values] if way == "down" else [values]
    if kind == "date":
        target.number_format = "yyyy-mm-dd"

    return {"success": True, "sheet": sheet.name, "written_range": target.address,
            "count": count, "first": _json_value(values[0]),
            "last": _json_value(values[-1]), "backup": backup}