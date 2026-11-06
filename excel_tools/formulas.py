import re

from excel_tools.workbook_manager import (
    EXCEL_ERRORS,
    ExcelToolError,
    backup_workbook,
    contiguous_blocks,
    get_sheet,
    load_dataframe,
    safe_tool,
    validate_address,
)

BACKUP_NOTE = "Backup is a copy of the last saved file. Unsaved edits are not in it."
COL_RE = re.compile(r"^[A-Za-z]{1,3}$")
MAX_CELLS = 100000


# ---------- helpers ----------

def _check_formula(formula: str) -> None:
    if not isinstance(formula, str) or not formula.startswith("=") or len(formula) < 2:
        raise ExcelToolError("Formula must start with '='. Example: =SUM(B2:B10)")


def _has_data(rng) -> bool:
    return any(v is not None for row in rng.options(ndim=2).value for v in row)


def _clean(address: str) -> str:
    return str(address).replace("$", "")


def _col_letter(number: int) -> str:
    letters = ""
    while number > 0:
        number, rem = divmod(number - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _is_error(value) -> bool:
    return isinstance(value, str) and value in EXCEL_ERRORS


def _formula_is_error(target) -> bool | None:
    """Ask Excel itself via ISERROR, because xlwings on macOS returns None for error cells.
    Uses an empty neighbour cell temporarily and clears its contents. Returns None if no free cell."""
    candidates = [target.offset(0, 1)]
    if target.row > 1:
        candidates.append(target.offset(-1, 0))
    for probe in candidates:
        if probe.value is None and not probe.formula:
            try:
                probe.formula = f"=ISERROR({_clean(target.address)})"
                return bool(probe.value)
            finally:
                probe.clear_contents()
    return None


def _write_cells(sheet, updates: dict) -> None:
    """updates: {(excel_row, excel_col): value}. Writes contiguous runs per column."""
    by_col = {}
    for (row, col), value in updates.items():
        by_col.setdefault(col, {})[row] = value
    for col, rows in by_col.items():
        for start, end in contiguous_blocks(sorted(rows)):
            sheet.range((start, col), (end, col)).value = [[rows[r]] for r in range(start, end + 1)]


# ---------- tools ----------

@safe_tool
def apply_excel_formula(cell: str, formula: str, overwrite: bool = False,
                        workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Put a formula in one cell and return the calculated result.

    PARAMETERS:
        cell: Single cell, example E2.
        formula: Formula starting with =, example =SUM(B2:B10).
        overwrite: true to replace an existing value (a backup is made).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, cell, formula, result, backup.
    """
    _check_formula(formula)
    address = validate_address(cell)
    if ":" in address:
        raise ExcelToolError("Give a single cell. For a range use apply_excel_formula_range.")
    sheet = get_sheet(workbook_name, sheet_name)
    target = sheet.range(address)
    backup = None
    if target.value is not None:
        if not overwrite:
            raise ExcelToolError(f"{address} already has a value. Use overwrite=true to replace it.")
        backup = backup_workbook(sheet.book)
    target.formula = formula
    result = target.value
    if _is_error(result):
        return {"success": False, "error": f"Formula error {result} in {address}", "formula": formula}
    if result is None:
        is_error = _formula_is_error(target)
        if is_error:
            return {"success": False,
                    "error": f"Formula error in {address} (Excel returned an error value).",
                    "formula": formula}
        if is_error is None:
            return {"success": True, "cell": address, "formula": formula, "result": None,
                    "backup": backup,
                    "warning": "Result empty; could not verify error state (no free neighbour cell)."}
    out = {"success": True, "cell": address, "formula": formula, "result": result, "backup": backup}
    if backup:
        out["note"] = BACKUP_NOTE
    return out


@safe_tool
def apply_excel_formula_range(target_range: str, formula: str, overwrite: bool = False,
                              workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Fill a range with a formula written for its first cell. References adjust per row and column.

    PARAMETERS:
        target_range: Range to fill, example E2:E100.
        formula: Formula for the first cell, example =B2*C2.
        overwrite: true to replace existing data (a backup is made).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, formula, cells_filled, error_cells, sample_results, backup, warning.
    """
    _check_formula(formula)
    address = validate_address(target_range)
    sheet = get_sheet(workbook_name, sheet_name)
    target = sheet.range(address)
    backup = None
    if _has_data(target):
        if not overwrite:
            raise ExcelToolError(f"{address} already has data. Use overwrite=true to replace it.")
        backup = backup_workbook(sheet.book)
    target.formula = formula
    values = [v for row in target.options(ndim=2).value for v in row]
    error_count = sum(1 for v in values if _is_error(v))
    out = {"success": True, "range": address, "formula": formula, "cells_filled": len(values),
           "error_cells": error_count, "sample_results": values[:5], "backup": backup}
    if error_count:
        out["warning"] = (f"Formula was written, but {error_count} cell(s) show an error. "
                          "Check the formula references.")
    if backup:
        out["note"] = BACKUP_NOTE
    return out


@safe_tool
def add_excel_formula_column(header: str, formula: str, column: str | None = None,
                             overwrite: bool = False, workbook_name: str | None = None,
                             sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Add a calculated column: write a header and fill the formula down to the last data row.

    PARAMETERS:
        header: New column header, example Total.
        formula: Formula for the first data row, example =B2*C2.
        column: Column letter to use (optional, default first free column after the data).
        overwrite: true to replace an existing column (a backup is made).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, column, header, range, rows_filled, error_cells, sample_results, backup.
    """
    header = str(header).strip()
    if not header:
        raise ExcelToolError("header cannot be empty.")
    _check_formula(formula)
    if column is not None and not COL_RE.match(str(column).strip()):
        raise ExcelToolError("column must be a column letter like H.")

    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        raise ExcelToolError("No data found. Add data first.")

    header_row = used.row
    first_row = header_row + 1
    last_row = used.last_cell.row
    if column is None:
        col_num = used.last_cell.column + 1
    else:
        col_num = sheet.range(f"{str(column).strip().upper()}1").column
    letter = _col_letter(col_num)

    header_cell = sheet.range((header_row, col_num))
    data = sheet.range((first_row, col_num), (last_row, col_num))
    backup = None
    if header_cell.value is not None or _has_data(data):
        if not overwrite:
            raise ExcelToolError(
                f"Column {letter} already has data. Use overwrite=true or choose another column.")
        backup = backup_workbook(sheet.book)

    header_cell.value = header
    data.formula = formula

    values = [v for row in data.options(ndim=2).value for v in row]
    error_count = sum(1 for v in values if _is_error(v))
    out = {"success": True, "column": letter, "header": header,
           "range": f"{letter}{first_row}:{letter}{last_row}", "rows_filled": len(values),
           "formula": formula, "error_cells": error_count, "sample_results": values[:5],
           "backup": backup}
    if error_count:
        out["warning"] = (f"Formula was written, but {error_count} cell(s) show an error. "
                          "Check the formula references.")
    if backup:
        out["note"] = BACKUP_NOTE
    return out


@safe_tool
def get_excel_formula(cell_range: str, workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Show the formulas and current values of cells in a range (read-only).

    PARAMETERS:
        cell_range: Cell or range, example E2 or E2:E20.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, formula_cells, items (cell, formula, value, first 50).
    """
    address = validate_address(cell_range)
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)
    if rng.count > 20000:
        raise ExcelToolError("Range too large. Use at most 20000 cells.")
    formulas = rng.options(ndim=2).formula
    values = rng.options(ndim=2).value
    items = []
    total = 0
    for r, row in enumerate(formulas):
        for c, text in enumerate(row):
            if isinstance(text, str) and text.startswith("="):
                total += 1
                if len(items) < 50:
                    items.append({"cell": _clean(sheet.range((rng.row + r, rng.column + c)).address),
                                  "formula": text, "value": values[r][c]})
    return {"success": True, "range": address, "formula_cells": total, "items": items}


@safe_tool
def convert_excel_formulas_to_values(cell_range: str, workbook_name: str | None = None,
                                     sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Replace formulas in a range with their current values (applies directly, a backup is made).

    PARAMETERS:
        cell_range: Range, example E2:E100.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, converted, skipped_unreadable, backup.
    """
    address = validate_address(cell_range)
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)
    if rng.count > MAX_CELLS:
        raise ExcelToolError(f"Range too large. Use at most {MAX_CELLS} cells.")
    formulas = rng.options(ndim=2).formula
    values = rng.options(ndim=2).value

    updates = {}
    skipped = 0
    for r, row in enumerate(formulas):
        for c, text in enumerate(row):
            if not (isinstance(text, str) and text.startswith("=")):
                continue
            value = values[r][c]
            if value is None or _is_error(value):
                skipped += 1  # empty or error result: keep the formula, do not lose it
                continue
            updates[(rng.row + r, rng.column + c)] = value

    if not updates:
        return {"success": True, "range": address, "converted": 0, "skipped_unreadable": skipped,
                "message": "No formulas to convert."}

    backup = backup_workbook(sheet.book)
    _write_cells(sheet, updates)
    return {"success": True, "range": address, "converted": len(updates),
            "skipped_unreadable": skipped, "backup": backup, "note": BACKUP_NOTE}