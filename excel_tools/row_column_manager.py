import sys

from excel_tools.workbook_manager import ExcelToolError, backup_workbook, get_sheet, safe_tool

IS_WINDOWS = sys.platform.startswith("win")
MAX_COUNT = 1000
MAX_ROW = 1048576
MAX_COL = 16384
BACKUP_NOTE = "Backup is a copy of the last saved file. Unsaved edits are not in it."
REF_NOTE = "Formulas that pointed to the deleted cells may show #REF!."


# ---------- helpers ----------

def _check_index(value, label: str) -> int:
    if not isinstance(value, int) or isinstance(value, bool) or value < 1:
        raise ExcelToolError(f"{label} must be a whole number 1 or higher.")
    return value


def _check_count(count) -> int:
    if not isinstance(count, int) or isinstance(count, bool) or not 1 <= count <= MAX_COUNT:
        raise ExcelToolError(f"count must be between 1 and {MAX_COUNT}.")
    return count


def _col_index(column) -> int:
    """Column number (1-based), digit text ('3') or letters ('C') -> number."""
    if isinstance(column, int) and not isinstance(column, bool):
        return _check_index(column, "column")
    text = str(column).strip()
    if text.isdigit():
        return _check_index(int(text), "column")
    if text.isalpha() and len(text) <= 3:
        n = 0
        for ch in text.upper():
            n = n * 26 + (ord(ch) - 64)
        return n
    raise ExcelToolError("column must be a number (3) or a letter (C).")


def _letter(n: int) -> str:
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _row_span(text: str) -> tuple[int, int]:
    parts = str(text).strip().split(":")
    if len(parts) > 2 or not all(p.strip().isdigit() and int(p) >= 1 for p in parts):
        raise ExcelToolError("rows format: '2' or '2:10'.")
    nums = sorted(int(p) for p in parts)
    if nums[1] > MAX_ROW:
        raise ExcelToolError(f"Row number cannot be more than {MAX_ROW}.")
    return nums[0], nums[1]


def _col_span(text: str) -> tuple[int, int]:
    parts = str(text).strip().split(":")
    if len(parts) > 2:
        raise ExcelToolError("columns format: 'C' or 'C:F'.")
    nums = sorted(_col_index(p) for p in parts)
    if nums[1] > MAX_COL:
        raise ExcelToolError("Column cannot be beyond XFD.")
    return nums[0], nums[1]


def _data_in(sheet, first_row, last_row, first_col, last_col) -> int:
    """Count non-empty cells, reading only the part that overlaps the used range (fast)."""
    used = sheet.used_range
    r1, r2 = max(first_row, used.row), min(last_row, used.last_cell.row)
    c1, c2 = max(first_col, used.column), min(last_col, used.last_cell.column)
    if r1 > r2 or c1 > c2:
        return 0
    value = sheet.range((r1, c1), (r2, c2)).options(ndim=2).value
    return sum(1 for row in value for c in row if c not in (None, ""))


# ---------- tools ----------

@safe_tool
def insert_excel_rows(row: int, count: int = 1, after: bool = False,
                      workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Insert empty rows. Existing rows shift down.

    PARAMETERS:
        row: Row number, example 5.
        count: How many rows to insert, 1-1000.
        after: true to insert after this row instead of before it.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sheet, inserted_at, count.
    """
    row, count = _check_index(row, "row"), _check_count(count)
    start = row + 1 if after else row
    if start + count - 1 > MAX_ROW:
        raise ExcelToolError("Not enough rows left in the sheet.")
    sheet = get_sheet(workbook_name, sheet_name)
    sheet.range(f"{start}:{start + count - 1}").insert()
    return {"success": True, "sheet": sheet.name, "inserted_at": start, "count": count}


@safe_tool
def insert_excel_columns(column: int | str, count: int = 1, after: bool = False,
                         workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Insert empty columns. Existing columns shift right.

    PARAMETERS:
        column: Column letter or number, example C or 3.
        count: How many columns to insert, 1-1000.
        after: true to insert after this column instead of before it.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sheet, inserted_at, count.
    """
    col, count = _col_index(column), _check_count(count)
    start = col + 1 if after else col
    if start + count - 1 > MAX_COL:
        raise ExcelToolError("Not enough columns left in the sheet.")
    sheet = get_sheet(workbook_name, sheet_name)
    sheet.range(f"{_letter(start)}:{_letter(start + count - 1)}").insert()
    return {"success": True, "sheet": sheet.name, "inserted_at": _letter(start), "count": count}


@safe_tool
def delete_excel_rows(row: int, count: int = 1, confirm: bool = False,
                      workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Delete rows. First call shows a preview, second call with confirm=true deletes.

    PARAMETERS:
        row: First row number to delete, example 5.
        count: How many rows to delete, 1-1000.
        confirm: Must be true to actually delete.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, deleted, rows, cells_with_data, backup.
    """
    row, count = _check_index(row, "row"), _check_count(count)
    last = row + count - 1
    if last > MAX_ROW:
        raise ExcelToolError("Rows go beyond the end of the sheet.")
    sheet = get_sheet(workbook_name, sheet_name)
    cells = _data_in(sheet, row, last, 1, MAX_COL)
    span = f"{row}-{last}"
    if not confirm:
        return {"success": True, "deleted": False, "needs_confirm": True, "rows": span,
                "cells_with_data": cells,
                "message": f"Rows {span} have {cells} cell(s) with data. Confirm to delete."}
    backup = backup_workbook(sheet.book)
    sheet.range(f"{row}:{last}").delete()
    return {"success": True, "deleted": True, "rows": span, "cells_with_data": cells,
            "backup": backup, "note": REF_NOTE + " " + BACKUP_NOTE}


@safe_tool
def delete_excel_columns(column: int | str, count: int = 1, confirm: bool = False,
                         workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Delete columns. First call shows a preview, second call with confirm=true deletes.

    PARAMETERS:
        column: First column letter or number to delete, example C or 3.
        count: How many columns to delete, 1-1000.
        confirm: Must be true to actually delete.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, deleted, columns, cells_with_data, backup.
    """
    col, count = _col_index(column), _check_count(count)
    last = col + count - 1
    if last > MAX_COL:
        raise ExcelToolError("Columns go beyond the end of the sheet.")
    sheet = get_sheet(workbook_name, sheet_name)
    span = f"{_letter(col)}:{_letter(last)}"
    cells = _data_in(sheet, 1, MAX_ROW, col, last)
    if not confirm:
        return {"success": True, "deleted": False, "needs_confirm": True, "columns": span,
                "cells_with_data": cells,
                "message": f"Columns {span} have {cells} cell(s) with data. Confirm to delete."}
    backup = backup_workbook(sheet.book)
    sheet.range(span).delete()
    return {"success": True, "deleted": True, "columns": span, "cells_with_data": cells,
            "backup": backup, "note": REF_NOTE + " " + BACKUP_NOTE}


@safe_tool
def hide_unhide_excel_rows_columns(rows: str | None = None, columns: str | None = None,
                                   hide: bool = True, workbook_name: str | None = None,
                                   sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Hide or unhide rows and/or columns. Data is not deleted.

    PARAMETERS:
        rows: Row or span, example 3 or 2:5 (optional).
        columns: Column or span, example C or C:F (optional).
        hide: true to hide, false to unhide.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sheet, hidden, rows, columns.
    """
    if rows is None and columns is None:
        raise ExcelToolError("Give rows and/or columns.")
    row_span = col_span = None
    if rows is not None:
        r1, r2 = _row_span(rows)
        row_span = f"{r1}:{r2}"
    if columns is not None:
        c1, c2 = _col_span(columns)
        col_span = f"{_letter(c1)}:{_letter(c2)}"

    sheet = get_sheet(workbook_name, sheet_name)
    hide = bool(hide)
    if row_span:
        rng = sheet.range(row_span)
        if IS_WINDOWS:
            rng.api.EntireRow.Hidden = hide
        elif hide:
            rng.row_height = 0          # Mac fallback (untested)
        else:
            rng.rows.autofit()
    if col_span:
        rng = sheet.range(col_span)
        if IS_WINDOWS:
            rng.api.EntireColumn.Hidden = hide
        elif hide:
            rng.column_width = 0        # Mac fallback (untested)
        else:
            rng.columns.autofit()
    return {"success": True, "sheet": sheet.name, "hidden": hide,
            "rows": row_span, "columns": col_span}