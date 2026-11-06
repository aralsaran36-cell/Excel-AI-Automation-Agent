"""Reader tools: read sheet data, read a range, describe columns, search values.

All tools are read-only and return {"success": bool, ...} (via @safe_tool).
"""
from datetime import date, datetime

import pandas as pd

from excel_tools.workbook_manager import (
    ExcelToolError, df_to_records, get_sheet, get_workbook, load_dataframe,
    safe_tool, validate_address,
)

MAX_ROWS_LIMIT = 1000
MAX_RANGE_CELLS = 20000
MAX_COLUMNS_INFO = 100
MAX_SEARCH_RESULTS = 100
SAMPLE_COUNT = 3
SAMPLE_CHARS = 30


# ---------------- helpers ----------------

def _check_limit(value, label: str, limit: int) -> int:
    """Validate a whole number from 1 to limit (bool rejected)."""
    if isinstance(value, bool) or not isinstance(value, int) or not 1 <= value <= limit:
        raise ExcelToolError(f"{label} must be a whole number from 1 to {limit}.")
    return value


def _col_letter(number: int) -> str:
    """1 -> A, 26 -> Z, 27 -> AA."""
    letters = ""
    while number:
        number, rem = divmod(number - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _json_value(value):
    """Make a cell value safe for JSON."""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, float) and value != value:  # NaN
        return None
    return value


def _text(value) -> str:
    """Cell value as plain text (5.0 -> '5')."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value)


# ---------------- 1. READ SHEET DATA ----------------

@safe_tool
def read_excel_data(workbook_name: str | None = None, sheet_name: str | None = None,
                    max_rows: int = 50) -> dict:
    """
    PURPOSE:
        Read data from a sheet (first row = headers). Returns the column names and up to max_rows rows.

    PARAMETERS:
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Sheet to read. Uses the active sheet if omitted.
        max_rows: Number of data rows to return, 1 to 1000.

    RETURNS:
        Columns, total_rows, returned_rows and the rows.
    """
    _check_limit(max_rows, "max_rows", MAX_ROWS_LIMIT)
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


# ---------------- 2. SHEET NAMES ----------------

@safe_tool
def get_excel_sheet_names(workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        List the sheet names of a workbook.

    PARAMETERS:
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Workbook name, sheet names, active sheet and open workbooks.
    """
    book = get_workbook(workbook_name)
    return {
        "success": True,
        "workbook": book.name,
        "sheets": [s.name for s in book.sheets],
        "active_sheet": book.sheets.active.name,
        "open_workbooks": [b.name for b in book.app.books],
    }


# ---------------- 3. READ A RANGE ----------------

@safe_tool
def read_excel_range(cell_range: str, max_rows: int = 50,
                     workbook_name: str | None = None,
                     sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Read the raw cell values of a specific range such as A1:D10. No header assumption.

    PARAMETERS:
        cell_range: Range like 'A1:D10' or a single cell like 'B5'.
        max_rows: Maximum number of rows to return, 1 to 1000.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Sheet to read. Uses the active sheet if omitted.

    RETURNS:
        Range address, total rows and columns, and the values as a list of rows.

    NOTES:
        - Read-only. Dates are returned as ISO text.
        - Large ranges are cut to max_rows rows.
    """
    address = validate_address(cell_range)
    _check_limit(max_rows, "max_rows", MAX_ROWS_LIMIT)
    sheet = get_sheet(workbook_name, sheet_name)

    rng = sheet.range(address)
    total_rows = int(rng.rows.count)
    total_cols = int(rng.columns.count)
    shown = min(total_rows, max_rows)
    if shown * total_cols > MAX_RANGE_CELLS:
        raise ExcelToolError(
            f"Range is too large to read at once (limit {MAX_RANGE_CELLS} cells). "
            "Use a smaller range or a smaller max_rows."
        )

    raw = rng.resize(shown, total_cols).options(ndim=2).value
    values = [[_json_value(cell) for cell in row] for row in raw]

    result = {
        "success": True,
        "sheet": sheet.name,
        "range": rng.address,
        "total_rows": total_rows,
        "total_columns": total_cols,
        "returned_rows": shown,
        "values": values,
    }
    if shown < total_rows:
        result["note"] = f"Only the first {shown} of {total_rows} rows are shown."
    return result


# ---------------- 4. COLUMN INFO ----------------

@safe_tool
def get_excel_column_info(column: str | None = None,
                          workbook_name: str | None = None,
                          sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Describe columns: data type, blank count, unique count and sample values.

    PARAMETERS:
        column: One column name. Omit to describe all columns.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Sheet to inspect. Uses the active sheet if omitted.

    RETURNS:
        One entry per column: name, letter, type (numeric, text, date, mixed,
        empty), non_blank_count, blank_count, unique_count, samples, and
        min/max for numeric columns.

    NOTES:
        - Read-only. The first row is treated as the header.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    names = list(df.columns)

    if column is not None and str(column).strip():
        wanted = str(column).strip().lower()
        selected = {i for i, n in enumerate(names) if n.strip().lower() == wanted}
        if not selected:
            raise ExcelToolError(f"Column '{column}' not found. Available: {names}")
    else:
        selected = set(range(min(len(names), MAX_COLUMNS_INFO)))

    info = []
    for idx, name in enumerate(names):
        if idx not in selected:
            continue
        series = df.iloc[:, idx]
        non_blank = series.dropna()
        count = int(len(non_blank))

        extra = {}
        if count == 0:
            kind = "empty"
        elif all(isinstance(v, (datetime, date)) for v in non_blank):
            kind = "date"
        else:
            numeric = pd.to_numeric(non_blank, errors="coerce")
            valid = int(numeric.notna().sum())
            if valid == count:
                kind = "numeric"
                extra = {"minimum": float(numeric.min()), "maximum": float(numeric.max())}
            elif valid == 0:
                kind = "text"
            else:
                kind = "mixed"

        samples = [_text(v)[:SAMPLE_CHARS]
                   for v in non_blank.drop_duplicates().head(SAMPLE_COUNT)]
        info.append({
            "column": name,
            "letter": _col_letter(used.column + idx),
            "type": kind,
            "non_blank_count": count,
            "blank_count": int(len(series) - count),
            "unique_count": int(non_blank.nunique()),
            "samples": samples,
            **extra,
        })

    return {"success": True, "sheet": sheet.name, "total_rows": len(df),
            "total_columns": len(names), "columns_returned": len(info),
            "columns": info}


# ---------------- 5. SEARCH A VALUE ----------------

@safe_tool
def search_excel_value(value: str | float, match_case: bool = False,
                       whole_cell: bool = False, max_results: int = 20,
                       workbook_name: str | None = None,
                       sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Find which cells contain a value (like Ctrl+F). Returns cell addresses.

    PARAMETERS:
        value: Text or number to search for.
        match_case: True = case-sensitive search.
        whole_cell: True = the whole cell must equal the value. False = partial match.
        max_results: Maximum number of addresses to return, 1 to 100.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        sheet_name: Sheet to search. Uses the active sheet if omitted.

    RETURNS:
        total_matches and a list of {cell, value}.

    NOTES:
        - Read-only. Does not change anything (use find_replace to replace).
        - Searches the used range of one sheet.
    """
    needle = _text(value).strip()
    if not needle:
        raise ExcelToolError("value cannot be empty.")
    _check_limit(max_results, "max_results", MAX_SEARCH_RESULTS)
    if not match_case:
        needle = needle.lower()

    sheet = get_sheet(workbook_name, sheet_name)
    used = sheet.used_range
    if used.count == 1 and used.value is None:
        return {"success": True, "sheet": sheet.name, "total_matches": 0,
                "returned": 0, "matches": []}

    grid = used.options(ndim=2).value
    matches = []
    total = 0
    for i, row in enumerate(grid):
        for j, cell in enumerate(row):
            if cell is None:
                continue
            text = _text(cell)
            hay = text if match_case else text.lower()
            hit = hay.strip() == needle if whole_cell else needle in hay
            if not hit:
                continue
            total += 1
            if len(matches) < max_results:
                address = f"{_col_letter(used.column + j)}{used.row + i}"
                matches.append({"cell": address, "value": text[:60]})

    return {"success": True, "sheet": sheet.name, "total_matches": total,
            "returned": len(matches), "matches": matches}