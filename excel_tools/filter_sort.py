import sys

import pandas as pd

from excel_tools.workbook_manager import (
    ExcelToolError,
    backup_workbook,
    build_condition_mask,
    get_sheet,
    get_workbook,
    load_dataframe,
    safe_tool,
    write_dataframe_to_sheet,
)

IS_WINDOWS = sys.platform.startswith("win")
FILTER_PREFIX = "Filtered"
BACKUP_NOTE = "Backup is a copy of the last saved file. Unsaved edits are not in it."
CONDITIONS = ("equals, not_equals, greater_than, less_than, greater_or_equal, "
              "less_or_equal, contains, is_empty, between")


# ---------- helpers ----------

def _blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _jsonable(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def _resolve_column(df, name):
    """Case-insensitive column lookup."""
    match = next((c for c in df.columns if str(c).strip().lower() == str(name).strip().lower()), None)
    if match is None:
        raise ExcelToolError(f"Column '{name}' not found. Available: {list(df.columns)}")
    return match


def _number(value, label: str) -> float:
    if value is None:
        raise ExcelToolError(f"{label} is required for 'between'.")
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ExcelToolError(f"{label} must be a number.")


def _crit_text(value) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


def _mask(series, condition, value, value2=None):
    cond = str(condition).lower().strip()
    if cond == "between":
        low, high = _number(value, "value"), _number(value2, "value2")
        if low > high:
            raise ExcelToolError("value must be less than or equal to value2.")
        nums = pd.to_numeric(series, errors="coerce")
        return ((nums >= low) & (nums <= high)).fillna(False).astype(bool)
    return build_condition_mask(series, cond, value)


def _output_name(name) -> str:
    """Result sheets always start with 'Filtered' so clear_excel_filter can remove them safely."""
    name = str(name).strip() or "Filtered_Data"
    if not name.startswith(FILTER_PREFIX):
        name = f"{FILTER_PREFIX}_{name}"
    return name[:31]


def _write_result(sheet, df, mask, output_sheet, criteria: str) -> dict:
    out = _output_name(output_sheet)
    if out == sheet.name:
        raise ExcelToolError("output_sheet cannot be the same as the source sheet.")
    result = df[mask]
    if result.empty:
        return {"success": True, "matched_rows": 0, "total_rows": len(df), "criteria": criteria,
                "message": "No matching rows found. No sheet was created."}
    existed = out in [s.name for s in sheet.book.sheets]
    write_dataframe_to_sheet(sheet.book, out, result, FILTER_PREFIX)
    res = {"success": True, "matched_rows": len(result), "total_rows": len(df),
           "output_sheet": out, "criteria": criteria}
    if existed:
        res["note"] = f"Existing sheet '{out}' was replaced with the new result."
    return res


def _autofilter_criteria(condition, value, value2):
    """Return (criteria1, criteria2) in Excel AutoFilter syntax."""
    cond = str(condition).lower().strip()
    ops = {"equals": "=", "equal": "=", "not_equals": "<>", "greater_than": ">",
           "less_than": "<", "greater_or_equal": ">=", "less_or_equal": "<="}
    if cond == "is_empty":
        return "=", None
    if cond == "between":
        low, high = _number(value, "value"), _number(value2, "value2")
        if low > high:
            raise ExcelToolError("value must be less than or equal to value2.")
        return f">={_crit_text(low)}", f"<={_crit_text(high)}"
    if value is None or str(value) == "":
        raise ExcelToolError("value is required for this condition.")
    if cond == "contains":
        return f"*{value}*", None
    if cond in ops:
        return f"{ops[cond]}{_crit_text(value)}", None
    raise ExcelToolError(f"Invalid condition '{condition}'. Use: {CONDITIONS}")


# ---------- tools ----------

@safe_tool
def filter_excel_data(column: str, condition: str, value: str | float | None = None,
                      value2: float | None = None, output_sheet: str = "Filtered_Data",
                      workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Filter rows by one condition and copy the matches to a separate sheet. Original data is untouched.

    PARAMETERS:
        column: Column header name (not case-sensitive).
        condition: equals, not_equals, greater_than, less_than, greater_or_equal, less_or_equal, contains, is_empty, between.
        value: Value to compare (not needed for is_empty).
        value2: Upper limit, only for between.
        output_sheet: Result sheet name (always starts with Filtered).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, matched_rows, total_rows, output_sheet, criteria.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    df, _ = load_dataframe(sheet)
    col = _resolve_column(df, column)
    mask = _mask(df[col], condition, value, value2)
    criteria = f"{col} {condition} {value}" + (f" and {value2}" if value2 is not None else "")
    return _write_result(sheet, df, mask, output_sheet, criteria)


@safe_tool
def filter_excel_multi(column1: str, condition1: str, value1: str | float | None = None,
                       column2: str | None = None, condition2: str | None = None,
                       value2: str | float | None = None, logic: str = "and",
                       output_sheet: str = "Filtered_Data", workbook_name: str | None = None,
                       sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Filter rows with two conditions joined by AND / OR and copy the matches to a separate sheet.

    PARAMETERS:
        column1: First column header name.
        condition1: equals, not_equals, greater_than, less_than, greater_or_equal, less_or_equal, contains, is_empty.
        value1: Value for the first condition.
        column2: Second column header name (optional).
        condition2: Second condition (same list as condition1).
        value2: Value for the second condition.
        logic: and (both must match) or or (either matches).
        output_sheet: Result sheet name (always starts with Filtered).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, matched_rows, total_rows, output_sheet, criteria.
    """
    logic = str(logic).lower().strip()
    if logic not in {"and", "or"}:
        raise ExcelToolError("logic must be 'and' or 'or'.")
    if (column2 is None) != (condition2 is None):
        raise ExcelToolError("column2 and condition2 must be given together.")
    for cond in (condition1, condition2):
        if cond is not None and str(cond).lower().strip() == "between":
            raise ExcelToolError("'between' is not supported here. Use filter_excel_data.")

    sheet = get_sheet(workbook_name, sheet_name)
    df, _ = load_dataframe(sheet)
    col1 = _resolve_column(df, column1)
    mask = _mask(df[col1], condition1, value1).fillna(False).astype(bool)
    criteria = f"{col1} {condition1} {value1}"
    if column2 is not None:
        col2 = _resolve_column(df, column2)
        mask2 = _mask(df[col2], condition2, value2).fillna(False).astype(bool)
        mask = (mask & mask2) if logic == "and" else (mask | mask2)
        criteria += f" {logic.upper()} {col2} {condition2} {value2}"
    return _write_result(sheet, df, mask, output_sheet, criteria)


@safe_tool
def apply_excel_autofilter(column: str | None = None, condition: str | None = None,
                           value: str | float | None = None, value2: float | None = None,
                           workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Turn on the filter dropdown arrows in place and optionally filter one column. Rows are hidden, not deleted. Call again with another column to add more filters.

    PARAMETERS:
        column: Column header name to filter (optional, without it only the arrows are added).
        condition: equals, not_equals, greater_than, less_than, greater_or_equal, less_or_equal, contains, is_empty, between.
        value: Value to compare (not needed for is_empty).
        value2: Upper limit, only for between.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sheet, column, condition, message.
    """
    if (column is None) != (condition is None):
        raise ExcelToolError("column and condition must be given together.")
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        raise ExcelToolError("No data found.")

    field = c1 = c2 = col = None
    if column is not None:
        col = _resolve_column(df, column)
        field = list(df.columns).index(col) + 1
        c1, c2 = _autofilter_criteria(condition, value, value2)
        if not IS_WINDOWS and c2 is not None:
            raise ExcelToolError("'between' in-place filter is supported only with Windows Excel. "
                                 "Use filter_excel_data instead.")

    try:
        if IS_WINDOWS:
            api = sheet.api
            if not api.AutoFilterMode:
                used.api.AutoFilter()
            if field is not None:
                kw = {"Field": field, "Criteria1": c1}
                if c2 is not None:
                    kw["Operator"] = 1  # xlAnd
                    kw["Criteria2"] = c2
                api.AutoFilter.Range.AutoFilter(**kw)
        else:
            # Mac (appscript), best effort and untested
            try:
                already_on = bool(sheet.api.autofilter_mode.get())
            except Exception:
                already_on = False
            if field is None:
                if not already_on:
                    used.api.autofilter()
            else:
                used.api.autofilter(field=field, criteria1=c1)
    except Exception as exc:
        raise ExcelToolError("In-place filter failed on this platform. Use filter_excel_data instead. "
                             f"Details: {exc}")

    return {"success": True, "sheet": sheet.name, "column": col, "condition": condition,
            "message": ("Filter arrows added." if col is None else f"Filter applied on '{col}'.")
            + " Rows are hidden, not deleted."}


@safe_tool
def clear_excel_filter(output_sheet: str = "Filtered_Data", remove_autofilter: bool = False,
                       workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Remove the filter result sheet made by filter_excel_data, or (remove_autofilter=true) remove the in-place filter and show all rows.

    PARAMETERS:
        output_sheet: Result sheet to delete (must start with Filtered).
        remove_autofilter: true to clear the in-place filter on the data sheet instead (Windows Excel only).
        workbook_name: Workbook name (optional).
        sheet_name: Data sheet name for remove_autofilter (optional).

    RETURNS:
        success, removed, sheet, message.
    """
    if remove_autofilter:
        if not IS_WINDOWS:
            raise ExcelToolError("Removing an in-place filter is supported only with Windows Excel (COM).")
        sheet = get_sheet(workbook_name, sheet_name)
        api = sheet.api
        if api.FilterMode:
            api.ShowAllData()
        api.AutoFilterMode = False
        return {"success": True, "removed": True, "sheet": sheet.name,
                "message": "In-place filter removed. All rows are visible."}

    out = str(output_sheet).strip()
    if not out.startswith(FILTER_PREFIX):
        raise ExcelToolError(f"Safety: only sheets starting with '{FILTER_PREFIX}' can be deleted.")
    book = get_workbook(workbook_name)
    if out not in [s.name for s in book.sheets]:
        return {"success": True, "removed": False,
                "message": f"Sheet '{out}' not found. For an in-place filter use remove_autofilter=true."}
    if len(book.sheets) == 1:
        raise ExcelToolError("This is the only sheet. It cannot be deleted.")
    book.app.display_alerts = False
    try:
        book.sheets[out].delete()
    finally:
        book.app.display_alerts = True
    return {"success": True, "removed": True, "sheet": out, "message": f"Sheet '{out}' deleted."}


def _sort_key(series):
    """Safe sort key: numbers stay numeric, mixed or text columns sort as lowercase text."""
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_datetime64_any_dtype(series):
        return series
    nums = pd.to_numeric(series, errors="coerce")
    if nums.notna().sum() == series.notna().sum():
        return nums
    return series.map(lambda v: None if _blank(v) else str(v).lower())


@safe_tool
def sort_excel_data(columns: list[str] | str, ascending: bool = True, confirm: bool = False,
                    workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Sort the data rows by one or more columns (header stays on top). First call shows a preview, second call with confirm=true sorts. Formulas in data rows become values.

    PARAMETERS:
        columns: Column name, or comma-separated names, example Age or Gender,Age.
        ascending: true = A to Z / smallest first, false = Z to A / largest first.
        confirm: Must be true to actually sort.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sorted, columns, rows, preview, backup.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        return {"success": True, "sorted": False, "message": "No data found."}

    if isinstance(columns, str):
        try:
            exact = _resolve_column(df, columns)  # a single column name that contains a comma
            names = [exact]
        except ExcelToolError:
            names = [p.strip() for p in columns.split(",") if p.strip()]
    else:
        names = [str(c).strip() for c in columns]
    if not names:
        raise ExcelToolError("Give at least one column.")
    resolved = [_resolve_column(df, n) for n in names]

    last_row = used.last_cell.row
    last_col = used.last_cell.column
    if len(df) != last_row - used.row:
        raise ExcelToolError("The data has blank rows inside it. Remove blank rows first, then sort.")

    grid = sheet.range((used.row + 1, used.column), (last_row, last_col)).options(ndim=2).formula
    formula_cells = sum(1 for row in grid for c in row if isinstance(c, str) and c.startswith("="))

    keys = pd.DataFrame({f"k{i}": _sort_key(df[c]) for i, c in enumerate(resolved)}).reset_index(drop=True)
    order = keys.sort_values(by=list(keys.columns), ascending=bool(ascending), kind="stable",
                             na_position="last").index
    sorted_df = df.reset_index(drop=True).iloc[list(order)]

    warning = None
    if formula_cells:
        warning = (f"{formula_cells} formula cell(s) in the data will become plain values after sorting.")

    if not confirm:
        head = sorted_df.head(5)
        records = [{k: _jsonable(v) for k, v in row.items()}
                   for row in head.astype(object).where(head.notna(), None).to_dict(orient="records")]
        out = {"success": True, "sorted": False, "needs_confirm": True, "columns": resolved,
               "rows": len(df), "preview": records,
               "message": f"{len(df)} rows will be sorted. Confirm to apply."}
        if warning:
            out["warning"] = warning
        return out

    backup = backup_workbook(sheet.book)
    native_ok = False
    if len(resolved) <= 3:
        try:
            from appscript import k as kw
            col_list = list(df.columns)
            rng = sheet.range((used.row, used.column), (last_row, last_col))
            order_const = kw.ascending if ascending else kw.descending
            sort_args = {}
            for i, c in enumerate(resolved):
                key_cell = sheet.range((used.row, used.column + col_list.index(c)))
                sort_args[f"key{i + 1}"] = key_cell.api
                sort_args[f"order{i + 1}"] = order_const
            rng.api.sort(header=kw.yes, **sort_args)
            first = sheet.range(
                (used.row + 1, used.column + col_list.index(resolved[0]))).value
            expected = sorted_df[resolved[0]].iloc[0]
            native_ok = (first == expected) or (pd.isna(expected) and first is None)
        except Exception:
            native_ok = False
    if not native_ok:
        body = sorted_df.astype(object).where(sorted_df.notna(), None)
        sheet.range((used.row + 1, used.column)).value = body.values.tolist()
    out = {"success": True, "sorted": True, "columns": resolved, "rows": len(df),
           "ascending": bool(ascending), "backup": backup, "note": BACKUP_NOTE}
    if warning:
        out["warning"] = warning
    return out