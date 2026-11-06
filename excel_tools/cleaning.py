import re

import pandas as pd

from excel_tools.workbook_manager import (
    ExcelToolError,
    backup_workbook,
    contiguous_blocks,
    delete_rows,
    df_to_records,
    get_sheet,
    load_dataframe,
    safe_tool,
)

BACKUP_NOTE = "Backup is a copy of the last saved file. Unsaved edits are not in it."


# ---------- helpers ----------

def _is_blank(value):
    """Check whether a cell value is blank."""
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def _is_plain(value) -> bool:
    """True for text and numbers (bool, dates and others are skipped)."""
    return isinstance(value, (str, int, float)) and not isinstance(value, bool)


def _text(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _coerce(text, original):
    """If the original cell was a number and the new text looks like a number, keep it numeric."""
    if isinstance(original, str):
        return text
    try:
        number = float(str(text).strip())
    except (TypeError, ValueError):
        return text
    return int(number) if number.is_integer() else number


def _jsonable(value):
    return value if isinstance(value, (str, int, float)) else str(value)


def _replace_ci(text, find_text, replace_text):
    return re.sub(re.escape(find_text), lambda m: replace_text, text, flags=re.IGNORECASE)


def _pick_columns(df, columns):
    if columns is None:
        return list(df.columns)
    if not columns:
        raise ExcelToolError("columns cannot be empty.")
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ExcelToolError(f"Columns not found: {missing}. Available: {list(df.columns)}")
    return list(columns)


def _formula_mask(sheet, used, nrows, ncols):
    """One bulk read: True where a data cell holds a formula."""
    block = sheet.range((used.row + 1, used.column),
                        (used.row + nrows, used.column + ncols - 1))
    grid = block.options(ndim=2).formula
    return [[isinstance(c, str) and c.startswith("=") for c in row] for row in grid]


def _write_cells(sheet, updates):
    """updates: {(excel_row, excel_col): value}. Writes contiguous runs per column (few COM calls)."""
    by_col = {}
    for (row, col), value in updates.items():
        by_col.setdefault(col, {})[row] = value
    for col, rows in by_col.items():
        for start, end in contiguous_blocks(sorted(rows)):
            sheet.range((start, col), (end, col)).value = [[rows[r]] for r in range(start, end + 1)]


def _remove_flagged_rows(sheet, df, used, mask, confirm=False, label="rows"):
    """Preview or delete rows selected by a boolean mask."""
    count = int(mask.sum())
    if count == 0:
        return {"success": True, "deleted": False, "found": 0, "message": f"No {label} found."}

    rows = [used.row + 1 + int(i) for i in mask[mask].index]
    if not confirm:
        return {"success": True, "deleted": False, "found": count, "excel_rows": rows[:50],
                "preview": df_to_records(df[mask], 5),
                "message": f"Found {count} {label}. Confirm to delete them."}

    backup = backup_workbook(sheet.book)
    delete_rows(sheet, rows)
    return {"success": True, "deleted": True, "removed": count, "backup": backup,
            "note": "Formulas that pointed to deleted rows may show #REF!. " + BACKUP_NOTE}


# ---------- tools ----------

@safe_tool
def remove_excel_duplicates(columns: list[str] | None = None, keep: str = "first",
                            confirm: bool = False, workbook_name: str | None = None,
                            sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Find and remove duplicate rows. First call shows a preview, second call with confirm=true deletes.

    PARAMETERS:
        columns: Column names that define a duplicate (optional, default all columns).
        keep: first or last, which copy to keep.
        confirm: Must be true to actually delete.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, deleted, found, removed, preview, backup.
    """
    keep = str(keep).lower().strip()
    if keep not in {"first", "last"}:
        raise ExcelToolError("keep must be 'first' or 'last'.")
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        return {"success": True, "deleted": False, "found": 0, "message": "No data found."}
    if columns is not None:
        _pick_columns(df, columns)
    mask = df.duplicated(subset=columns, keep=keep)
    return _remove_flagged_rows(sheet, df, used, mask, confirm, "duplicate rows")


@safe_tool
def remove_excel_blank_rows(confirm: bool = False, workbook_name: str | None = None,
                            sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Find and remove completely blank data rows. First call shows a preview, second call with confirm=true deletes.

    PARAMETERS:
        confirm: Must be true to actually delete.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, deleted, found, removed, preview, backup.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        return {"success": True, "deleted": False, "found": 0, "message": "No data found."}
    mask = df.apply(lambda row: all(_is_blank(v) for v in row), axis=1)
    return _remove_flagged_rows(sheet, df, used, mask, confirm, "blank rows")


@safe_tool
def find_replace_excel_data(find_value: str, replace_value: str = "", columns: list[str] | None = None,
                            match_case: bool = False, whole_cell: bool = False,
                            confirm: bool = False, workbook_name: str | None = None,
                            sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Find and replace values in sheet data. First call shows a preview, second call with confirm=true replaces. Formula cells are skipped.

    PARAMETERS:
        find_value: Text or number to find.
        replace_value: Replacement (empty string clears the match).
        columns: Column names to search (optional, default all columns).
        match_case: true for case-sensitive search.
        whole_cell: true to match only when the whole cell equals find_value.
        confirm: Must be true to actually replace.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, found, replaced, skipped_formulas, preview, backup.
    """
    if find_value is None or str(find_value) == "":
        raise ExcelToolError("find_value cannot be empty.")
    replace_text = "" if replace_value is None else str(replace_value)
    search = str(find_value)

    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        return {"success": True, "found": 0, "replaced": 0, "message": "No data found."}
    selected = _pick_columns(df, columns)
    all_cols = list(df.columns)
    formulas = _formula_mask(sheet, used, len(df), len(all_cols))

    matches = []
    skipped = 0
    for column in selected:
        ci = all_cols.index(column)
        for ri, value in enumerate(df[column].tolist()):
            if _is_blank(value) or not _is_plain(value):
                continue
            text = _text(value)
            source, target = (text, search) if match_case else (text.lower(), search.lower())
            if whole_cell:
                hit = source == target
                new_text = replace_text
            else:
                hit = target in source
                new_text = (text.replace(search, replace_text) if match_case
                            else _replace_ci(text, search, replace_text))
            if not hit:
                continue
            if formulas[ri][ci]:
                skipped += 1
                continue
            matches.append({"row": used.row + 1 + ri, "col": used.column + ci, "column": column,
                            "old_value": _jsonable(value), "new_value": _coerce(new_text, value)})

    if not matches:
        return {"success": True, "found": 0, "replaced": 0, "skipped_formulas": skipped,
                "message": "No matching values found."}

    if not confirm:
        preview = [{"row": m["row"], "column": m["column"], "old_value": m["old_value"],
                    "new_value": m["new_value"]} for m in matches[:50]]
        return {"success": True, "found": len(matches), "replaced": 0, "skipped_formulas": skipped,
                "preview": preview, "message": "Preview only. Confirm to replace."}

    backup = backup_workbook(sheet.book)
    _write_cells(sheet, {(m["row"], m["col"]): m["new_value"] for m in matches})
    return {"success": True, "found": len(matches), "replaced": len(matches),
            "skipped_formulas": skipped, "backup": backup, "note": BACKUP_NOTE}


@safe_tool
def fill_missing_excel_values(fill_value=None, columns: list[str] | None = None, method: str = "value",
                              confirm: bool = False, workbook_name: str | None = None,
                              sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Fill blank cells in selected columns. First call shows a preview, second call with confirm=true fills.

    PARAMETERS:
        fill_value: Value to use when method is value.
        columns: Column names (optional, default all columns).
        method: value (use fill_value), forward (previous non-blank value) or backward (next non-blank value).
        confirm: Must be true to actually fill.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, filled, updated, preview, backup.
    """
    method = str(method).lower().strip()
    if method not in {"value", "forward", "backward"}:
        raise ExcelToolError("method must be value, forward or backward.")
    if method == "value" and fill_value is None:
        raise ExcelToolError("fill_value is required when method is value.")

    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        return {"success": True, "filled": 0, "updated": False, "message": "No data found."}
    selected = _pick_columns(df, columns)
    all_cols = list(df.columns)
    formulas = _formula_mask(sheet, used, len(df), len(all_cols))
    constant = _coerce(fill_value, 0) if (method == "value" and isinstance(fill_value, str)) else fill_value

    updates = []
    for column in selected:
        ci = all_cols.index(column)
        values = df[column].tolist()
        blanks = [_is_blank(v) for v in values]
        order = range(len(values)) if method != "backward" else range(len(values) - 1, -1, -1)
        last = None
        for ri in order:
            if not blanks[ri]:
                last = values[ri]
                continue
            new = constant if method == "value" else last
            if new is None or formulas[ri][ci]:
                continue
            updates.append({"row": used.row + 1 + ri, "col": used.column + ci,
                            "column": column, "value": new})

    if not updates:
        return {"success": True, "filled": 0, "updated": False, "message": "No missing values to fill."}

    if not confirm:
        preview = [{"row": u["row"], "column": u["column"], "value": _jsonable(u["value"])}
                   for u in updates[:50]]
        return {"success": True, "filled": len(updates), "updated": False, "preview": preview,
                "message": "Preview only. Confirm to fill."}

    backup = backup_workbook(sheet.book)
    _write_cells(sheet, {(u["row"], u["col"]): u["value"] for u in updates})
    return {"success": True, "filled": len(updates), "updated": True, "backup": backup,
            "note": BACKUP_NOTE}


TEXT_OPS = {"trim", "upper", "lower", "proper", "remove_nonprinting"}
TEXT_OP_ALIASES = {"uppercase": "upper", "lowercase": "lower", "title": "proper",
                   "titlecase": "proper", "title_case": "proper", "clean": "remove_nonprinting",
                   "nonprinting": "remove_nonprinting", "strip": "trim"}
WORD_RE = re.compile(r"[^\W\d_]+(?:'[^\W\d_]+)*")


def _clean_text(text: str, ops: set) -> str:
    if "remove_nonprinting" in ops:
        text = re.sub(r"[\x00-\x1f\x7f]", "", text.replace("\xa0", " "))
    if "trim" in ops:
        text = re.sub(r" +", " ", text.replace("\xa0", " ")).strip()
    if "upper" in ops:
        text = text.upper()
    elif "lower" in ops:
        text = text.lower()
    elif "proper" in ops:
        text = WORD_RE.sub(lambda m: m.group(0).capitalize(), text)
    return text


@safe_tool
def clean_excel_text(columns: list[str] | None = None, operations: str = "trim",
                     workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Clean text cells: trim extra spaces, change case, remove non-printing characters. Applies directly with a backup. Numbers and formulas are not touched.

    PARAMETERS:
        columns: Column names (optional, default all columns).
        operations: Comma list of trim, upper, lower, proper, remove_nonprinting. Only one of upper, lower, proper.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, changed, columns, operations, skipped_formulas, backup.
    """
    raw = [p for p in re.split(r"[,\s]+", str(operations).lower().strip()) if p]
    ops = {TEXT_OP_ALIASES.get(p, p) for p in raw}
    bad = sorted(ops - TEXT_OPS)
    if not ops or bad:
        raise ExcelToolError(f"Invalid operations {bad}. Use: {sorted(TEXT_OPS)}")
    if len(ops & {"upper", "lower", "proper"}) > 1:
        raise ExcelToolError("Use only one of upper, lower or proper.")

    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        return {"success": True, "changed": 0, "message": "No data found."}
    selected = _pick_columns(df, columns)
    all_cols = list(df.columns)
    formulas = _formula_mask(sheet, used, len(df), len(all_cols))

    updates = {}
    skipped = 0
    for column in selected:
        ci = all_cols.index(column)
        for ri, value in enumerate(df[column].tolist()):
            if not isinstance(value, str):
                continue
            new = _clean_text(value, ops)
            if new == value:
                continue
            if formulas[ri][ci]:
                skipped += 1
                continue
            updates[(used.row + 1 + ri, used.column + ci)] = new

    if not updates:
        return {"success": True, "changed": 0, "skipped_formulas": skipped,
                "message": "Nothing needed cleaning."}

    backup = backup_workbook(sheet.book)
    _write_cells(sheet, updates)
    return {"success": True, "changed": len(updates), "columns": selected,
            "operations": sorted(ops), "skipped_formulas": skipped, "backup": backup,
            "note": BACKUP_NOTE}