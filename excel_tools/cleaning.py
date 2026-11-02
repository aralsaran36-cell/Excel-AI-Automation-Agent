from excel_tools.workbook_manager import (
    ExcelToolError, backup_workbook, delete_rows, df_to_records, get_sheet,
    load_dataframe, safe_tool,
)


def _remove_flagged_rows(sheet, df, used, mask, confirm: bool, label: str) -> dict:
    """Shared by duplicate and blank-row removal: preview first, delete only with confirm=True."""
    count = int(mask.sum())
    if count == 0:
        return {"success": True, "deleted": False, "found": 0, "message": f"{label} edhuvum illa."}
    rows = [used.row + 1 + int(i) for i in mask[mask].index]
    if not confirm:
        return {"success": True, "deleted": False, "found": count, "excel_rows": rows[:50],
                "preview": df_to_records(df[mask], 5),
                "message": f"{count} {label} kandupidichen. Delete pannanum-na confirm=True kudunga."}
    backup = backup_workbook(sheet.book)
    delete_rows(sheet, rows)
    return {"success": True, "deleted": True, "removed": count, "backup": backup}


@safe_tool
def remove_excel_duplicates(columns: list[str] | None = None, keep: str = "first",
                            confirm: bool = False, workbook_name: str | None = None,
                            sheet_name: str | None = None) -> dict:
    """Find duplicate rows (all columns, or only 'columns'). Deletes only when confirm=True."""
    if keep not in ("first", "last"):
        raise ExcelToolError("keep 'first' or 'last' aa irukkanum.")
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if columns:
        missing = [c for c in columns if c not in df.columns]
        if missing:
            raise ExcelToolError(f"Columns illa: {missing}. Available: {list(df.columns)}")
    mask = df.duplicated(subset=columns or None, keep=keep)
    return _remove_flagged_rows(sheet, df, used, mask, confirm, "duplicate rows")


@safe_tool
def remove_excel_blank_rows(confirm: bool = False, workbook_name: str | None = None,
                            sheet_name: str | None = None) -> dict:
    """Find fully blank rows inside the data. Deletes only when confirm=True."""
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    mask = df.isna().all(axis=1)
    return _remove_flagged_rows(sheet, df, used, mask, confirm, "blank rows")