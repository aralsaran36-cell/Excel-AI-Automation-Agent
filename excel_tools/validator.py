from excel_tools.workbook_manager import get_sheet, load_dataframe, safe_tool


@safe_tool
def validate_excel_data(workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """Check data quality: missing values, duplicate rows, blank rows, mixed types, unnamed columns."""
    sheet = get_sheet(workbook_name, sheet_name)
    df, _ = load_dataframe(sheet)
    issues = []

    unnamed = [c for c in df.columns if c.startswith("Unnamed_")]
    if unnamed:
        issues.append({"type": "unnamed_columns", "columns": unnamed})

    missing = {c: int(n) for c, n in df.isna().sum().items() if n > 0}
    if missing:
        issues.append({"type": "missing_values", "per_column": missing})

    blank_rows = int(df.isna().all(axis=1).sum())
    if blank_rows:
        issues.append({"type": "blank_rows", "count": blank_rows})

    duplicate_rows = int(df.duplicated().sum())
    if duplicate_rows:
        issues.append({"type": "duplicate_rows", "count": duplicate_rows})

    mixed = {}
    for col in df.columns:
        kinds = {type(v).__name__ for v in df[col].dropna()}
        if len(kinds) > 1:
            mixed[col] = sorted(kinds)
    if mixed:
        issues.append({"type": "mixed_types", "columns": mixed})

    return {"success": True, "rows": len(df), "columns": len(df.columns),
            "is_clean": not issues, "issues": issues}