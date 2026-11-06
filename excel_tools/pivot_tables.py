import sys

import pandas as pd

from excel_tools.workbook_manager import (
    ExcelToolError,
    backup_workbook,
    get_sheet,
    get_workbook,
    load_dataframe,
    safe_tool,
    write_dataframe_to_sheet,
)

IS_WINDOWS = sys.platform.startswith("win")
PIVOT_PREFIX = "Pivot"
AGG_ALIASES = {"sum": "sum", "total": "sum", "average": "mean", "avg": "mean", "mean": "mean",
               "count": "count", "min": "min", "minimum": "min", "max": "max", "maximum": "max"}
AGG_LABELS = {"sum": "Sum", "mean": "Average", "count": "Count", "min": "Min", "max": "Max"}
AGG_COM = {"sum": -4157, "mean": -4106, "count": -4112, "min": -4139, "max": -4136}
SORTS = ("descending", "ascending", "none")
NATIVE_HINT = "Native PivotTables are supported only with Windows Excel (COM). Use create_excel_pivot_summary instead."


# ---------- helpers ----------

def _need_com() -> None:
    if not IS_WINDOWS:
        raise ExcelToolError(NATIVE_HINT)


def _find_col(df, name, label: str = "Column"):
    """Case-insensitive column lookup."""
    key = str(name).strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == key:
            return col
    raise ExcelToolError(f"{label} '{name}' not found. Available: {list(df.columns)}")


def _agg(aggregation: str) -> str:
    key = str(aggregation).lower().strip()
    if key not in AGG_ALIASES:
        raise ExcelToolError("aggregation must be sum, average, count, min or max.")
    return AGG_ALIASES[key]


def _output_name(name) -> str:
    """Result sheets always start with 'Pivot' so delete_excel_pivot can remove them safely."""
    name = str(name).strip() or "Pivot_Summary"
    if not name.startswith(PIVOT_PREFIX):
        name = f"{PIVOT_PREFIX}_{name}"
    return name[:31]


def _unique_name(book, base: str) -> str:
    existing = {s.name for s in book.sheets}
    if base not in existing:
        return base
    for i in range(2, 100):
        candidate = f"{base[:28]}_{i}"
        if candidate not in existing:
            return candidate
    raise ExcelToolError("Could not find a free sheet name. Choose another output_sheet.")


def _native_pivots(book, sheet_name: str | None = None):
    """Yield (sheet, pivot) for every native PivotTable (Windows only)."""
    sheets = [book.sheets[sheet_name]] if sheet_name else list(book.sheets)
    for sh in sheets:
        try:
            tables = sh.api.PivotTables()
            count = int(tables.Count)
        except Exception:
            continue
        for i in range(1, count + 1):
            yield sh, tables.Item(i)


# ---------- tools ----------

@safe_tool
def create_excel_pivot_summary(group_by: str, value_column: str, aggregation: str = "sum",
                               column_group: str | None = None, sort: str = "descending",
                               output_sheet: str = "Pivot_Summary",
                               workbook_name: str | None = None,
                               sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Summarize data (sum, average, count, min, max) grouped by one column, or by two columns as a cross-table. Writes a static table to a new sheet (not a native PivotTable). Original data is untouched.

    PARAMETERS:
        group_by: Column for the rows, example Region (not case-sensitive).
        value_column: Column to calculate, example Sales.
        aggregation: sum, average, count, min or max.
        column_group: Second column for a cross-table, example Product (optional).
        sort: descending, ascending or none (sorts by the result values or totals).
        output_sheet: Result sheet name (always starts with Pivot).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, output_sheet, groups, top_rows, note.
    """
    agg = _agg(aggregation)
    order = str(sort).lower().strip()
    if order not in SORTS:
        raise ExcelToolError("sort must be descending, ascending or none.")

    sheet = get_sheet(workbook_name, sheet_name)
    out = _output_name(output_sheet)
    if out == sheet.name:
        raise ExcelToolError("output_sheet cannot be the same as the source sheet.")

    df, _ = load_dataframe(sheet)
    if df.empty:
        raise ExcelToolError("No data found.")
    g = _find_col(df, group_by, "group_by column")
    v = _find_col(df, value_column, "value_column")
    c = _find_col(df, column_group, "column_group") if column_group is not None else None
    if g == v:
        raise ExcelToolError("group_by and value_column must be different columns.")
    if c is not None and c in (g, v):
        raise ExcelToolError("column_group must be different from group_by and value_column.")

    values = df[v]
    if agg != "count":
        values = pd.to_numeric(values, errors="coerce")
        if values.notna().sum() == 0:
            raise ExcelToolError(f"'{v}' is not a numeric column.")

    label = f"{AGG_LABELS[agg]} of {v}"
    asc = order == "ascending"

    if c is None:
        temp = pd.DataFrame({g: df[g], v: values})
        summary = temp.groupby(g, dropna=True)[v].agg(agg).reset_index()
        summary.columns = [str(g), label]
        if order != "none":
            summary = summary.sort_values(label, ascending=asc)
        groups = len(summary)
    else:
        temp = pd.DataFrame({g: df[g], c: df[c], v: values})
        table = pd.pivot_table(temp, index=g, columns=c, values=v, aggfunc=agg)
        if agg in {"sum", "count"}:
            table = table.fillna(0)
            table["Total"] = table.sum(axis=1)
            if order != "none":
                table = table.sort_values("Total", ascending=asc)
        table.columns = [str(x) for x in table.columns]
        summary = table.reset_index()
        summary.columns = [str(x) for x in summary.columns]
        groups = len(summary)

    existed = out in [s.name for s in sheet.book.sheets]
    write_dataframe_to_sheet(sheet.book, out, summary, PIVOT_PREFIX)

    top = summary.head(5).astype(object).where(summary.head(5).notna(), None)
    result = {"success": True, "output_sheet": out, "groups": int(groups),
              "top_rows": top.to_dict("records"),
              "note": "Static summary. It does not update when the data changes; run it again to refresh."}
    if existed:
        result["note"] = f"Existing sheet '{out}' was replaced with the new result. " + result["note"]
    return result


@safe_tool
def create_excel_native_pivot(row_field: str, value_field: str, aggregation: str = "sum",
                              column_field: str | None = None, output_sheet: str = "Pivot_Native",
                              workbook_name: str | None = None,
                              sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Create a real Excel PivotTable on a new sheet (Windows Excel only).

    PARAMETERS:
        row_field: Column for the rows, example Region (not case-sensitive).
        value_field: Column to calculate, example Sales.
        aggregation: sum, average, count, min or max.
        column_field: Column for the columns, example Product (optional).
        output_sheet: New sheet name (always starts with Pivot).
        workbook_name: Workbook name (optional).
        sheet_name: Source sheet name (optional).

    RETURNS:
        success, pivot_name, output_sheet, source_range, row_field, value_field.
    """
    _need_com()
    agg = _agg(aggregation)
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if df.empty:
        raise ExcelToolError("No data found.")
    r = _find_col(df, row_field, "row_field")
    v = _find_col(df, value_field, "value_field")
    c = _find_col(df, column_field, "column_field") if column_field is not None else None
    if r == v or (c is not None and c in (r, v)):
        raise ExcelToolError("row_field, column_field and value_field must be different columns.")

    book = sheet.book
    out = _unique_name(book, _output_name(output_sheet))
    src = sheet.range((used.row, used.column), (used.last_cell.row, used.last_cell.column))

    new_sheet = book.sheets.add(name=out, after=book.sheets[-1])
    try:
        cache = book.api.PivotCaches().Create(SourceType=1, SourceData=src.api)
        pivot = cache.CreatePivotTable(TableDestination=new_sheet.range("A3").api, TableName=out)
        field = pivot.PivotFields(str(r))
        field.Orientation = 1
        field.Position = 1
        if c is not None:
            col_field = pivot.PivotFields(str(c))
            col_field.Orientation = 2
            col_field.Position = 1
        pivot.AddDataField(pivot.PivotFields(str(v)), f"{AGG_LABELS[agg]} of {v}", AGG_COM[agg])
    except Exception as exc:
        try:
            book.app.display_alerts = False
            new_sheet.delete()
        finally:
            book.app.display_alerts = True
        raise ExcelToolError(f"Could not create the PivotTable: {exc}")

    return {"success": True, "pivot_name": out, "output_sheet": out,
            "source_range": src.address.replace("$", ""), "row_field": str(r),
            "column_field": str(c) if c is not None else None,
            "value_field": str(v), "aggregation": agg}


@safe_tool
def list_excel_pivots(workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        List PivotTables (name, sheet, location, source) and summary sheets made by create_excel_pivot_summary. Read-only.

    PARAMETERS:
        workbook_name: Workbook name (optional).

    RETURNS:
        success, native_count, native_pivots, summary_sheets, note.
    """
    book = get_workbook(workbook_name)
    natives = []
    if IS_WINDOWS:
        for sh, pivot in _native_pivots(book):
            item = {"name": str(pivot.Name), "sheet": sh.name}
            try:
                item["location"] = pivot.TableRange2.Address.replace("$", "")
            except Exception:
                item["location"] = None
            try:
                item["source"] = str(pivot.SourceData)
            except Exception:
                item["source"] = None
            natives.append(item)

    summaries = [s.name for s in book.sheets if s.name.startswith(PIVOT_PREFIX)]
    out = {"success": True, "native_count": len(natives), "native_pivots": natives,
           "summary_sheets": summaries}
    if not IS_WINDOWS:
        out["note"] = "Native PivotTables can be listed only with Windows Excel. Summary sheets are shown."
    return out


@safe_tool
def modify_excel_pivot(pivot_name: str | None = None, refresh: bool = True,
                       workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Refresh native PivotTables so they show the latest data (Windows Excel only). Without pivot_name all PivotTables are refreshed. Summary sheets from create_excel_pivot_summary are static: run that tool again.

    PARAMETERS:
        pivot_name: PivotTable name (optional, see list_excel_pivots).
        refresh: true to refresh.
        workbook_name: Workbook name (optional).
        sheet_name: Only look on this sheet (optional).

    RETURNS:
        success, refreshed, pivots, message.
    """
    _need_com()
    book = get_workbook(workbook_name)
    found = list(_native_pivots(book, sheet_name))
    if not found:
        raise ExcelToolError("No native PivotTables found. Use list_excel_pivots to check.")
    if pivot_name is not None:
        key = str(pivot_name).strip().lower()
        matched = [(s, p) for s, p in found if str(p.Name).strip().lower() == key]
        if not matched:
            raise ExcelToolError(f"PivotTable '{pivot_name}' not found. Available: {[str(p.Name) for _, p in found]}")
        found = matched
    if not refresh:
        return {"success": True, "refreshed": False, "pivots": [str(p.Name) for _, p in found],
                "message": "No changes requested."}
    for _, pivot in found:
        pivot.RefreshTable()
    names = [str(p.Name) for _, p in found]
    return {"success": True, "refreshed": True, "pivots": names,
            "message": f"Refreshed {len(names)} PivotTable(s)."}


@safe_tool
def delete_excel_pivot(name: str, confirm: bool = False, workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Delete a summary sheet (name starts with Pivot) or a native PivotTable. First call shows a preview, second call with confirm=true deletes.

    PARAMETERS:
        name: Summary sheet name or PivotTable name (see list_excel_pivots).
        confirm: Must be true to actually delete.
        workbook_name: Workbook name (optional).

    RETURNS:
        success, deleted, name, kind, backup, message.
    """
    book = get_workbook(workbook_name)
    target = str(name).strip()
    sheet_match = next((s for s in book.sheets
                        if s.name.lower() == target.lower() and s.name.startswith(PIVOT_PREFIX)), None)

    pivot_match = None
    if sheet_match is None and IS_WINDOWS:
        pivot_match = next(((s, p) for s, p in _native_pivots(book)
                            if str(p.Name).strip().lower() == target.lower()), None)

    if sheet_match is None and pivot_match is None:
        raise ExcelToolError(
            f"'{name}' not found. Only sheets starting with '{PIVOT_PREFIX}' or native PivotTables can be deleted. "
            "Use list_excel_pivots.")

    kind = "summary_sheet" if sheet_match is not None else "native_pivot"
    real_name = sheet_match.name if sheet_match is not None else str(pivot_match[1].Name)

    if not confirm:
        return {"success": True, "deleted": False, "needs_confirm": True, "name": real_name,
                "kind": kind, "message": f"'{real_name}' will be deleted. Please confirm."}

    if sheet_match is not None and len(book.sheets) == 1:
        raise ExcelToolError("This is the only sheet. It cannot be deleted.")

    backup = backup_workbook(book)
    if sheet_match is not None:
        book.app.display_alerts = False
        try:
            sheet_match.delete()
        finally:
            book.app.display_alerts = True
    else:
        pivot_match[1].TableRange2.Clear()
    return {"success": True, "deleted": True, "name": real_name, "kind": kind, "backup": backup,
            "message": f"'{real_name}' deleted. Backup is a copy of the last saved file."}