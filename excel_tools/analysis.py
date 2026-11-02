from datetime import date, datetime

import pandas as pd

from excel_tools.workbook_manager import (
    ExcelToolError, get_sheet, load_dataframe, safe_tool,
)


def _is_date_like(series: pd.Series) -> bool:
    return pd.api.types.is_datetime64_any_dtype(series) or all(
        isinstance(v, (datetime, date)) for v in series
    )


@safe_tool
def analyze_excel_column(column: str, workbook_name: str | None = None,
                         sheet_name: str | None = None) -> dict:
    """Statistics for one column: numeric (sum, mean, median, min, max, std),
    date (earliest, latest) or text (unique, top values)."""
    sheet = get_sheet(workbook_name, sheet_name)
    df, _ = load_dataframe(sheet)
    if column not in df.columns:
        raise ExcelToolError(f"Column '{column}' illa. Available: {list(df.columns)}")
    series = df[column]
    info = {"success": True, "column": column, "total_rows": len(series),
            "missing": int(series.isna().sum())}
    non_null = series.dropna()
    if non_null.empty:
        return {**info, "type": "empty"}
    if _is_date_like(non_null):
        return {**info, "type": "date", "earliest": str(non_null.min()), "latest": str(non_null.max())}
    numeric = pd.to_numeric(non_null, errors="coerce")
    if numeric.notna().all():
        return {**info, "type": "numeric", "sum": float(numeric.sum()), "mean": float(numeric.mean()),
                "median": float(numeric.median()), "min": float(numeric.min()),
                "max": float(numeric.max()),
                "std": float(numeric.std()) if len(numeric) > 1 else 0.0}
    counts = non_null.astype(str).value_counts().head(5)
    return {**info, "type": "text", "unique": int(non_null.nunique()),
            "top_values": {k: int(v) for k, v in counts.items()}}







@safe_tool
def calculate_excel_total(column: str, year: int | None = None, group_by: str | None = None,
                          date_column: str | None = None, workbook_name: str | None = None,
                          sheet_name: str | None = None) -> dict:
    """Read-only total (sum) of a numeric column. Optional: year (filters rows of that
    year using the date column), group_by (one total per value of another column).
    Use this for questions like '2024 total revenue' or 'revenue by region'."""
    sheet = get_sheet(workbook_name, sheet_name)
    df, _ = load_dataframe(sheet)
    if column not in df.columns:
        raise ExcelToolError(f"Column '{column}' illa. Available: {list(df.columns)}")
    values = pd.to_numeric(df[column], errors="coerce")
    if values.notna().sum() == 0:
        raise ExcelToolError(f"Column '{column}' numeric illa, total kandupidikka mudiyaadhu.")

    rows_used = df
    if year is not None:
        if date_column is None:
            date_column = next((c for c in df.columns if "date" in str(c).lower()), None)
        if date_column is None or date_column not in df.columns:
            raise ExcelToolError(f"Date column kidaikkala. Available: {list(df.columns)}")
        dates = pd.to_datetime(df[date_column], errors="coerce")
        if dates.notna().sum() == 0:
            raise ExcelToolError(f"'{date_column}' la valid dates illa.")
        mask = dates.dt.year == int(year)
        rows_used = df[mask]
        values = values[mask]

    result = {"success": True, "column": column, "year": year,
              "rows_matched": int(len(rows_used)), "total": round(float(values.sum()), 2)}
    if group_by:
        if group_by not in df.columns:
            raise ExcelToolError(f"Column '{group_by}' illa. Available: {list(df.columns)}")
        grouped = values.groupby(rows_used[group_by]).sum().round(2)
        result["by_group"] = {str(k): float(v) for k, v in grouped.items()}
    return result