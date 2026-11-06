"""Dataset analyser: understands a sheet's structure for the dashboard modules.

Read-only. analyse_dataframe() is pure pandas (testable without Excel).
Numeric columns are classified with measure_rules:
  sum/mean/binary -> "numeric" (usable measures), discrete -> "categorical", id -> "id".
Test: python -m dashboard.dataset_analyser
"""
from __future__ import annotations

from typing import Any

import pandas as pd

from dashboard.measure import classify_measure
from excel_tools.workbook_manager import get_sheet, load_dataframe

CATEGORY_MAX_UNIQUE = 30      # <= this many distinct values -> usable as a category
CATEGORY_MAX_RATIO = 0.5      # and not more than half of the rows
DATE_TEXT_PATTERN = r"^\d{4}-\d{1,2}-\d{1,2}([ T]\d{1,2}:\d{2}(:\d{2})?)?$"


def parse_date_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy where text columns that are all ISO dates ('2024-07-01') become real dates.

    In memory only: the Excel sheet is never changed.
    """
    out = df.copy()
    for col in out.columns:
        series = out[col]
        if (pd.api.types.is_numeric_dtype(series) or pd.api.types.is_bool_dtype(series)
                or pd.api.types.is_datetime64_any_dtype(series)):
            continue
        non_null = series.dropna()
        if non_null.empty or not non_null.map(lambda v: isinstance(v, str)).all():
            continue
        if non_null.astype(str).str.strip().str.match(DATE_TEXT_PATTERN).all():
            out[col] = pd.to_datetime(series, errors="coerce")
    return out


def _classify(name: str, series: pd.Series) -> tuple[str, str | None]:
    """Return (kind, measure_kind). kind: date, numeric, categorical, id, text, empty."""
    non_null = series.dropna()
    if non_null.empty:
        return "empty", None
    if pd.api.types.is_datetime64_any_dtype(series):
        return "date", None
    if pd.api.types.is_bool_dtype(series):
        return "categorical", None
    if pd.api.types.is_numeric_dtype(series):
        measure = classify_measure(pd.DataFrame({name: series}), name)
        if measure == "id":
            return "id", measure
        if measure == "discrete":
            if non_null.nunique() <= CATEGORY_MAX_UNIQUE:
                return "categorical", measure     # Pclass, SibSp, year behave like categories
            return "numeric", "mean"
        if measure == "skip":
            return "empty", None
        return "numeric", measure                 # sum, mean or binary
    if series.map(lambda v: hasattr(v, "year") and hasattr(v, "month")).all():
        return "date", None
    unique = non_null.nunique()
    if unique <= CATEGORY_MAX_UNIQUE and unique / len(non_null) <= CATEGORY_MAX_RATIO:
        return "categorical", None
    return ("id" if unique == len(non_null) else "text"), None


def analyse_dataframe(df: pd.DataFrame) -> dict[str, Any]:
    """Analyse a DataFrame. Returns a plain dict used by the kpi/chart selectors."""
    if df is None or df.empty:
        return {"success": False, "error": "Dataset is empty."}

    columns: dict[str, dict[str, Any]] = {}
    measures: dict[str, str] = {}
    for name in df.columns:
        series = df[name]
        kind, measure = _classify(str(name), series)
        info: dict[str, Any] = {
            "kind": kind,
            "missing": int(series.isna().sum()),
            "unique": int(series.nunique(dropna=True)),
        }
        if kind == "numeric":
            info.update(sum=float(series.sum()), mean=float(series.mean()),
                        min=float(series.min()), max=float(series.max()), measure=measure)
            measures[str(name)] = measure
        elif kind == "categorical":
            info["top_values"] = [str(v) for v in series.value_counts().head(5).index]
        elif kind == "date":
            info.update(min=str(series.min()), max=str(series.max()))
        columns[str(name)] = info

    def of_kind(kind: str) -> list[str]:
        return [c for c, i in columns.items() if i["kind"] == kind]

    return {
        "success": True,
        "rows": len(df),
        "column_count": len(df.columns),
        "columns": columns,
        "numeric": of_kind("numeric"),
        "measures": measures,                       # column -> sum | mean | binary
        "binary": [c for c, m in measures.items() if m == "binary"],
        "categorical": of_kind("categorical"),
        "date": of_kind("date"),
        "id": of_kind("id"),
        "text": of_kind("text"),
        "missing_total": int(df.isna().sum().sum()),
        "duplicate_rows": int(df.duplicated().sum()),
    }


def analyse_sheet(workbook_name: str | None = None, sheet_name: str | None = None) -> dict[str, Any]:
    """Read a sheet (active by default) and analyse it. Never writes to Excel."""
    try:
        sheet = get_sheet(workbook_name, sheet_name)
        df, _ = load_dataframe(sheet)
    except Exception as exc:
        return {"success": False, "error": f"{type(exc).__name__}: {exc}"}
    result = analyse_dataframe(df)
    if result.get("success"):
        result.update(workbook=sheet.book.name, sheet=sheet.name)
    return result


if __name__ == "__main__":
    out = analyse_sheet()
    if not out["success"]:
        print("ERROR:", out["error"])
    else:
        print(f"{out['workbook']} / {out['sheet']}: {out['rows']} rows, {out['column_count']} columns")
        for col, info in out["columns"].items():
            print(f"  {col:<14} {info['kind']:<12} {info.get('measure') or '':<7} unique={info['unique']}")
        print("measures   :", out["measures"])
        print("categorical:", out["categorical"])
        print("date       :", out["date"])
        print("id         :", out["id"])