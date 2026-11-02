"""Dataset analyser: understands a sheet's structure for the dashboard modules.

Read-only. analyse_dataframe() is pure pandas (testable without Excel).
analyse_sheet() reads the active (or named) sheet through excel_tools.
Test: python -m dashboard.dataset_analyser
"""
from __future__ import annotations

from typing import Any

import pandas as pd

from excel_tools.workbook_manager import get_sheet, load_dataframe

CATEGORY_MAX_UNIQUE = 30      # <= this many distinct values -> usable as a category
CATEGORY_MAX_RATIO = 0.5      # and not more than half of the rows
ID_NAME_HINTS = ("id", "rukku", "index", "no", "number", "code")
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


def _is_id_like(name: str, series: pd.Series) -> bool:
    """Unique-per-row columns or id-named integer columns are not real measures."""
    non_null = series.dropna()
    if non_null.empty:
        return False
    all_unique = non_null.nunique() == len(non_null) and len(non_null) > 1
    lowered = name.strip().lower()
    named_id = lowered in ID_NAME_HINTS or lowered.endswith(("_id", " id"))
    return all_unique and named_id 


def _classify(name: str, series: pd.Series) -> str:
    """Return one of: date, numeric, categorical, id, text, empty."""
    non_null = series.dropna()
    if non_null.empty:
        return "empty"
    if pd.api.types.is_datetime64_any_dtype(series):
        return "date"
    if pd.api.types.is_bool_dtype(series):
        return "categorical"
    if pd.api.types.is_numeric_dtype(series):
        return "id" if _is_id_like(name, series) else "numeric"
    if series.map(lambda v: hasattr(v, "year") and hasattr(v, "month")).all():
        return "date"
    unique = non_null.nunique()
    if unique <= CATEGORY_MAX_UNIQUE and unique / len(non_null) <= CATEGORY_MAX_RATIO:
        return "categorical"
    return "id" if unique == len(non_null) else "text"


def analyse_dataframe(df: pd.DataFrame) -> dict[str, Any]:
    """Analyse a DataFrame. Returns a plain dict used by kpi/chart selectors."""
    if df is None or df.empty:
        return {"success": False, "error": "Dataset empty-a irukku."}

    columns: dict[str, dict[str, Any]] = {}
    for name in df.columns:
        series = df[name]
        kind = _classify(name, series)
        info: dict[str, Any] = {
            "kind": kind,
            "missing": int(series.isna().sum()),
            "unique": int(series.nunique(dropna=True)),
        }
        if kind == "numeric":
            info.update(sum=float(series.sum()), mean=float(series.mean()),
                        min=float(series.min()), max=float(series.max()))
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
            print(f"  {col:<12} {info['kind']:<12} missing={info['missing']} unique={info['unique']}")
        print("numeric    :", out["numeric"])
        print("categorical:", out["categorical"])
        print("date       :", out["date"])
        print("id         :", out["id"])
        print("missing_total:", out["missing_total"], "| duplicate_rows:", out["duplicate_rows"])