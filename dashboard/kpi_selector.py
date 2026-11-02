"""KPI selector: chooses and computes KPIs from an analysis result + DataFrame.

Pure pandas, no Excel access.
- auto_kpis(df, analysis): Auto Dashboard picks KPIs itself.
- build_kpis(df, requests): Guided Dashboard computes exactly what the user asked.
Test (no Excel needed): python -m dashboard.kpi_selector
"""
from __future__ import annotations

from typing import Any

import pandas as pd

MAX_AUTO_KPIS = 6
AGGREGATIONS = {"sum", "mean", "count", "min", "max", "median", "nunique"}
# Measures where adding values up is meaningless (prices, rates, percentages).
NO_SUM_HINTS = ("price", "rate", "percent", "pct", "ratio", "avg", "average", "margin")
# Preferred order when picking the main measure.
MAIN_MEASURE_HINTS = ("revenue", "sales", "amount", "total", "profit", "cost")

LABELS = {"sum": "Total", "mean": "Average", "count": "Count", "min": "Min",
          "max": "Max", "median": "Median", "nunique": "Unique"}


def _is_additive(column: str) -> bool:
    return not any(h in column.lower() for h in NO_SUM_HINTS)


def _pretty(column: str) -> str:
    return column.replace("_", " ").title()


def _compute(df: pd.DataFrame, column: str | None, agg: str) -> float | int:
    if agg == "count" and column is None:
        return int(len(df))
    series = df[column]
    if agg == "nunique":
        return int(series.nunique())
    if agg == "count":
        return int(series.count())
    value = getattr(series, agg)()
    return 0.0 if pd.isna(value) else float(value)


def make_kpi(df: pd.DataFrame, column: str | None, agg: str, label: str | None = None) -> dict[str, Any]:
    """Compute one KPI. column=None with agg='count' means row count."""
    if agg not in AGGREGATIONS:
        raise ValueError(f"Unknown aggregation '{agg}'. Allowed: {sorted(AGGREGATIONS)}")
    if column is not None and column not in df.columns:
        raise ValueError(f"Column '{column}' illa. Available: {list(df.columns)}")
    if column is not None and agg in {"sum", "mean", "min", "max", "median"} \
            and not pd.api.types.is_numeric_dtype(df[column]):
        raise ValueError(f"'{column}' numeric illa, '{agg}' pannamudiyaadhu.")
    if label is None:
        label = "Order Count" if column is None else f"{LABELS[agg]} {_pretty(column)}"
    return {"label": label, "column": column, "agg": agg, "value": _compute(df, column, agg)}


def auto_kpis(df: pd.DataFrame, analysis: dict[str, Any]) -> list[dict[str, Any]]:
    """Pick a small, relevant KPI set automatically."""
    if not analysis.get("success"):
        return []
    numeric: list[str] = analysis.get("numeric", [])
    ordered = sorted(numeric, key=lambda c: (
        0 if any(h in c.lower() for h in MAIN_MEASURE_HINTS) else 1, numeric.index(c)))

    kpis: list[dict[str, Any]] = [make_kpi(df, None, "count", "Total Records")]
    for column in ordered:
        if _is_additive(column):
            kpis.append(make_kpi(df, column, "sum"))
    for column in ordered:
        if not _is_additive(column):
            kpis.append(make_kpi(df, column, "mean"))
    if ordered and _is_additive(ordered[0]):
        kpis.append(make_kpi(df, ordered[0], "mean"))
    return kpis[:MAX_AUTO_KPIS]


def build_kpis(df: pd.DataFrame, requests: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Guided: requests = [{"column": "revenue", "agg": "sum", "label": optional}, ...]."""
    return [make_kpi(df, r.get("column"), r.get("agg", "sum"), r.get("label")) for r in requests]


def format_value(kpi: dict[str, Any]) -> str:
    """Human readable KPI value for the dashboard card."""
    value = kpi["value"]
    if kpi["agg"] in {"count", "nunique"} or float(value).is_integer() and abs(value) >= 1000:
        return f"{value:,.0f}"
    return f"{value:,.2f}"


if __name__ == "__main__":
    demo = pd.DataFrame({
        "rukku": range(1, 5), "region": ["N", "S", "N", "E"],
        "units": [10, 20, 30, 40], "unit_price": [5.0, 5.0, 10.0, 2.5],
        "revenue": [50.0, 100.0, 300.0, 100.0],
    })
    from dashboard.dataset_analyser import analyse_dataframe
    info = analyse_dataframe(demo)
    print("AUTO KPIs:")
    for k in auto_kpis(demo, info):
        print(f"  {k['label']:<22} {format_value(k)}")
    print("GUIDED KPIs:")
    for k in build_kpis(demo, [{"column": "revenue", "agg": "max"}, {"column": "region", "agg": "nunique"}]):
        print(f"  {k['label']:<22} {format_value(k)}")
    try:
        make_kpi(demo, "region", "sum")
    except ValueError as exc:
        print("Expected error:", exc)