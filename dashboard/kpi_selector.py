"""KPI selector: chooses and computes KPIs from an analysis result + DataFrame.

Pure pandas, no Excel access. Uses measure_rules, so Age is never summed:
  sum measure -> Total, mean measure -> Average, 0/1 column -> Rate %, ids -> not a KPI.
- auto_kpis(df, analysis, variant=0): variant > 0 gives a different KPI combination.
- build_kpis(df, requests): Guided Dashboard computes exactly what the user asked.
Test (no Excel needed): python -m dashboard.kpi_selector
"""
from __future__ import annotations

from typing import Any

import pandas as pd

from dashboard.measure import classify_measure, rank_measures

from dashboard.measure import classify_measure, rank_measures, tokens

DOMAIN_COUNT_LABEL = {"Sales": "Total Orders", "HR": "Headcount", "Education": "Students",
                      "Passengers": "Total Passengers", "Customer": "Customers",
                      "Healthcare": "Patients", "Inventory": "Items", "Survey": "Responses",
                      "Transport": "Trips", "Marketing": "Records", "Finance": "Transactions"}
DOMAIN_FOCUS = {"Sales": {"revenue", "sales", "amount", "profit"},
                "HR": {"salary", "experience", "tenure"},
                "Education": {"marks", "score", "grade"},
                "Finance": {"profit", "income", "expense", "balance"},
                "Passengers": {"survived", "fare", "age"},
                "Customer": {"churn", "tenure"},
                "Survey": {"rating", "satisfaction", "score"},
                "Marketing": {"conversion", "clicks", "ctr"},
                "Transport": {"fare", "distance"}}
MAX_AUTO_KPIS = 6
AGGREGATIONS = {"sum", "mean", "count", "min", "max", "median", "nunique", "rate"}
LABELS = {"sum": "Total", "mean": "Average", "count": "Count", "min": "Min",
          "max": "Max", "median": "Median", "nunique": "Unique", "rate": ""}


def _pretty(column: str) -> str:
    return str(column).replace("_", " ").title()


def _compute(df: pd.DataFrame, column: str | None, agg: str) -> float | int:
    if agg == "count" and column is None:
        return int(len(df))
    series = df[column]
    if agg == "nunique":
        return int(series.nunique())
    if agg == "count":
        return int(series.count())
    if agg == "rate":
        value = series.mean() * 100
        return 0.0 if pd.isna(value) else float(value)
    value = getattr(series, agg)()
    return 0.0 if pd.isna(value) else float(value)


def make_kpi(df: pd.DataFrame, column: str | None, agg: str, label: str | None = None) -> dict[str, Any]:
    """Compute one KPI. column=None with agg='count' means row count."""
    if agg not in AGGREGATIONS:
        raise ValueError(f"Unknown aggregation '{agg}'. Allowed: {sorted(AGGREGATIONS)}")
    if column is not None and column not in df.columns:
        raise ValueError(f"Column '{column}' not found. Available: {list(df.columns)}")
    if column is not None and agg in {"sum", "mean", "min", "max", "median", "rate"} \
            and not pd.api.types.is_numeric_dtype(df[column]):
        raise ValueError(f"'{column}' is not numeric, '{agg}' cannot be used.")
    if agg == "rate" and classify_measure(df, column) != "binary":
        raise ValueError(f"'{column}' is not a 0/1 column, 'rate' cannot be used.")
    if label is None:
        if column is None:
            label = "Total Records"
        elif agg == "rate":
            label = f"{_pretty(column)} Rate"
        else:
            label = f"{LABELS[agg]} {_pretty(column)}"
    return {"label": label, "column": column, "agg": agg, "value": _compute(df, column, agg)}


def auto_kpis(df: pd.DataFrame, analysis: dict[str, Any], variant: int = 0,
              domain: str | None = None) -> list[dict[str, Any]]:
    """Pick a small, relevant KPI set. variant > 0 rotates to a different combination."""
    if not analysis.get("success"):
        return []
    kinds: dict[str, str] = analysis.get("measures") or {
        c: classify_measure(df, c) for c in analysis.get("numeric", [])}
    ranked = rank_measures(kinds)

    candidates: list[dict[str, Any]] = []
    for col in ranked:                                   # 0/1 outcomes first: Survival Rate
        if kinds[col] == "binary":
            candidates.append(make_kpi(df, col, "rate"))
    for col in ranked:
        if kinds[col] == "sum":
            candidates.append(make_kpi(df, col, "sum"))
    for col in ranked:
        if kinds[col] == "mean":
            candidates.append(make_kpi(df, col, "mean"))
    for col in ranked:                                   # extra options (used by variants)
        if kinds[col] == "sum":
            candidates.append(make_kpi(df, col, "mean"))
            candidates.append(make_kpi(df, col, "max"))
        elif kinds[col] == "mean":
            candidates.append(make_kpi(df, col, "median"))
            candidates.append(make_kpi(df, col, "max"))
    for col in analysis.get("categorical", []):
        if 2 <= df[col].nunique() <= 30:
            candidates.append(make_kpi(df, col, "nunique", f"Unique {_pretty(col)}"))

    seen, rest = set(), []
    for k in candidates:
        key = (k["column"], k["agg"])
        if key not in seen:
            seen.add(key)
            rest.append(k)

    focus = DOMAIN_FOCUS.get(domain or "", set())        # domain knowledge: important columns first
    if focus:
        rest.sort(key=lambda k: 0 if (k["column"] and tokens(k["column"]) & focus) else 1)

    if variant and len(rest) > 1:
        shift = (variant * 2) % len(rest)
        rest = rest[shift:] + rest[:shift]
    count_label = DOMAIN_COUNT_LABEL.get(domain or "", "Total Records")
    return ([make_kpi(df, None, "count", count_label)] + rest)[:MAX_AUTO_KPIS]


def build_kpis(df: pd.DataFrame, requests: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Guided: requests = [{"column": "revenue", "agg": "sum", "label": optional}, ...]."""
    return [make_kpi(df, r.get("column"), r.get("agg", "sum"), r.get("label")) for r in requests]


def format_value(kpi: dict[str, Any]) -> str:
    """Human readable KPI value for the dashboard card."""
    value, agg = kpi["value"], kpi["agg"]
    if agg == "rate":
        return f"{value:.1f}%"
    if agg in {"count", "nunique"} or (float(value).is_integer() and abs(value) >= 1000):
        return f"{value:,.0f}"
    return f"{value:,.2f}"


if __name__ == "__main__":
    from dashboard.dataset_analyser import analyse_dataframe
    titanic = pd.DataFrame({
        "PassengerId": range(1, 11), "Survived": [0, 1, 1, 0, 1, 0, 0, 1, 1, 0],
        "Pclass": [3, 1, 3, 1, 2, 3, 1, 2, 3, 3], "Sex": list("MFFMFMMFFM"),
        "Age": [22, 38, 26, 35, 54, 2, 27, 14, 4, 58.5],
        "Fare": [7.25, 71.3, 7.9, 53.1, 8.05, 21.1, 11.1, 30.1, 16.7, 26.5]})
    info = analyse_dataframe(titanic)
    for v in (0, 1):
        print(f"TITANIC variant {v}:")
        for k in auto_kpis(titanic, info, v):
            print(f"  {k['label']:<22} {format_value(k)}")