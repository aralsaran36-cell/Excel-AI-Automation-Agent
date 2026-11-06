"""Dashboard filter: filters a DataFrame by user rules. Pure pandas, never touches Excel.

A rule is a dict:
  {"column": "region", "op": "in",  "value": ["EMEA", "APAC"]}
  {"column": "units",  "op": ">",   "value": 30}
  {"column": "units",  "op": "between", "value": [10, 50]}
  {"column": "date",   "op": "between", "value": ["2025-01-01", "2025-03-31"]}
Ops: ==, !=, >, >=, <, <=, in, not_in, between, contains
Test (no Excel needed): python -m dashboard.dashboard_filter
"""
from __future__ import annotations

from typing import Any

import pandas as pd

OPS = {"==", "!=", ">", ">=", "<", "<=", "in", "not_in", "between", "contains"}


def suggest_filters(analysis: dict[str, Any], max_values: int = 10) -> list[dict[str, Any]]:
    """Useful filter options (categorical values, numeric ranges, date ranges)."""
    options: list[dict[str, Any]] = []
    for col, info in analysis.get("columns", {}).items():
        if info["kind"] == "categorical":
            options.append({"column": col, "kind": "categorical", "top_values": info.get("top_values", [])[:max_values]})
        elif info["kind"] == "numeric":
            options.append({"column": col, "kind": "numeric", "min": info["min"], "max": info["max"]})
        elif info["kind"] == "date":
            options.append({"column": col, "kind": "date", "min": info["min"], "max": info["max"]})
    return options


def _match_column(df: pd.DataFrame, name: str) -> str:
    """Exact match first, then case/space-insensitive."""
    if name in df.columns:
        return name
    wanted = str(name).strip().lower()
    for col in df.columns:
        if str(col).strip().lower() == wanted:
            return col
    raise ValueError(f"Column '{name}' not found. Available: {list(df.columns)}")


def _mask(df: pd.DataFrame, rule: dict[str, Any]) -> pd.Series:
    col = _match_column(df, rule.get("column", ""))
    op, value = rule.get("op", "=="), rule.get("value")
    if op not in OPS:
        raise ValueError(f"Operator '{op}' is not supported. Allowed: {sorted(OPS)}")
    series = df[col]

    if op in {"in", "not_in"}:
        values = value if isinstance(value, (list, tuple, set)) else [value]
        if pd.api.types.is_numeric_dtype(series):
            hit = series.isin([float(v) for v in values])
        else:
            lowered = {str(v).strip().lower() for v in values}
            hit = series.astype(str).str.strip().str.lower().isin(lowered)
        return hit if op == "in" else ~hit

    if op == "contains":
        return series.astype(str).str.contains(str(value), case=False, na=False, regex=False)

    if pd.api.types.is_datetime64_any_dtype(series):
        convert = pd.to_datetime
    elif pd.api.types.is_numeric_dtype(series):
        convert = float
    else:
        if op in {"==", "!="}:
            same = series.astype(str).str.strip().str.lower() == str(value).strip().lower()
            return same if op == "==" else ~same
        raise ValueError(f"'{col}' is a text column, '{op}' cannot be used.")

    if op == "between":
        if not isinstance(value, (list, tuple)) or len(value) != 2:
            raise ValueError("between needs [low, high].")
        low, high = convert(value[0]), convert(value[1])
        return series.between(low, high)
    target = convert(value)
    return {"==": series == target, "!=": series != target, ">": series > target,
            ">=": series >= target, "<": series < target, "<=": series <= target}[op]


def apply_filters(df: pd.DataFrame, rules: list[dict[str, Any]] | None) -> dict[str, Any]:
    """Apply all rules (AND). Returns the filtered DataFrame; the original is never changed."""
    if not rules:
        return {"success": True, "df": df, "applied": [], "rows_before": len(df), "rows_after": len(df)}
    try:
        mask = pd.Series(True, index=df.index)
        applied = []
        for rule in rules:
            mask &= _mask(df, rule).fillna(False)
            applied.append(f"{rule.get('column')} {rule.get('op', '==')} {rule.get('value')}")
    except (ValueError, TypeError) as exc:
        return {"success": False, "error": str(exc)}
    out = df[mask].reset_index(drop=True)
    if out.empty:
        return {"success": False, "error": "No rows match this filter.", "applied": applied}
    return {"success": True, "df": out, "applied": applied, "rows_before": len(df), "rows_after": len(out)}


if __name__ == "__main__":
    demo = pd.DataFrame({
        "region": ["EMEA", "APAC", "EMEA", "AMER"],
        "units": [10, 20, 30, 40],
        "day": pd.to_datetime(["2025-01-05", "2025-02-10", "2025-03-15", "2025-04-20"]),
    })

    def show(label: str, rules):
        r = apply_filters(demo, rules)
        print(f"{label:<28}", r["rows_after"] if r["success"] else "ERROR: " + r["error"])

    show("no rules", [])
    show("region == emea (case)", [{"column": "Region", "op": "==", "value": "emea"}])
    show("region in EMEA,APAC", [{"column": "region", "op": "in", "value": ["EMEA", "APAC"]}])
    show("units > 15", [{"column": "units", "op": ">", "value": 15}])
    show("units between 15 and 35", [{"column": "units", "op": "between", "value": [15, 35]}])
    show("day between Feb-Mar", [{"column": "day", "op": "between", "value": ["2025-02-01", "2025-03-31"]}])
    show("EMEA and units>15", [{"column": "region", "op": "==", "value": "EMEA"}, {"column": "units", "op": ">", "value": 15}])
    show("no match", [{"column": "units", "op": ">", "value": 999}])
    show("bad column", [{"column": "xyz", "op": "==", "value": 1}])
    show("text with >", [{"column": "region", "op": ">", "value": "A"}])
    print("original rows untouched:", len(demo))