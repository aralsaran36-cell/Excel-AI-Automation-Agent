"""Chart selector: scores candidate charts and keeps only the useful ones.

Returns plans only (no Excel access). Number of charts is dynamic (1..MAX_CHARTS),
decided by how many candidates score above MIN_SCORE.
- auto_charts(df, analysis): Auto Dashboard.
- build_charts(df, requests): Guided Dashboard validates user-chosen charts.
Test (no Excel needed): python -m dashboard.chart_selector
"""
from __future__ import annotations

from typing import Any

import pandas as pd

MAX_CHARTS = 6
SCATTER_ENABLED = False   # scatter plotted wrong on Mac (x = row number); re-enable after the fix
MIN_CHARTS = 1
MIN_SCORE = 50
PIE_MAX_CATEGORIES = 5
BAR_MAX_CATEGORIES = 15
TOP_N = 10
CHART_TYPES = {"bar", "line", "pie", "column", "scatter", "stacked"}
AGGS = {"sum", "mean", "count", "min", "max", "median"}
NO_SUM_HINTS = ("price", "rate", "percent", "pct", "ratio", "avg", "average", "margin")
MAIN_MEASURE_HINTS = ("revenue", "sales", "amount", "total", "profit", "cost", "salary", "marks", "score")


def _pretty(name: str) -> str:
    return name.replace("_", " ").title()


def _is_additive(column: str) -> bool:
    return not any(h in column.lower() for h in NO_SUM_HINTS)


def _measure_rank(numeric: list[str]) -> list[str]:
    return sorted(numeric, key=lambda c: (
        0 if _is_additive(c) else 1,
        0 if any(h in c.lower() for h in MAIN_MEASURE_HINTS) else 1,
        numeric.index(c)))


def pick_main_measure(analysis: dict[str, Any]) -> str | None:
    numeric: list[str] = analysis.get("numeric", [])
    return _measure_rank(numeric)[0] if numeric else None


def make_chart(df: pd.DataFrame, chart_type: str, category: str, measure: str,
               agg: str = "sum", title: str | None = None,
               measure2: str | None = None, series: str | None = None) -> dict[str, Any]:
    """Validate one chart request and return its plan."""
    if chart_type not in CHART_TYPES:
        raise ValueError(f"Chart type '{chart_type}' support illa. Allowed: {sorted(CHART_TYPES)}")
    needed = [category, measure] + ([measure2] if measure2 else []) + ([series] if series else [])
    for col in needed:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' illa. Available: {list(df.columns)}")
    if not pd.api.types.is_numeric_dtype(df[measure]):
        raise ValueError(f"'{measure}' numeric illa.")
    if agg not in AGGS:
        raise ValueError(f"Aggregation '{agg}' support illa.")
    n_cat = int(df[category].nunique())
    if chart_type == "pie" and (n_cat > PIE_MAX_CATEGORIES or agg not in {"sum", "count"}):
        raise ValueError("Pie chart: max 5 categories and sum/count mattum (misleading-a aagum).")
    if chart_type == "pie" and (df[measure] < 0).any():
        raise ValueError("Pie chart-la negative values irukku.")
    if chart_type == "scatter":
        if not measure2 or not pd.api.types.is_numeric_dtype(df[measure2]):
            raise ValueError("Scatter-ku rendu numeric columns venum.")
    if chart_type == "stacked" and not series:
        raise ValueError("Stacked chart-ku series column venum.")
    if title is None:
        verb = {"sum": "", "mean": "Average "}.get(agg, f"{agg.title()} ")
        if chart_type == "scatter":
            title = f"{_pretty(measure)} vs {_pretty(measure2)}"
        elif chart_type == "stacked":
            title = f"{verb}{_pretty(measure)} by {_pretty(category)} and {_pretty(series)}"
        elif chart_type == "line":
            title = f"{verb}{_pretty(measure)} Trend"
        else:
            title = f"{verb}{_pretty(measure)} by {_pretty(category)}"
    plan = {"type": chart_type, "category": category, "measure": measure,
            "agg": agg, "title": title, "categories": n_cat}
    if measure2:
        plan["measure2"] = measure2
    if series:
        plan["series"] = series
    return plan


def _candidates(df: pd.DataFrame, analysis: dict[str, Any]) -> list[tuple[int, dict[str, Any]]]:
    numeric = _measure_rank(analysis.get("numeric", []))
    cats = [c for c in analysis.get("categorical", []) if df[c].nunique() >= 2]
    dates = analysis.get("date", [])
    out: list[tuple[int, dict[str, Any]]] = []
    if not numeric:
        return out
    main = numeric[0]
    main_agg = "sum" if _is_additive(main) else "mean"

    # Trend over time: strongest chart when a date exists.
    for d in dates[:1]:
        if df[d].nunique() >= 3:
            out.append((95, make_chart(df, "line", d, main, main_agg)))

    # Category x measure (bar / pie / top-N).
    for ci, cat in enumerate(cats):
        n = int(df[cat].nunique())
        for mi, meas in enumerate(numeric[:3]):
            agg = "sum" if _is_additive(meas) else "mean"
            base = 90 - ci * 6 - mi * 22          # main measure first, later ones lower
            if n <= PIE_MAX_CATEGORIES and agg == "sum" and (df[meas] >= 0).all() and n <= 3 and mi == 0:
                out.append((base, make_chart(df, "pie", cat, meas, agg)))
            elif n <= BAR_MAX_CATEGORIES:
                out.append((base - (8 if n > 10 else 0), make_chart(df, "bar", cat, meas, agg)))
            else:
                plan = make_chart(df, "bar", cat, meas, agg,
                                  title=f"Top {TOP_N} {_pretty(cat)} by {_pretty(meas)}")
                plan["top_n"] = TOP_N
                out.append((base - 10, plan))

    # Two categories -> stacked column (only if the grid is readable).
    if len(cats) >= 2:
        a, b = cats[0], cats[1]
        if df[a].nunique() <= 8 and df[b].nunique() <= 5:
            out.append((70, make_chart(df, "stacked", a, main, main_agg, series=b)))

    # Two numerics with a real relationship -> scatter.
    if SCATTER_ENABLED and len(numeric) >= 2 and len(df) >= 10:
        x, y = numeric[0], numeric[1]
        corr = df[x].corr(df[y])
        if pd.notna(corr) and abs(corr) >= 0.3:
            out.append((int(55 + 40 * abs(corr)), make_chart(df, "scatter", x, x, "sum", measure2=y)))
    return out


def auto_charts(df: pd.DataFrame, analysis: dict[str, Any]) -> list[dict[str, Any]]:
    """Pick useful charts. Count depends on the data (MIN_CHARTS..MAX_CHARTS)."""
    if not analysis.get("success"):
        return []
    scored = sorted(_candidates(df, analysis), key=lambda s: -s[0])
    chosen, seen, used_categories = [], set(), set()
    for score, plan in scored:
        key = (plan["type"], plan["category"], plan["measure"], plan.get("series"), plan.get("measure2"))
        if key in seen:
            continue
        if plan["type"] in {"bar", "pie"} and plan["category"] in used_categories:
            continue  # one chart per category (no "Units by Region" after "Revenue by Region")
        seen.add(key)
        if score >= MIN_SCORE or len(chosen) < MIN_CHARTS:
            chosen.append(plan)
            if plan["type"] in {"bar", "pie"}:
                used_categories.add(plan["category"])
        if len(chosen) >= MAX_CHARTS:
            break
    return chosen


def build_charts(df: pd.DataFrame, requests: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Guided: requests = [{"type": "bar", "category": "region", "measure": "revenue", "agg": "sum"}]."""
    return [make_chart(df, r.get("type", "bar"), r["category"], r["measure"], r.get("agg", "sum"),
                       r.get("title"), r.get("measure2"), r.get("series")) for r in requests]


if __name__ == "__main__":
    from dashboard.dataset_analyser import analyse_dataframe
    import numpy as np
    rng = np.random.default_rng(1)

    sales = pd.DataFrame({
        "region": rng.choice(["N", "S", "E", "W", "C"], 60),
        "channel": rng.choice(["web", "shop", "app"], 60),
        "units": rng.integers(1, 50, 60)})
    sales["revenue"] = sales["units"] * rng.uniform(5, 6, 60)

    hr = pd.DataFrame({
        "department": rng.choice(["IT", "HR", "Ops", "Fin"], 80),
        "gender": rng.choice(["M", "F"], 80),
        "city": rng.choice([f"city{i}" for i in range(25)], 80),
        "hire_date": pd.date_range("2022-01-01", periods=80, freq="7D"),
        "salary": rng.integers(30000, 90000, 80)})

    tiny = pd.DataFrame({"fruit": ["a", "b", "a", "c"], "qty": [3, 4, 5, 6]})

    for label, frame in [("sales-like", sales), ("hr-like", hr), ("tiny", tiny)]:
        plans = auto_charts(frame, analyse_dataframe(frame))
        print(f"{label}: {len(plans)} charts")
        for p in plans:
            print(f"   {p['type']:<8} {p['title']}")
    try:
        make_chart(sales, "pie", "region", "revenue", "mean")
    except ValueError as exc:
        print("Expected error:", exc)