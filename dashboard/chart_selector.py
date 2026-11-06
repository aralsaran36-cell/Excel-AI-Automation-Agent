"""Chart selector: scores candidate charts and keeps only the useful ones.

Returns plans only (no Excel access). Uses measure rules, so the aggregation is right:
sum measure -> sum, everything else -> average, 0/1 column -> rate by category.
- auto_charts(df, analysis, variant=0): chart COUNT (3..6) depends on how rich the data is.
- build_charts(df, requests): Guided Dashboard validates user-chosen charts.
Constants live in dashboard_config.py.
Test (no Excel needed): python -m dashboard.chart_selector
"""
from __future__ import annotations

from typing import Any

import pandas as pd

from dashboard.dashboard_config import (
    BAR_MAX_CATEGORIES, MAX_CHARTS, MIN_CHARTS, MIN_SCORE, MIN_TARGET,
    PIE_MAX_CATEGORIES, SCATTER_ENABLED, TOP_N,
)
from dashboard.measure import agg_for, classify_measure, rank_measures

CHART_TYPES = {"bar", "line", "pie", "column", "scatter", "stacked", "hist"}
AGGS = {"sum", "mean", "count", "min", "max", "median"}


def _pretty(name: str) -> str:
    return str(name).replace("_", " ").title()


def make_chart(df: pd.DataFrame, chart_type: str, category: str, measure: str,
               agg: str = "sum", title: str | None = None,
               measure2: str | None = None, series: str | None = None) -> dict[str, Any]:
    """Validate one chart request and return its plan."""
    if chart_type not in CHART_TYPES:
        raise ValueError(f"Chart type '{chart_type}' is not supported. Allowed: {sorted(CHART_TYPES)}")
    needed = [category, measure] + ([measure2] if measure2 else []) + ([series] if series else [])
    for col in needed:
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found. Available: {list(df.columns)}")
    if not pd.api.types.is_numeric_dtype(df[measure]):
        raise ValueError(f"'{measure}' is not numeric.")
    if agg not in AGGS:
        raise ValueError(f"Aggregation '{agg}' is not supported.")
    n_cat = int(df[category].nunique())
    if chart_type == "pie" and (n_cat > PIE_MAX_CATEGORIES or agg not in {"sum", "count"}):
        raise ValueError("Pie chart: max 5 categories and sum/count only (otherwise misleading).")
    if chart_type == "pie" and (df[measure] < 0).any():
        raise ValueError("Pie chart cannot have negative values.")
    if chart_type == "scatter":
        if not measure2 or not pd.api.types.is_numeric_dtype(df[measure2]):
            raise ValueError("Scatter needs two numeric columns.")
    if chart_type == "stacked" and not series:
        raise ValueError("Stacked chart needs a series column.")
    if title is None:
        rate = agg == "mean" and classify_measure(df, measure) == "binary"
        verb = "" if agg == "sum" else ("Average " if agg == "mean" else f"{agg.title()} ")
        name = f"{_pretty(measure)} Rate" if rate else f"{verb}{_pretty(measure)}"
        if chart_type == "scatter":
            title = f"{_pretty(measure)} vs {_pretty(measure2)}"
        elif chart_type == "stacked":
            title = f"{name} by {_pretty(category)} and {_pretty(series)}"
        elif chart_type == "line":
            title = f"{name} Trend"
        else:
            title = f"{name} by {_pretty(category)}"
    plan = {"type": chart_type, "category": category, "measure": measure,
            "agg": agg, "title": title, "categories": n_cat}
    if measure2:
        plan["measure2"] = measure2
    if series:
        plan["series"] = series
    return plan


def _add(out: list, score: int, *args, **kwargs) -> dict[str, Any] | None:
    """Add a candidate; a request that fails validation is skipped, not fatal."""
    try:
        plan = make_chart(*args, **kwargs)
    except ValueError:
        return None
    out.append((score, plan))
    return plan


def _context(df: pd.DataFrame, analysis: dict[str, Any]):
    kinds: dict[str, str] = analysis.get("measures") or {
        c: classify_measure(df, c) for c in analysis.get("numeric", [])}
    ranked = rank_measures(kinds)
    cats = [c for c in analysis.get("categorical", []) if df[c].nunique() >= 2]
    dates = analysis.get("date", [])
    return kinds, ranked, cats, dates


def _target_count(df: pd.DataFrame, kinds, ranked, cats, dates, variant: int = 0) -> int:
    """How many charts this dataset deserves: richer data -> more charts (3..6)."""
    n = 1
    if dates and df[dates[0]].nunique() >= 3:
        n += 1                                                    # trend chart
    n += min(len(cats), 2)                                        # one chart per main category
    if len([m for m in ranked if kinds[m] in {"sum", "mean"}]) >= 2:
        n += 1                                                    # second measure
    if any(kinds[m] == "binary" for m in ranked):
        n += 1                                                    # rate chart (e.g. survival)
    if variant % 2 == 1:                                          # "another dashboard" changes the size a bit
        n = n + 1 if n < MAX_CHARTS else n - 1
    return max(MIN_TARGET, min(MAX_CHARTS, n))


def _candidates(df: pd.DataFrame, analysis: dict[str, Any], variant: int = 0) -> list[tuple[int, dict[str, Any]]]:
    kinds, ranked, cats, dates = _context(df, analysis)
    out: list[tuple[int, dict[str, Any]]] = []
    if not ranked:
        return out
    if variant:                                   # different look = different measure/category order
        ranked = ranked[variant % len(ranked):] + ranked[:variant % len(ranked)]
        if len(cats) > 1:
            shift = variant % len(cats)
            cats = cats[shift:] + cats[:shift]
    main = ranked[0]
    main_agg = agg_for(kinds[main])

    # Trend over time: strongest chart when a date exists.
    for d in dates[:1]:
        if df[d].nunique() >= 3:
            _add(out, 95, df, "line", d, main, main_agg)

    # Distribution of an attribute (Age Distribution, Fare Distribution).
    for m in ranked:
        if kinds[m] == "mean" and df[m].nunique() >= 10:
            _add(out, 78, df, "hist", m, m, "count", f"{_pretty(m)} Distribution")
            break

    # Record counts by category (Passengers by Class, Orders by Region).
    count_col = next((c for c in analysis.get("id", []) if pd.api.types.is_numeric_dtype(df[c])), main)
    for ci, cat in enumerate(cats[:2]):
        n = int(df[cat].nunique())
        title = f"Records by {_pretty(cat)}"
        if n <= PIE_MAX_CATEGORIES:
            _add(out, 84 - ci * 6, df, "pie", cat, count_col, "count", title)
        elif n <= BAR_MAX_CATEGORIES:
            _add(out, 80 - ci * 6, df, "bar", cat, count_col, "count", title)

    # Category x measure (bar / pie / top-N). 0/1 columns become "rate by category".
    for ci, cat in enumerate(cats):
        n = int(df[cat].nunique())
        for mi, meas in enumerate(ranked[:3]):
            agg = agg_for(kinds[meas])
            base = 90 - ci * 6 - mi * 22
            if n <= 3 and agg == "sum" and mi == 0 and (df[meas] >= 0).all():
                _add(out, base, df, "pie", cat, meas, agg)
            elif n <= BAR_MAX_CATEGORIES:
                _add(out, base - (8 if n > 10 else 0), df, "bar", cat, meas, agg)
            else:
                plan = _add(out, base - 10, df, "bar", cat, meas, agg,
                            f"Top {TOP_N} {_pretty(cat)} by {_pretty(meas)}")
                if plan:
                    plan["top_n"] = TOP_N

    # Two categories -> stacked column (only if the grid is readable).
    if len(cats) >= 2:
        a, b = cats[0], cats[1]
        if df[a].nunique() <= 8 and df[b].nunique() <= 5:
            _add(out, 70, df, "stacked", a, main, main_agg, series=b)

    # Two numerics with a real relationship -> scatter.
    if SCATTER_ENABLED and len(ranked) >= 2 and len(df) >= 10:
        x, y = ranked[0], ranked[1]
        corr = df[x].corr(df[y])
        if pd.notna(corr) and abs(corr) >= 0.3:
            _add(out, int(55 + 40 * abs(corr)), df, "scatter", x, x, "sum", measure2=y)
    return out


def auto_charts(df: pd.DataFrame, analysis: dict[str, Any], variant: int = 0) -> list[dict[str, Any]]:
    """Pick useful charts. Count = _target_count (3..6), at most one bar/pie per category in pass 1."""
    if not analysis.get("success"):
        return []
    kinds, ranked, cats, dates = _context(df, analysis)
    if not ranked:
        return []
    target = _target_count(df, kinds, ranked, cats, dates, variant)
    scored = sorted(_candidates(df, analysis, variant), key=lambda s: -s[0])
    chosen: list[dict[str, Any]] = []
    seen: set = set()
    used_cat: set = set()
    used_slot: set = set()
    for strict in (True, False):                  # pass 2 relaxes "one chart per category" if still short
        for score, plan in scored:
            if len(chosen) >= target:
                break
            key = (plan["type"], plan["category"], plan["measure"], plan["agg"],
                   plan.get("series"), plan.get("measure2"))
            if key in seen:
                continue
            if plan["type"] in {"bar", "pie"}:
                slot = (plan["category"], plan["agg"] == "count")
                if slot in used_slot or (strict and plan["category"] in used_cat):
                    continue
            if score < MIN_SCORE and len(chosen) >= MIN_CHARTS:
                continue
            seen.add(key)
            chosen.append(plan)
            if plan["type"] in {"bar", "pie"}:
                used_cat.add(plan["category"])
                used_slot.add((plan["category"], plan["agg"] == "count"))
    return chosen


def build_charts(df: pd.DataFrame, requests: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Guided: requests = [{"type": "bar", "category": "region", "measure": "revenue", "agg": "sum"}]."""
    return [make_chart(df, r.get("type", "bar"), r["category"], r["measure"], r.get("agg", "sum"),
                       r.get("title"), r.get("measure2"), r.get("series")) for r in requests]


if __name__ == "__main__":
    from dashboard.dataset_analyser import analyse_dataframe
    titanic = pd.DataFrame({
        "PassengerId": range(1, 13), "Survived": [0, 1, 1, 0, 1, 0, 0, 1, 1, 0, 1, 0],
        "Pclass": [3, 1, 3, 1, 2, 3, 1, 2, 3, 3, 1, 2], "Sex": list("MFFMFMMFFMFM"),
        "Age": [22, 38, 26, 35, 54, 2, 27, 14, 4, 58.5, 30, 41],
        "Fare": [7.25, 71.3, 7.9, 53.1, 8.05, 21.1, 11.1, 30.1, 16.7, 26.5, 40, 12]})
    info = analyse_dataframe(titanic)
    for v in (0, 1, 2):
        plans = auto_charts(titanic, info, v)
        print(f"variant {v}: {len(plans)} charts")
        for p in plans:
            print(f"   {p['type']:<8} {p['agg']:<6} {p['title']}")