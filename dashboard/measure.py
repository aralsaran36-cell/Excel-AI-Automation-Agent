"""Measure rules: decides what a numeric column MEANS, so KPIs/charts aggregate correctly.

classify_measure(df, col) -> "sum" | "mean" | "binary" | "discrete" | "id" | "skip"
  sum      additive amounts (revenue, units, expense)  -> Total
  mean     attributes (age, fare, price, score)        -> Average   (DEFAULT when unsure)
  binary   only 0/1 (survived, churned)                -> Rate %
  discrete few whole numbers (pclass, sibsp, year)     -> behaves like a category
  id       unique identifier                           -> ignored
Pure pandas, no Excel. Test: python -m dashboard.measure_rules
"""
from __future__ import annotations

import re

import pandas as pd

SUM_HINTS = {"revenue", "sales", "sale", "amount", "profit", "cost", "expense", "income", "salary",
             "units", "unit", "quantity", "qty", "orders", "spend", "budget", "payment", "sold", "volume"}
AVG_HINTS = {"age", "price", "rate", "percent", "pct", "ratio", "avg", "average", "margin", "score",
             "marks", "mark", "rating", "height", "weight", "temperature", "fare", "duration",
             "experience", "tenure", "discount", "grade", "cgpa", "gpa", "speed", "distance"}
DISCRETE_HINTS = {"class", "year", "month", "week", "level", "rank", "zip", "pin", "phone", "pclass"}
ID_HINTS = {"id", "index", "idx", "key", "uuid"}
MAIN_HINTS = {"revenue", "sales", "amount", "profit", "cost", "income", "salary"}
DISCRETE_MAX_UNIQUE = 10


def tokens(name: str) -> set[str]:
    spaced = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", str(name))
    out = set()
    for t in re.split(r"[^a-z0-9]+", spaced.lower()):
        if t:
            out.add(t)
            out.add(t.rstrip("s"))
    return out


def classify_measure(df: pd.DataFrame, col: str) -> str:
    s = df[col].dropna()
    if s.empty or not pd.api.types.is_numeric_dtype(s):
        return "skip"
    toks, n = tokens(col), s.nunique()
    if toks & ID_HINTS and n == len(s) and len(s) > 1:
        return "id"
    if n <= 2 and set(s.astype(float).unique()) <= {0.0, 1.0}:
        return "binary"
    if "total" in toks:
        return "sum"
    if toks & AVG_HINTS:
        return "mean"
    if toks & SUM_HINTS:
        return "sum"
    if toks & DISCRETE_HINTS:
        return "discrete"
    if (s % 1 == 0).all() and n <= DISCRETE_MAX_UNIQUE:
        return "discrete"
    return "mean"          # safe default: averaging is never nonsense, summing often is


def classify_all(df: pd.DataFrame, numeric: list[str]) -> dict[str, str]:
    return {c: classify_measure(df, c) for c in numeric}


def agg_for(kind: str) -> str:
    return "sum" if kind == "sum" else "mean"      # binary -> mean = rate


def rank_measures(kinds: dict[str, str]) -> list[str]:
    """Usable measures, best first: additive amounts, then 0/1 outcomes, then averages."""
    order = {"sum": 0, "binary": 1, "mean": 2}
    cols = [c for c, k in kinds.items() if k in order]
    return sorted(cols, key=lambda c: (order[kinds[c]],
                                       0 if tokens(c) & MAIN_HINTS else 1, cols.index(c)))


if __name__ == "__main__":
    titanic = pd.DataFrame({
        "PassengerId": range(1, 11), "Survived": [0, 1, 1, 0, 1, 0, 0, 1, 1, 0],
        "Pclass": [3, 1, 3, 1, 2, 3, 1, 2, 3, 3], "Age": [22, 38, 26, 35, 54, 2, 27, 14, 4, 58.5],
        "SibSp": [1, 1, 0, 1, 0, 3, 0, 1, 0, 0], "Fare": [7.25, 71.3, 7.9, 53.1, 8.05, 21.1, 11.1, 30.1, 16.7, 26.5]})
    sales = pd.DataFrame({"order_id": range(1, 7), "units": [3, 5, 2, 8, 1, 4],
                          "unit_price": [5.0, 6.5, 5.0, 7.0, 9.0, 5.5], "revenue": [15, 32.5, 10, 56, 9, 22]})
    for name, frame in [("titanic", titanic), ("sales", sales)]:
        print(name, classify_all(frame, list(frame.columns)))