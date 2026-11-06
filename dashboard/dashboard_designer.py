"""Dashboard designer: decides the LOOK of a dashboard (theme + layout). No Excel access (testable).

- pick_design(): theme + layout per dataset (deterministic by dataset fingerprint,
  biased by domain / data shape). regenerate=True gives the next UNUSED combination
  for the same dataset, remembered in ~/.auto_dashboard_history.json.
- build_insights(): short data-driven insight cards (top share, trend, concentration).
Geometry lives in dashboard_layout.py, constants in dashboard_config.py.
Test: python -m dashboard.dashboard_designer
"""
from __future__ import annotations

import hashlib
import json
import random
from typing import Any

import pandas as pd

from dashboard.dashboard_config import GRID, GRID_LEFT, HISTORY_FILE, MAX_INSIGHTS  # noqa: F401
from dashboard.dashboard_layout import LAYOUTS, order_charts, plan_layout  # noqa: F401
from dashboard.dashboard_styler import THEMES
from dashboard.measure import classify_measure

DOMAIN_THEMES = {
    "Sales": ["ocean_blue", "neon_purple", "cyber_teal"],
    "HR": ["rose_noir", "neon_purple", "aurora"],
    "Education": ["sunset", "aurora", "ocean_blue"],
    "Finance": ["emerald", "midnight_gold", "ocean_blue"],
    "Inventory": ["cyber_teal", "midnight_gold", "sunset"],
    "Healthcare": ["cyber_teal", "emerald", "aurora"],
    "Passengers": ["midnight_gold", "ocean_blue", "rose_noir"],
    "Survey": ["aurora", "sunset", "neon_purple"],
    "Customer": ["neon_purple", "rose_noir", "cyber_teal"],
    "Marketing": ["sunset", "neon_purple", "midnight_gold"],
    "Transport": ["cyber_teal", "ocean_blue", "emerald"],
}


def _pretty(name: str) -> str:
    return str(name).replace("_", " ").title()


# ---------- design choice ----------
def fingerprint(df: pd.DataFrame) -> str:
    sig = "|".join(f"{c}:{df[c].dtype}" for c in df.columns) + f"|{len(df)}"
    return hashlib.md5(sig.encode()).hexdigest()[:10]


def _load_history() -> dict[str, list]:
    try:
        return json.loads(HISTORY_FILE.read_text())
    except Exception:
        return {}


def _save_history(hist: dict[str, list]) -> None:
    try:
        HISTORY_FILE.write_text(json.dumps(hist))
    except Exception:
        pass


def _ranked_pairs(fp: str, domain: str, has_trend: bool) -> list[tuple[str, str]]:
    rng = random.Random(fp)
    pairs = [(t, l) for t in THEMES for l in LAYOUTS]
    rng.shuffle(pairs)
    pool = DOMAIN_THEMES.get(domain, [])
    suited = ["hero", "command"] if has_trend else ["magazine", "sidebar", "command"]
    pairs.sort(key=lambda p: (p[0] not in pool, p[1] not in suited))   # stable sort
    return pairs


def pick_design(df: pd.DataFrame, domain: str, has_trend: bool, theme: str | None = None,
                layout: str | None = None, regenerate: bool = False) -> dict[str, Any]:
    if theme and theme not in THEMES:
        raise ValueError(f"Theme '{theme}' not found. Available: {sorted(THEMES)}")
    if layout and layout not in LAYOUTS:
        raise ValueError(f"Layout '{layout}' not found. Available: {LAYOUTS}")
    fp = fingerprint(df)
    hist = _load_history()
    used: list = hist.get(fp, [])
    if used and not regenerate and not (theme or layout):
        choice = tuple(used[-1])
    else:
        pairs = _ranked_pairs(fp, domain, has_trend)
        fresh = [p for p in pairs if list(p) not in used]
        if not fresh:                       # everything tried -> start over
            used, fresh = [], pairs
        if used and regenerate:             # prefer a different theme AND a different layout
            last = used[-1]
            fresh.sort(key=lambda p: (p[0] == last[0], p[1] == last[1]))
        choice = fresh[0]
        if theme or layout:
            choice = (theme or choice[0], layout or choice[1])
        if list(choice) not in used:
            used.append(list(choice))
    hist[fp] = used
    _save_history(hist)
    return {"theme": choice[0], "layout": choice[1], "variant": len(used), "fingerprint": fp}


# ---------- insights ----------
def build_insights(df: pd.DataFrame, charts: list[dict[str, Any]],
                   max_items: int = MAX_INSIGHTS) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    cats = [c for c in charts if c["type"] in {"bar", "pie", "stacked"}]
    lines = [c for c in charts if c["type"] == "line"]

    if cats:
        p = cats[0]
        g = df.groupby(p["category"])[p["measure"]].agg(p["agg"]).sort_values(ascending=False)
        if len(g):
            top = str(g.index[0])
            if p["agg"] == "sum" and g.sum() > 0:
                share = float(g.iloc[0] / g.sum())
                out.append({"title": f"Top {_pretty(p['category'])}", "headline": f"{top} leads with {share:.0%}",
                            "detail": f"Share of total {_pretty(p['measure'])}", "pct": share})
                if len(g) >= 5:
                    c3 = float(g.head(3).sum() / g.sum())
                    out.append({"title": "Concentration", "headline": f"Top 3 = {c3:.0%} of total",
                                "detail": f"{_pretty(p['category'])} dependency", "pct": c3})
            else:
                rate = p["agg"] == "mean" and classify_measure(df, p["measure"]) == "binary"
                value = f"{g.iloc[0]:.1%}" if rate else f"{g.iloc[0]:,.2f}"
                label = f"{_pretty(p['measure'])} rate" if rate else f"{p['agg'].title()} {_pretty(p['measure'])}"
                out.append({"title": f"Best {_pretty(p['category'])}", "headline": f"{top} is highest",
                            "detail": f"{label}: {value}", "pct": float(g.iloc[0]) if rate else None})

    if lines:
        p = lines[0]
        key = df[p["category"]]
        if pd.api.types.is_datetime64_any_dtype(key):
            key = key.dt.to_period("M")
        s = df.groupby(key)[p["measure"]].agg(p["agg"]).sort_index()
        if len(s) >= 2 and s.iloc[-2] not in (0, None):
            change = float((s.iloc[-1] - s.iloc[-2]) / abs(s.iloc[-2]))
            arrow = "▲" if change >= 0 else "▼"
            out.insert(1 if out else 0, {"title": "Latest Trend", "headline": f"{arrow} {abs(change):.1%} vs previous",
                                         "detail": f"{_pretty(p['measure'])}, last period", "pct": min(abs(change), 1.0)})
        if len(s) >= 3:
            out.append({"title": "Peak Period", "headline": str(s.idxmax()),
                        "detail": f"Highest {_pretty(p['measure'])}: {s.max():,.2f}", "pct": None})
    return out[:max_items]


if __name__ == "__main__":
    import numpy as np
    rng = np.random.default_rng(1)
    demo = pd.DataFrame({"region": rng.choice(list("NSEWC"), 60),
                         "date": pd.date_range("2025-01-01", periods=60, freq="7D"),
                         "revenue": rng.integers(100, 900, 60)})
    for regen in (False, False, True, True):
        d = pick_design(demo, "Sales", True, regenerate=regen)
        print(f"regenerate={regen!s:<5} ->", d["theme"], d["layout"], "variant", d["variant"])