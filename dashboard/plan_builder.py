"""DATA brain: dataset (DataFrame) -> plan JSON.
Decides columns, KPIs, charts, sizes, positions, insights, validation.
No colours, no Excel, no drawing.

Chart choice (v2): every possible chart is a CANDIDATE with a score 0-100 that comes from the
data itself (shares, concentration, correlation, skew, trend length). The best, most varied
2-5 charts are picked (budget depends on how rich the dataset is). Layout depends on the
final chart count (landscape). Every skipped chart gets a reason in plan["charts_skipped"]."""
from itertools import combinations

import numpy as np
import pandas as pd

from dashboard.time_analysis import analyse_time, to_datetime_safe

MAX_KPIS = 6
MAX_CHARTS = 5          # hard upper limit (budget decides 2..5)
MIN_CHARTS = 2
MIN_SCORE = 50          # below this a chart is skipped (unless we have fewer than MIN_CHARTS)
GRID_COLS = 12
BANNER_H, KPI_H, CHART_H, INSIGHT_H = 1, 2, 4, 2
SIZES = {"wide": 8, "medium": 6, "small": 4}   # kept for compatibility

_BUSINESS = ("revenue", "sales", "amount", "total", "profit", "income", "price",
             "cost", "spend", "salary", "fare", "quantity", "units", "value", "score")
_AVG_WORDS = ("price", "rate", "score", "age", "ratio", "percent", "pct",
              "temperature", "rating", "fare", "salary")
_FAMILY = {"pie": "share", "donut": "share", "line": "trend", "area": "trend",
           "bar": "compare", "hbar": "compare", "scatter": "scatter",
           "histogram": "distribution"}


# ---------- helpers ----------
def _py(x):
    if x is None:
        return None
    try:
        f = float(x)
    except (TypeError, ValueError):
        return x
    if np.isnan(f) or np.isinf(f):
        return None
    return round(f, 2)


def _agg_for(name):
    low = str(name).lower()
    return "mean" if any(w in low for w in _AVG_WORDS) else "sum"


def _is_id_name(name):
    n = str(name).strip()
    low = n.lower()
    return (low in ("id", "index", "uuid", "guid", "sr no", "sr.no", "s.no", "sno", "s no")
            or low.endswith("_id") or low.endswith(" id") or low.startswith("id_")
            or low.startswith("serial")
            or (len(n) > 2 and (n.endswith("Id") or n.endswith("ID"))))


def _is_integer_like(s):
    s = s.dropna()
    return len(s) > 0 and bool((s % 1 == 0).all())


# ---------- column roles ----------
def _classify(s, name, n, nun, blank):
    low = str(name).lower()
    if blank > 60:
        return "dropped", f"{blank:.0f}% blank", None
    if nun <= 1:
        return "dropped", "constant value", None
    if pd.api.types.is_bool_dtype(s):
        return "label", "true/false flag", None
    if pd.api.types.is_datetime64_any_dtype(s):
        return "date", "date column", s
    if pd.api.types.is_numeric_dtype(s):
        if nun == 2:
            return "label", "two-value flag", None
        if _is_id_name(name):
            return "id", "id-like name", None
        if any(w in low for w in _BUSINESS):
            return "measure", "business measure", None
        integer = _is_integer_like(s)
        if integer and n > 20 and nun / n > 0.98:
            return "id", "unique running number", None
        if integer and nun <= 8:
            return "dimension", "few distinct values", None
        if "year" in low and integer and nun <= 40:
            return "dimension", "year", None
        return "measure", "numeric", None
    dt = to_datetime_safe(s)
    if dt is not None:
        return "date", "date-like text", dt
    txt = s.dropna().astype(str)
    avg_len = float(txt.str.len().mean())
    ratio = nun / max(n, 1)
    if nun == 2:
        return "label", "two-value text", None
    if _is_id_name(name) or (n > 20 and ratio > 0.95 and avg_len <= 20):
        return "id", "unique identifier", None
    if avg_len >= 40 or (ratio > 0.8 and avg_len >= 20):
        return "text", "free text", None
    if nun <= 50 or ratio <= 0.5:
        return "dimension", "category", None
    return "dropped", "too many unique values", None


def profile_columns(df):
    """-> (profile list, working df with date columns converted)."""
    n = len(df)
    work = df.copy()
    profile = []
    for col in df.columns:
        s = df[col]
        nun = int(s.nunique(dropna=True))
        blank = round(float(s.isna().mean() * 100), 1) if n else 100.0
        role, why, converted = _classify(s, col, n, nun, blank)
        if converted is not None:
            work[col] = converted
        profile.append({"name": col, "role": role, "blank_pct": blank,
                        "unique": nun, "why": why})
    return profile, work


def _rank_measures(df, cols):
    scored = []
    for c in cols:
        s = df[c].dropna()
        if s.empty:
            continue
        mean = s.mean()
        cv = (s.std() / abs(mean)) if mean else 0
        cv = 0 if pd.isna(cv) else cv
        bw = 1 if any(w in str(c).lower() for w in _BUSINESS) else 0
        scored.append((2 * bw + min(cv, 3) / 3, c))
    scored.sort(key=lambda t: -t[0])
    return [c for _, c in scored]


def _rank_dims(df, cols):
    def key(c):
        k = df[c].nunique(dropna=True)
        return (0 if 3 <= k <= 12 else 1, k)
    return sorted(cols, key=key)


def _group(df, dim, measure, agg):
    d = df[dim].astype("object").where(df[dim].notna(), "(blank)").astype(str)
    if measure:
        g = df.groupby(d)[measure].agg(agg)
    else:
        g = d.value_counts()
    g = g.dropna().sort_values(ascending=False)
    return [str(i) for i in g.index], np.array(g.values, dtype=float)


def _hist(values, bins=10):
    counts, edges = np.histogram(values, bins=bins)
    labels = [f"{edges[i]:.4g}-{edges[i + 1]:.4g}" for i in range(len(counts))]
    return labels, [int(c) for c in counts]


def _kpi(label, value, fmt, why, sub=None, delta=None, delta_label=None):
    return {"label": label, "value": _py(value), "format": fmt, "sub": sub,
            "delta": _py(delta), "delta_label": delta_label, "why": why}


# ---------- layout (depends on the FINAL chart count; landscape) ----------
def _put(c, row, col, w, h):
    c["pos"] = {"row": row, "col": col, "w": w, "h": h}
    c["size"] = "wide" if w >= 8 else "medium" if w >= 6 else "small"


def _row(items, row, h):
    """Equal-width row filling the 12 columns."""
    k = len(items)
    w = GRID_COLS // k
    for i, c in enumerate(items):
        _put(c, row, i * w, w if i < k - 1 else GRID_COLS - w * (k - 1), h)


def _pack(charts, start_row):
    """Place charts by count. Returns the first free row after the charts.
    1: one big | 2: side by side | 3: big left + 2 stacked right | 4: 2x2 | 5+: [8|4] then rows of 3."""
    n = len(charts)
    H = CHART_H
    r = start_row
    if n == 0:
        return r
    if n == 1:
        _put(charts[0], r, 0, 12, 6)
        return r + 6
    if n == 2:
        _row(charts, r, 6)
        return r + 6
    if n == 3:
        _put(charts[0], r, 0, 8, 2 * H)
        _put(charts[1], r, 8, 4, H)
        _put(charts[2], r + H, 8, 4, H)
        return r + 2 * H
    if n == 4:
        _row(charts[:2], r, H)
        _row(charts[2:], r + H, H)
        return r + 2 * H
    _put(charts[0], r, 0, 8, H)
    _put(charts[1], r, 8, 4, H)
    r += H
    rest = charts[2:]
    for i in range(0, len(rest), 3):
        _row(rest[i:i + 3], r, H)
        r += H
    return r


# ---------- chart selection ----------
def _budget(n_rows, usable):
    b = 2 if usable <= 2 else 3 if usable <= 3 else 4 if usable <= 5 else 5
    if n_rows < 30:
        b = min(b, 3)
    return max(MIN_CHARTS, min(MAX_CHARTS, b))


def _select(cands, budget):
    """Greedy: highest adjusted score first. Same chart family / same dimension is penalised
    so the dashboard stays varied. Returns (chosen, skipped)."""
    chosen, pool = [], list(cands)

    def adj(c):
        s = c["score"]
        s -= 14 * sum(1 for x in chosen if _FAMILY.get(x["type"]) == _FAMILY.get(c["type"]))
        if c.get("dim") and any(x.get("dim") == c["dim"] for x in chosen):
            s -= 20
        return s

    while pool and len(chosen) < budget:
        best = max(pool, key=adj)
        a = adj(best)
        if a < MIN_SCORE and len(chosen) >= MIN_CHARTS:
            break
        best["adj"] = a
        chosen.append(best)
        pool.remove(best)
    skipped = []
    for c in pool:
        why = ("budget reached" if len(chosen) >= budget else f"score {adj(c):.0f} below {MIN_SCORE}")
        skipped.append({"title": c["title"], "type": c["type"], "score": round(c["score"]), "why": why})
    return chosen, skipped


# ---------- main ----------
def build_plan(df, title=None, domain=None, style_request=None):
    n = len(df)
    style_request = style_request or {"theme": None, "variant": None, "color": None}
    domain = domain or "general"
    base = {"title": title or (f"{domain.title()} Dashboard" if domain != "general" else "Data Dashboard"),
            "domain": domain, "grid": {"cols": GRID_COLS}, "style_request": style_request,
            "kpis": [], "charts": [], "charts_skipped": [], "insights": [], "dropped_columns": [],
            "warnings": [], "validation": {"ok": True, "checks": []}}
    if n == 0 or df.shape[1] == 0:
        base["warnings"].append("empty dataset")
        return base

    profile, work = profile_columns(df)
    roles = {}
    for p in profile:
        roles.setdefault(p["role"], []).append(p["name"])
    measures = _rank_measures(work, roles.get("measure", []))
    dims = _rank_dims(work, roles.get("dimension", []))
    labels = roles.get("label", [])
    texts = roles.get("text", [])
    dates = roles.get("date", [])
    main = measures[0] if measures else None
    main_agg = _agg_for(main) if main else None

    ts, date_col = None, None
    if dates:
        date_col = max(dates, key=lambda c: work[c].nunique())
        ts = analyse_time(work[date_col], work[main] if main else None, agg=main_agg or "sum")

    # ----- KPIs -----
    kpis = []
    if main:
        val = work[main].sum() if main_agg == "sum" else work[main].mean()
        delta = ts["growth"] * 100 if ts and ts["growth"] is not None else None
        kpis.append(_kpi(("Total " if main_agg == "sum" else "Average ") + str(main), val,
                         "compact", "main measure (highest importance score)", delta=delta,
                         delta_label=ts["growth_basis"] if delta is not None else None))
        if main_agg == "sum":
            kpis.append(_kpi(f"Average {main}", work[main].mean(), "compact", "typical value per record"))
        else:
            kpis.append(_kpi(f"Highest {main}", work[main].max(), "compact", "upper range"))
    kpis.append(_kpi("Records", n, "int", "row count"))
    if main and dims:
        cats, vals = _group(work, dims[0], main, main_agg)
        if len(cats) >= 2:
            if main_agg == "sum" and vals.sum() > 0:
                kpis.append(_kpi(f"Top {dims[0]}", vals[0] / vals.sum() * 100, "percent",
                                 "largest share of main measure", sub=cats[0]))
            else:
                kpis.append(_kpi(f"Best {dims[0]}", vals[0], "compact", "highest value", sub=cats[0]))
    for m in measures[1:3]:
        a = _agg_for(m)
        v = work[m].sum() if a == "sum" else work[m].mean()
        kpis.append(_kpi(("Total " if a == "sum" else "Average ") + str(m), v, "compact", "second measure"))
    if not main:
        for lab in labels[:1]:
            vc = work[lab].astype(str).value_counts()
            for name, cnt in list(vc.items())[:2]:
                kpis.append(_kpi(f"{name} share", cnt / max(vc.sum(), 1) * 100, "percent",
                                 f"split of {lab}", sub=f"{int(cnt):,} records"))
        for d in dims[:1]:
            kpis.append(_kpi(f"Unique {d}", work[d].nunique(), "int", "distinct categories"))
        for t in texts[:1]:
            kpis.append(_kpi(f"Avg {t} length", work[t].dropna().astype(str).str.len().mean(),
                             "int", "text size"))
    kpis = kpis[:MAX_KPIS]

    # ----- chart CANDIDATES (each with a data-driven score) -----
    cands, groups = [], {}

    def cand(score, ctype, ttl, cats, series, why, **extra):
        c = {"score": float(score), "type": ctype, "title": ttl, "categories": cats,
             "series": series, "why": why}
        c.update(extra)
        cands.append(c)
        return c

    # 1) trend over time
    if ts and ts["periods"] >= 3:
        what = main or "Records"
        extra = {"note": f"last {ts['grain']} is partial"} if ts["partial_last"] else {}
        kind = "area" if (ts["periods"] >= 24 and main_agg != "mean") else "line"
        score = 92 if ts["periods"] >= 6 else 70
        cand(score, kind, f"{what} by {ts['grain']}", ts["labels"], {what: ts["values"]},
             f"{date_col} + {what} = trend over {ts['periods']} {ts['grain']}s", **extra)

    # 2) one candidate per category column
    additive = (main is None) or main_agg == "sum"
    sname = str(main) if main else "Count"
    prefix = "Average " if (main and main_agg == "mean") else ""
    for dim in dims[:5]:
        cats, vals = _group(work, dim, main, main_agg)
        k = len(cats)
        if k < 2:
            continue
        total = float(vals.sum())
        top = (vals[0] / total) if total > 0 else 0.0
        score = 66 + (8 if 3 <= k <= 8 else 0) - (10 if k > 12 else 0) - (4 if not main else 0)
        if top >= 0.30:
            score += 10
        if total > 0 and top <= (1.0 / k) * 1.15:
            score -= 12                      # almost flat = not interesting
        ttl_core = f"{prefix}{main} by {dim}" if main else f"{dim} count"
        if additive and k <= 5 and vals.min() >= 0 and total > 0:
            kind = "pie" if k <= 3 else "donut"
            c = cand(score, kind, f"{main or 'Records'} share by {dim}", cats,
                     {sname: [_py(v) for v in vals]}, f"{k} categories, shares of a whole",
                     dim=dim, center={"value": _py(total), "label": "Total"})
        elif k <= 8:
            c = cand(score, "bar", ttl_core, cats, {sname: [_py(v) for v in vals]},
                     f"{k} categories, sorted comparison", dim=dim)
        else:
            c = cand(score, "hbar", ttl_core, cats[:10], {sname: [_py(v) for v in vals[:10]]},
                     f"{k} categories, top 10 only", dim=dim, note=f"Top 10 of {k}")
        groups[c["title"]] = (dim, cats, vals)

    # 3) label (flag / class) split
    for lab in labels[:2]:
        vc = work[lab].astype(str).value_counts()
        if len(vc) < 2:
            continue
        cand(78 if not main else 62, "donut", f"{lab} split", [str(i) for i in vc.index],
             {"Count": [int(v) for v in vc.values]}, "label column = count split",
             dim=f"label:{lab}", center={"value": int(vc.sum()), "label": "Records"})

    # 4) correlation between two measures
    if len(measures) >= 2:
        best = None
        for a, b in combinations(measures[:4], 2):
            sub = work[[a, b]].dropna()
            if len(sub) < 10:
                continue
            r = sub[a].corr(sub[b])
            if pd.notna(r) and abs(r) >= 0.5 and (best is None or abs(r) > abs(best[0])):
                best = (r, a, b)
        if best:
            r, a, b = best
            sub = work[[a, b]].dropna()
            sub = sub.sample(min(len(sub), 300), random_state=0)
            cand(min(95, 40 + 60 * abs(r)), "scatter", f"{b} vs {a}", sub[a].round(2).tolist(),
                 {str(b): sub[b].round(2).tolist()}, f"correlation r={r:.2f}",
                 x_name=str(a), corr=round(float(r), 2))

    # 5) distribution of the main measure
    if main and work[main].dropna().nunique() > 5:
        s_main = work[main].dropna()
        skew = float(s_main.skew()) if len(s_main) > 2 else 0.0
        lab_, cnt_ = _hist(s_main.values)
        cand(62 if abs(skew) >= 1 else 48, "histogram", f"{main} distribution", lab_,
             {"Count": cnt_}, f"spread of {main} (skew {skew:.1f})")

    # 6) text length distribution
    for t in texts[:1]:
        lens = work[t].dropna().astype(str).str.len().values
        if len(lens) >= 5:
            lab_, cnt_ = _hist(lens)
            cand(58, "histogram", f"{t} length (characters)", lab_, {"Count": cnt_},
                 "text column = length distribution")

    # ----- choose, order, build final charts -----
    usable = len(measures) + len(dims) + len(labels) + (1 if ts else 0) + (1 if texts else 0)
    budget = _budget(n, usable)
    chosen, skipped = _select(cands, budget)
    order = {"trend": 0, "compare": 1}
    chosen.sort(key=lambda c: (order.get(_FAMILY.get(c["type"]), 2), -c.get("adj", c["score"])))

    charts, checks = [], []
    for i, c in enumerate(chosen):
        item = {k: v for k, v in c.items() if k not in ("adj",)}
        item["priority"] = i + 1
        item["score"] = round(c["score"])
        item["size"] = "medium"
        charts.append(item)
        if main and main_agg == "sum":
            if c["title"] in groups:
                got = float(groups[c["title"]][2].sum())
                exp = float(work[main].sum())
                checks.append({"chart": c["title"], "chart_total": _py(got), "source_total": _py(exp),
                               "ok": abs(got - exp) <= max(0.01, abs(exp) * 1e-6)})
            elif ts and c["type"] in ("line", "area"):
                exp = float(work.loc[work[date_col].notna(), main].sum())
                got = float(sum(ts["values"]))
                checks.append({"chart": c["title"], "chart_total": _py(got), "source_total": _py(exp),
                               "ok": abs(got - exp) <= max(0.01, abs(exp) * 1e-6)})
    base["charts_skipped"] = skipped
    if not charts:
        base["warnings"].append("no chartable columns found")

    # ----- insights -----
    ins = []
    first_dim = next((groups[c["title"]] for c in chosen if c["title"] in groups), None)
    if main and first_dim:
        dim, cats, vals = first_dim
        if main_agg == "sum" and vals.sum() > 0:
            ins.append(f"{cats[0]} leads {main} with {vals[0] / vals.sum() * 100:.0f}%")
            if len(cats) >= 6:
                ins.append(f"Top 3 {dim} = {vals[:3].sum() / vals.sum() * 100:.0f}% of {main}")
        else:
            ins.append(f"{cats[0]} has the highest {main} ({vals[0]:,.2f})")
    if ts and ts["growth"] is not None:
        what = main or "Records"
        ins.append(f"{what} is {'up' if ts['growth'] >= 0 else 'down'} "
                   f"{abs(ts['growth']) * 100:.0f}% ({ts['growth_basis']})")
        if ts["partial_last"]:
            ins.append(f"Latest {ts['grain']} is partial, so it is left out of growth")
    for lab in labels[:1]:
        vc = work[lab].astype(str).value_counts()
        if len(vc) >= 2:
            ins.append(f"{vc.index[0]} is {vc.iloc[0] / vc.sum() * 100:.0f}% of {lab}")
    ins = ins[:4]

    # ----- layout (by final counts) -----
    k = max(len(kpis), 1)
    w = GRID_COLS // k
    for i, kp in enumerate(kpis):
        kp["priority"] = i + 1
        kp["pos"] = {"row": BANNER_H, "col": i * w, "h": KPI_H,
                     "w": w if i < k - 1 else GRID_COLS - w * (k - 1)}
    end_row = _pack(charts, BANNER_H + KPI_H)
    insights_pos = {"row": end_row, "col": 0, "w": GRID_COLS, "h": INSIGHT_H} if ins else None

    base.update({
        "subtitle": f"{n:,} records · {df.shape[1]} columns" + (f" · {ts['start']} to {ts['end']}" if ts else ""),
        "banner": {"row": 0, "col": 0, "w": GRID_COLS, "h": BANNER_H},
        "kpis": kpis, "charts": charts, "insights": ins, "insights_pos": insights_pos,
        "dropped_columns": [{"col": str(p["name"]), "role": p["role"], "why": p["why"]}
                            for p in profile if p["role"] in ("id", "text", "dropped")],
        "columns": [{"col": str(p["name"]), "role": p["role"], "why": p["why"]} for p in profile],
        "main_measure": str(main) if main else None,
        "date_column": str(date_col) if date_col else None,
        "validation": {"ok": all(c["ok"] for c in checks), "checks": checks},
        "meta": {"rows": n, "columns": int(df.shape[1]), "chart_budget": budget},
    })
    return base