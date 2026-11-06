"""KPI selection (DATA side). Candidates -> score -> keep only useful ones.
No colours, no Excel, no drawing."""
import numpy as np
import pandas as pd

MIN_SCORE = 55
MIN_KPIS = 2
MAX_KPIS = 6


def _num(x):
    try:
        f = float(x)
    except (TypeError, ValueError):
        return None
    if np.isnan(f) or np.isinf(f):
        return None
    return round(f, 2)


def _cand(label, value, fmt, score, why, skip="below usefulness threshold",
          sub=None, delta=None, delta_label=None):
    return {"label": label, "value": _num(value), "format": fmt, "sub": sub,
            "delta": _num(delta), "delta_label": delta_label,
            "score": score, "why": why, "skip_reason": skip}


def _top_group(work, dim, measure, agg):
    d = work[dim].astype("object").where(work[dim].notna(), "(blank)").astype(str)
    return work.groupby(d)[measure].agg(agg).dropna().sort_values(ascending=False)


def _budget(measures, ts, dims, labels):
    b = 2 + len(measures) + (1 if ts else 0) + (1 if dims else 0) + (1 if labels else 0)
    return max(MIN_KPIS + 1, min(MAX_KPIS, b))


def select_kpis(work, measures, labels, texts, dims, ts, main, charts, agg_for):
    """-> (chosen KPI list, skipped list with reasons)."""
    n = len(work)
    c = []
    main_agg = agg_for(main) if main else None

    # 1. main measure (always the hero KPI)
    if main:
        s = work[main].dropna()
        val = s.sum() if main_agg == "sum" else s.mean()
        delta = ts["growth"] * 100 if ts and ts.get("growth") is not None else None
        c.append(_cand(("Total " if main_agg == "sum" else "Average ") + str(main), val,
                       "compact", 100, "main measure, the number this dashboard is about",
                       delta=delta, delta_label=ts["growth_basis"] if delta is not None else None))
        if main_agg == "sum" and len(s) > 1:
            mean = s.mean()
            cv = float(s.std() / abs(mean)) if mean else 0.0
            cv = 0.0 if np.isnan(cv) else cv
            c.append(_cand(f"Average {main}", mean, "compact", 72 if cv >= 0.3 else 45,
                           f"values vary (cv={cv:.2f}), average per record is informative",
                           skip=f"values are almost the same (cv={cv:.2f}), average adds nothing"))
        elif main_agg == "mean":
            c.append(_cand(f"Highest {main}", s.max(), "compact", 50, "upper range",
                           skip="range is already visible in the charts"))

    # 2. best period (only with a real time series)
    if ts and ts["periods"] >= 6:
        c.append(_cand(f"Best {ts['grain']}", ts["peak"]["value"], "compact", 60,
                       f"peak {ts['grain']} of a {ts['periods']}-period trend", sub=ts["peak"]["label"]))
    elif ts:
        c.append(_cand(f"Best {ts['grain']}", ts["peak"]["value"], "compact", 30,
                       "peak period", skip="too few periods for a meaningful peak", sub=ts["peak"]["label"]))

    # 3. leader of a dimension that is also drawn as a chart
    dim = next((ch["dim"] for ch in charts if ch.get("dim")), None)
    if main and dim:
        g = _top_group(work, dim, main, main_agg)
        k = len(g)
        if main_agg == "sum" and k >= 1 and g.sum() > 0:
            share = float(g.iloc[0] / g.sum())
            lead = share * k
            if k < 3:
                c.append(_cand(f"Top {dim}", share * 100, "percent", 45, "leader share",
                               skip=f"only {k} categories, the chart already shows it", sub=str(g.index[0])))
            elif share >= 0.25 and lead >= 1.5:
                c.append(_cand(f"Top {dim}", share * 100, "percent", 80,
                               f"clear leader: {share * 100:.0f}% across {k} categories", sub=str(g.index[0])))
            else:
                c.append(_cand(f"Top {dim}", share * 100, "percent", 40, "leader share",
                               skip=f"no clear leader (top is only {share * 100:.0f}% of {k})", sub=str(g.index[0])))
        elif k >= 3:
            c.append(_cand(f"Best {dim}", g.iloc[0], "compact", 60, "highest average category",
                           sub=str(g.index[0])))
    elif not main and dims:
        c.append(_cand(f"Unique {dims[0]}", work[dims[0]].nunique(), "int", 55,
                       "how many distinct categories", sub=None))

    # 4. other measures (skip those that just echo the main one)
    extra = [m for m in measures if m != main][:2]
    for i, m in enumerate(extra):
        a = agg_for(m)
        s = work[m].dropna()
        if s.empty:
            continue
        v = s.sum() if a == "sum" else s.mean()
        r = 0.0
        if main:
            r = work[[main, m]].dropna().corr().iloc[0, 1]
            r = 0.0 if pd.isna(r) else float(r)
        label = ("Total " if a == "sum" else "Average ") + str(m)
        if abs(r) > 0.9:
            c.append(_cand(label, v, "compact", 35, "second measure",
                           skip=f"moves with {main} (r={r:.2f}), adds nothing new"))
        else:
            c.append(_cand(label, v, "compact", 66 if i == 0 else 52,
                           "independent second measure",
                           skip="third-priority measure, no space"))

    # 5. label columns: ONE class (the minority) is enough
    for lab in labels[:2]:
        vc = work[lab].astype(str).value_counts()
        if len(vc) < 2:
            continue
        name, cnt = vc.index[-1], int(vc.iloc[-1])
        c.append(_cand(f"{name} share", cnt / vc.sum() * 100, "percent",
                       88 if not main else 62, f"split of {lab}; the other class is just 100% minus this",
                       sub=f"{cnt:,} of {int(vc.sum()):,}"))

    # 6. record count
    c.append(_cand("Records", n, "int", 90 if not main else 58,
                   "size of the dataset", skip="not useful here"))

    # 7. data quality, only when it is a real problem
    blank = float(work.isna().mean().mean() * 100) if n else 0.0
    c.append(_cand("Missing data", blank, "percent", 70 if blank > 5 else 20,
                   f"{blank:.0f}% of cells are empty, the reader should know",
                   skip="data is clean, no need to show"))

    # ---- choose ----
    budget = _budget(measures, ts, dims, labels)
    c.sort(key=lambda k: -k["score"])
    chosen = [k for k in c if k["score"] >= MIN_SCORE][:budget]
    if len(chosen) < MIN_KPIS:
        for k in c:
            if k not in chosen:
                chosen.append(k)
            if len(chosen) >= MIN_KPIS:
                break
    skipped = []
    for k in c:
        if k not in chosen:
            reason = k["skip_reason"] if k["score"] < MIN_SCORE else f"over budget ({budget} KPIs for this dataset)"
            skipped.append({"label": k["label"], "score": k["score"], "why": reason})
    for k in chosen:
        k.pop("skip_reason", None)
    return chosen, skipped