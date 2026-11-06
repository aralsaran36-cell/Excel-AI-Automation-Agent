"""Date/time analysis (DATA side). No drawing, no Excel."""
import re
import warnings

import numpy as np
import pandas as pd

_DATE_HINT = re.compile(
    r"\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}"
    r"|\d{1,2}\s*[A-Za-z]{3,9}\s*\d{2,4}"
    r"|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{2,4}"
)
_GRAIN_NAME = {"D": "day", "W": "week", "M": "month", "Y": "year"}


def to_datetime_safe(series, min_ratio=0.8):
    """Return a datetime Series if the column is date-like, else None."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    if pd.api.types.is_numeric_dtype(series) or pd.api.types.is_bool_dtype(series):
        return None
    s = series.dropna()
    if s.empty:
        return None
    sample = s.head(50)
    if not all(hasattr(v, "year") and hasattr(v, "month") for v in sample):
        strs = sample.astype(str)
        if strs.map(lambda t: bool(_DATE_HINT.search(t))).mean() < min_ratio:
            return None
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        parsed = pd.to_datetime(series, errors="coerce")
    if parsed.notna().sum() / max(len(s), 1) < min_ratio:
        return None
    return parsed


def _pick_freq(span_days):
    if span_days <= 45:
        return "D"
    if span_days <= 200:
        return "W"
    if span_days <= 365 * 4:
        return "M"
    return "Y"


def _label(period, freq):
    ts = period.to_timestamp()
    if freq == "M":
        return ts.strftime("%b %Y")
    if freq == "Y":
        return ts.strftime("%Y")
    return ts.strftime("%d %b")


def analyse_time(dates, values=None, agg="sum"):
    """Group a measure over time. values=None -> count of records.

    Returns dict (grain, labels, values, partial_last, growth, trend, peak, low) or None.
    growth is a fraction (0.12 = +12%) and ignores a partial last period.
    """
    d = pd.DataFrame({"d": pd.to_datetime(dates, errors="coerce")})
    if values is None:
        d["v"] = 1.0
        agg = "sum"
    else:
        d["v"] = pd.to_numeric(values, errors="coerce")
    d = d.dropna(subset=["d"])
    if d.empty:
        return None

    span_days = int((d["d"].max() - d["d"].min()).days) + 1
    freq = _pick_freq(span_days)
    grouped = d.groupby(d["d"].dt.to_period(freq))["v"].agg(agg)
    full = pd.period_range(grouped.index.min(), grouped.index.max(), freq=freq)
    grouped = grouped.reindex(full)
    if agg in ("sum", "count"):
        grouped = grouped.fillna(0)
    else:
        grouped = grouped.interpolate().fillna(0)

    labels = [_label(p, freq) for p in full]
    vals = [round(float(v), 2) for v in grouped.values]

    last = full[-1]
    partial = freq != "D" and last.end_time.normalize() > d["d"].max().normalize()
    core_vals = vals[:-1] if partial and len(vals) > 1 else vals
    core_labels = labels[:-1] if partial and len(labels) > 1 else labels

    growth, basis = None, None
    if len(core_vals) >= 2 and core_vals[-2] != 0:
        growth = (core_vals[-1] - core_vals[-2]) / abs(core_vals[-2])
        basis = f"{core_labels[-1]} vs {core_labels[-2]}"

    trend = "flat"
    if len(core_vals) >= 3:
        slope = np.polyfit(range(len(core_vals)), core_vals, 1)[0]
        mean = abs(np.mean(core_vals)) or 1.0
        rel = slope * len(core_vals) / mean
        trend = "up" if rel > 0.05 else "down" if rel < -0.05 else "flat"

    peak_i = int(np.argmax(core_vals))
    low_i = int(np.argmin(core_vals))
    return {
        "grain": _GRAIN_NAME[freq], "freq": freq,
        "labels": labels, "values": vals, "periods": len(vals),
        "partial_last": bool(partial),
        "growth": growth, "growth_basis": basis, "trend": trend,
        "peak": {"label": core_labels[peak_i], "value": core_vals[peak_i]},
        "low": {"label": core_labels[low_i], "value": core_vals[low_i]},
        "start": str(d["d"].min().date()), "end": str(d["d"].max().date()),
        "span_days": span_days,
    }