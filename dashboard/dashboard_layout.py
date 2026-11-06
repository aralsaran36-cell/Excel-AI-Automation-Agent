"""Dashboard layout: blueprints and tile planning. No Excel access (testable).

plan_layout() turns a blueprint into tiles (row, col, w, h) on a 24-col grid.
Charts are spread over the bands so there are no lonely wide tiles or empty strips.
"""
from __future__ import annotations

import math
from typing import Any

from dashboard.dashboard_config import GRID, GRID_LEFT, TOP_ROW

# (kind, height_rows, spans). Spans sum to the usable width; fewer items -> spans re-divided.
BLUEPRINTS: dict[str, dict[str, Any]] = {
    "command": {"sidebar": 0, "bands": [
        ("kpi", 4, [6, 6, 6, 6]), ("chart", 15, [16, 8]),
        ("chart", 13, [8, 8, 8]), ("insight", 5, [8, 8, 8])]},
    "hero": {"sidebar": 0, "bands": [
        ("chart", 17, [24]), ("kpi", 4, [4, 4, 4, 4, 4, 4]),
        ("chart", 13, [12, 12]), ("chart", 13, [8, 8, 8]), ("insight", 5, [8, 8, 8])]},
    "sidebar": {"sidebar": 6, "bands": [
        ("chart", 16, [18]), ("chart", 13, [10, 8]), ("chart", 13, [6, 6, 6])]},
    "magazine": {"sidebar": 0, "bands": [
        ("kpi", 4, [8, 8, 8]), ("chart", 14, [12, 12]),
        ("insight", 5, [8, 8, 8]), ("chart", 13, [8, 8, 8])]},
    "split": {"sidebar": 0, "bands": [
        ("kpi", 4, [4, 4, 4, 4, 4, 4]), ("chart", 15, [14, 10]),
        ("insight", 5, [8, 8, 8]), ("chart", 13, [12, 12])]},
    "tiles": {"sidebar": 0, "bands": [
        ("insight", 5, [8, 8, 8]), ("kpi", 4, [6, 6, 6, 6]),
        ("chart", 14, [12, 12]), ("chart", 13, [8, 8, 8])]},
}
LAYOUTS = list(BLUEPRINTS)


def _even(total: int, n: int) -> list[int]:
    base, extra = divmod(total, n)
    return [base + (1 if i < extra else 0) for i in range(n)]


def _split_even(n: int, cap: int) -> list[int]:
    """n items in rows of at most cap, rows as equal as possible (6, cap 4 -> [3, 3])."""
    rows = math.ceil(n / cap)
    base, extra = divmod(n, rows)
    return [base + (1 if i < extra else 0) for i in range(rows)]


def _alloc(n: int, caps: list[int]) -> list[int]:
    """Spread n charts over bands (capacity caps). Uses as few bands as needed, balanced;
    extra charts go to the LAST bands first so the first band stays big/wide."""
    if n <= 0:
        return []
    used, total = 0, 0
    while used < len(caps) and total < n:
        total += caps[used]
        used += 1
    counts = [1] * used
    rem = n - used
    while rem > 0:
        moved = False
        for i in range(used - 1, -1, -1):
            if rem > 0 and counts[i] < caps[i]:
                counts[i] += 1
                rem -= 1
                moved = True
        if not moved:
            break
    return counts


def order_charts(charts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Wide-friendly charts first (trend, stacked), pies last (small slots)."""
    rank = {"line": 0, "stacked": 1, "bar": 2, "column": 2, "hist": 2, "scatter": 3, "pie": 4}
    return sorted(charts, key=lambda c: rank.get(c["type"], 5))


def plan_layout(name: str, n_kpi: int, n_chart: int, n_ins: int) -> tuple[list[dict[str, Any]], int]:
    bp = BLUEPRINTS[name]
    side = bp["sidebar"]
    main_w = GRID - side
    main_left = GRID_LEFT + side
    tiles: list[dict[str, Any]] = []
    side_end = 0

    if side:                                                     # KPIs + insights stacked on the left
        r = TOP_ROW
        for i in range(n_kpi):
            tiles.append({"kind": "kpi", "idx": i, "row": r, "col": GRID_LEFT, "w": side, "h": 4})
            r += 4
        for i in range(n_ins):
            tiles.append({"kind": "insight", "idx": i, "row": r, "col": GRID_LEFT, "w": side, "h": 5})
            r += 5
        side_end = r

    bands = list(bp["bands"])
    chart_bands = [b for b in bands if b[0] == "chart"]
    extra_bands = [("chart", 13, _even(main_w, 3))] * 3          # overflow rows if ever needed
    counts = _alloc(n_chart, [len(b[2]) for b in chart_bands + extra_bands])

    seq: list[tuple[str, int, list[int], int]] = []              # (kind, height, spans, count)
    ci = 0
    for kind, h, spans in bands:
        if kind == "kpi":
            if not side and n_kpi:
                for cnt in _split_even(n_kpi, len(spans)):
                    seq.append(("kpi", h, spans, cnt))
        elif kind == "chart":
            cnt = counts[ci] if ci < len(counts) else 0
            ci += 1
            if cnt:
                seq.append(("chart", h, spans, cnt))
        elif kind == "insight":
            if not side and n_ins:
                seq.append(("insight", h, spans, min(len(spans), n_ins)))
    for k in range(len(chart_bands), len(counts)):
        if counts[k]:
            seq.append(("chart", 13, _even(main_w, 3), counts[k]))

    done = {"kpi": 0, "chart": 0, "insight": 0}
    row = TOP_ROW
    last_chart_row = None
    for kind, h, spans, cnt in seq:
        sp = spans if cnt == len(spans) else _even(sum(spans), cnt)
        col = main_left
        for w in sp:
            tiles.append({"kind": kind, "idx": done[kind], "row": row, "col": col, "w": w, "h": h})
            done[kind] += 1
            col += w
        if kind == "chart":
            last_chart_row = row
        row += h

    if side and last_chart_row is not None and side_end > row:   # fill the strip under the main area
        extra = side_end - row
        for t in tiles:
            if t["kind"] == "chart" and t["row"] == last_chart_row:
                t["h"] += extra
        row = side_end

    total_rows = max(row, side_end) + 1
    return tiles, total_rows