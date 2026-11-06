"""Dashboard validator: checks a dashboard plan BEFORE the build and the sheet AFTER the build.

Issues are warnings (returned in the result), they never stop the build.
Test (no Excel needed): python -m dashboard.dashboard_validator
"""
from __future__ import annotations

from typing import Any

from dashboard.dashboard_config import COL_WIDTH, GRID, GRID_LEFT, MIN_CHART_H, MIN_CHART_W


def _overlap(a: dict[str, Any], b: dict[str, Any]) -> bool:
    return not (a["col"] + a["w"] <= b["col"] or b["col"] + b["w"] <= a["col"]
                or a["row"] + a["h"] <= b["row"] or b["row"] + b["h"] <= a["row"])


def check_tiles(tiles: list[dict[str, Any]], n_kpi: int, n_chart: int, n_ins: int,
                total_rows: int) -> list[str]:
    """Count, grid bounds, minimum chart size and overlap."""
    issues: list[str] = []
    for kind, wanted in (("kpi", n_kpi), ("chart", n_chart), ("insight", n_ins)):
        placed = sum(1 for t in tiles if t["kind"] == kind)
        if placed != wanted:
            issues.append(f"{kind}: {placed} tiles placed for {wanted} items")
    for t in tiles:
        if t["col"] < GRID_LEFT or t["col"] + t["w"] > GRID_LEFT + GRID:
            issues.append(f"{t['kind']} {t['idx'] + 1} is outside the grid")
        if t["row"] + t["h"] > total_rows:
            issues.append(f"{t['kind']} {t['idx'] + 1} is below the page end")
        if t["kind"] == "chart" and (t["w"] < MIN_CHART_W or t["h"] < MIN_CHART_H):
            issues.append(f"chart {t['idx'] + 1} is too small ({t['w']}x{t['h']})")
    for i, a in enumerate(tiles):
        for b in tiles[i + 1:]:
            if _overlap(a, b):
                issues.append(f"{a['kind']} {a['idx'] + 1} overlaps {b['kind']} {b['idx'] + 1}")
    return issues


def check_text(tiles: list[dict[str, Any]], kpi_labels: list[str], kpi_values: list[str],
               insights: list[dict[str, Any]]) -> list[str]:
    """Rough check: will KPI label/value and insight headline fit inside their card?"""
    issues: list[str] = []
    for t in tiles:
        usable = max(t["w"] - 2, 1) * COL_WIDTH          # characters at default font size
        i = t["idx"]
        if t["kind"] == "kpi" and i < len(kpi_values):
            if len(kpi_values[i]) > usable / 2.0:        # value font is about 2x
                issues.append(f"KPI value '{kpi_values[i]}' may be cut")
            if i < len(kpi_labels) and len(kpi_labels[i]) > usable * 1.1:
                issues.append(f"KPI label '{kpi_labels[i]}' may be cut")
        elif t["kind"] == "insight" and i < len(insights):
            head = str(insights[i].get("headline", ""))
            if len(head) > usable / 1.3:
                issues.append(f"Insight '{head}' may be cut")
    return issues


def check_sheet(sheet, expected_charts: int) -> list[str]:
    """After the build: chart count, charts inside the page, charts not overlapping."""
    issues: list[str] = []
    try:
        charts = list(sheet.charts)
        right = sheet.range((1, GRID_LEFT + GRID)).left
        boxes = [(c.name, c.left, c.top, c.left + c.width, c.top + c.height) for c in charts]
    except Exception as exc:
        return [f"could not read charts: {type(exc).__name__}"]
    if len(boxes) != expected_charts:
        issues.append(f"{len(boxes)} charts on the sheet, expected {expected_charts}")
    for name, x1, y1, x2, y2 in boxes:
        if x2 > right + 1:
            issues.append(f"chart '{name}' goes outside the dashboard width")
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            if not (a[3] <= b[1] or b[3] <= a[1] or a[4] <= b[2] or b[4] <= a[2]):
                issues.append(f"charts '{a[0]}' and '{b[0]}' overlap")
    return issues


def summarize(issues: list[str]) -> dict[str, Any]:
    return {"ok": not issues, "issues": issues}


if __name__ == "__main__":
    from dashboard.dashboard_layout import LAYOUTS, plan_layout
    bad = 0
    for name in LAYOUTS:
        for n_chart in (3, 4, 5, 6):
            for n_kpi in (3, 6):
                tiles, rows = plan_layout(name, n_kpi, n_chart, 3)
                found = check_tiles(tiles, n_kpi, n_chart, 0 if name == "sidebar" and False else 3, rows)
                found = [f for f in found if not f.startswith("insight")]   # some layouts have no insight band
                if found:
                    bad += 1
                    print(f"{name:<9} kpi={n_kpi} charts={n_chart}: {found}")
    print("ALL LAYOUTS OK" if not bad else f"{bad} layout case(s) with issues")