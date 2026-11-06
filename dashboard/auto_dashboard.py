"""Auto Dashboard: analyses the dataset, then lets the designer choose theme + layout.

Every dataset gets its own look. Call again with regenerate=True (same dataset)
to get a NEW theme, layout, KPI set and chart set as a new sheet (old sheets are kept).
Settings are saved in a masked meta cell so refresh/modify can rebuild the same dashboard.

Two engines:
  1. IMAGE engine (default): plan_builder (DATA) -> renderer (DESIGN) -> matplotlib pictures.
  2. NATIVE engine (fallback): the old Excel native charts. Used when the image engine fails
     (traceback saved to dashboard/image_error.log) or USE_IMAGE_DASHBOARD = False.
"""
from __future__ import annotations

import json
import math
from typing import Any

import pandas as pd

from dashboard.chart_selector import auto_charts
from dashboard.dashboard_config import (
    BACK_COL, BANNER_ROW_H, COL_WIDTH, GENERATED_PREFIXES, GRID, GRID_LEFT, HIST_BINS,
    KPI_VALUE_ROW_H, MARGIN_WIDTH, META_COL, META_PREFIX, ROW_H, SCATTER_MAX_POINTS,
)
from dashboard.dashboard_designer import build_insights, pick_design
from dashboard.dashboard_filter import apply_filters
from dashboard.dashboard_layout import order_charts, plan_layout
from dashboard.dashboard_styler import (
    hide_gridlines, mask_range, set_column_widths, style_banner, style_chart,
    style_insight_card, style_kpi_card, style_page, style_sidebar_bg, style_tile,
)
from dashboard.dashboard_validator import check_sheet, check_text, check_tiles, summarize
from dashboard.dataset_analyser import analyse_dataframe, parse_date_columns
from dashboard.domain_detector import detect_domain
from dashboard.kpi_selector import auto_kpis, format_value
from dashboard.measure import classify_measure
from excel_tools.charts import create_excel_chart
from excel_tools.workbook_manager import ExcelToolError, get_sheet, load_dataframe

USE_IMAGE_DASHBOARD = True   # False = old native-chart dashboard only


def _a1(row: int, col: int) -> str:
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def _py(value: Any) -> Any:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if hasattr(value, "item"):
        value = value.item()
    return round(value, 2) if isinstance(value, float) else value


def _is_generated(name: str) -> bool:
    return str(name).lower().startswith(GENERATED_PREFIXES)


def _pick_source(workbook_name: str | None, sheet_name: str | None):
    """Data sheet for the dashboard. If the active/requested sheet is a generated one
    (Dashboard, Dashboard 2, Filtered..., Pivot...), use the first real data sheet."""
    source = get_sheet(workbook_name, sheet_name)
    if not _is_generated(source.name):
        return source
    for s in source.book.sheets:
        if not _is_generated(s.name) and s.used_range.count > 1:
            return s
    raise ExcelToolError("No data sheet found. Open the sheet that has your data.")


def _table(df: pd.DataFrame, plan: dict[str, Any]) -> tuple[list[str], list[list[Any]]]:
    kind, cat, meas, agg = plan["type"], plan["category"], plan["measure"], plan["agg"]
    if kind == "hist":
        s = pd.to_numeric(df[meas], errors="coerce").dropna()
        if s.empty:
            return ["Range", "Count"], []
        counts = pd.cut(s, bins=HIST_BINS).value_counts(sort=False)
        fmt = "{:.1f}" if (s.max() - s.min()) < 10 else "{:.0f}"
        return ["Range", "Count"], [[f"{fmt.format(iv.left)}-{fmt.format(iv.right)}", int(v)]
                                    for iv, v in counts.items()]
    if kind == "scatter":
        sub = df[[meas, plan["measure2"]]].dropna()
        if len(sub) > SCATTER_MAX_POINTS:
            sub = sub.sample(SCATTER_MAX_POINTS, random_state=1)
        return [meas, plan["measure2"]], [[_py(a), _py(b)] for a, b in sub.values.tolist()]
    if kind == "stacked":
        pv = df.pivot_table(index=cat, columns=plan["series"], values=meas, aggfunc=agg, fill_value=0)
        pv = pv.loc[pv.sum(axis=1).sort_values(ascending=False).index]
        header = [cat] + [str(c) for c in pv.columns]
        return header, [[str(i)] + [_py(v) for v in row] for i, row in zip(pv.index, pv.values.tolist())]

    scale = 100 if (agg == "mean" and classify_measure(df, meas) == "binary") else 1   # 0/1 -> percent
    if kind == "line":
        key = df[cat]
        if pd.api.types.is_datetime64_any_dtype(key):
            many = key.nunique() > 24
            key = key.dt.to_period("M").dt.to_timestamp() if many else key
            grouped = df.groupby(key)[meas].agg(agg).sort_index() * scale
            fmt = "%b %Y" if many else "%Y-%m-%d"
            return [cat, meas], [[k.strftime(fmt), _py(v)] for k, v in grouped.items()]
        grouped = df.groupby(key)[meas].agg(agg).sort_index() * scale
        return [cat, meas], [[str(k), _py(v)] for k, v in grouped.items()]
    grouped = df.groupby(cat)[meas].agg(agg).sort_values(ascending=False) * scale
    if plan.get("top_n"):
        grouped = grouped.head(plan["top_n"])
    return [cat, meas], [[str(k), _py(v)] for k, v in grouped.items()]


def _render_type(plan: dict[str, Any], variant: int) -> str:
    """Excel chart type. Odd variants use area / horizontal bars for a different look."""
    t = plan["type"]
    if t == "pie":
        return "doughnut"
    if t == "hist":
        return "column"
    if t == "stacked":
        return "stacked"
    if t == "bar":
        return "bar" if (variant % 2 == 1 and plan.get("categories", 99) <= 8) else "column"
    if t == "line":
        return "area" if variant % 2 == 1 else "line"
    return t


def _new_sheet(book, base: str = "Dashboard"):
    """Never delete an old dashboard: use the next free name (Dashboard, Dashboard 2, ...)."""
    names = {s.name.lower() for s in book.sheets}
    name, i = base, 2
    while name.lower() in names:
        name = f"{base} {i}"
        i += 1
    return book.sheets.add(name, after=book.sheets[-1])


# ----------------------------------------------------------------------------------
# NATIVE engine (old behaviour, unchanged)
# ----------------------------------------------------------------------------------
def build_dashboard(df: pd.DataFrame, kpis: list[dict[str, Any]], charts: list[dict[str, Any]],
                    book, source_name: str, filters_text: str, title: str,
                    design: dict[str, Any], meta: dict[str, Any] | None = None) -> dict[str, Any]:
    theme, layout, variant = design["theme"], design["layout"], design["variant"]
    charts = order_charts(charts)
    insights = build_insights(df, charts)
    tiles, total_rows = plan_layout(layout, len(kpis), len(charts), len(insights))
    total_cols = GRID_LEFT + GRID                      # + 1 right margin

    kpi_values = [format_value(k) for k in kpis]
    warnings = check_tiles(tiles, len(kpis), len(charts), len(insights), total_rows)
    warnings += check_text(tiles, [k["label"] for k in kpis], kpi_values, insights)

    sheet = _new_sheet(book)
    created: list[str] = []
    notes: list[str] = []
    style_report: dict[str, Any] = {}

    try:
        set_column_widths(sheet, 1, 1, MARGIN_WIDTH)
        set_column_widths(sheet, GRID_LEFT, GRID_LEFT + GRID - 1, COL_WIDTH)
        set_column_widths(sheet, GRID_LEFT + GRID, GRID_LEFT + GRID, MARGIN_WIDTH)
        sheet.range((1, 1), (total_rows + 3, 1)).row_height = ROW_H

        style_page(sheet, rows=total_rows + 3, cols=total_cols, theme=theme)
        subtitle = (f"Source: {source_name} | {len(df):,} records | {theme.replace('_', ' ')} / {layout}"
                    f" / v{variant}" + (f" | Filter: {filters_text}" if filters_text else ""))
        style_banner(sheet, total_cols, title, subtitle, theme)
        sheet.range((1, 1)).row_height = BANNER_ROW_H

        if layout == "sidebar":
            style_sidebar_bg(sheet, 3, 1, 7, total_rows, theme)      # sidebar strip incl. margin

        back_row = 2
        back_width = 2
        sheet.range((1, BACK_COL)).value = "Backing data (chart sources)"
        sheet.range((1, BACK_COL)).font.size = 8

        for tile in tiles:
            r, c, w, h, i = tile["row"], tile["col"], tile["w"], tile["h"], tile["idx"]
            if tile["kind"] == "kpi":
                kpi = kpis[i]
                pct = kpi["value"] / 100 if kpi["agg"] == "rate" else None
                style_kpi_card(sheet, r, c, w, h, kpi["label"], kpi_values[i], theme, i, pct)
                sheet.range((r + 1, 1)).row_height = KPI_VALUE_ROW_H
                continue
            if tile["kind"] == "insight":
                style_insight_card(sheet, r, c, w, h, insights[i], theme, i)
                continue

            plan = charts[i]
            header, rows = _table(df, plan)
            if not rows:
                notes.append(f"'{plan['title']}' skipped: no data")
                continue
            style_tile(sheet, r, c, w, h, plan["title"], theme)

            sheet.range((back_row, BACK_COL)).value = header
            sheet.range((back_row + 1, BACK_COL)).value = rows
            last = back_row + len(rows)
            back_width = max(back_width, len(header))
            source = f"{_a1(back_row, BACK_COL)}:{_a1(last, BACK_COL + len(header) - 1)}"
            back_row = last + 3

            inner = sheet.range((r + 1, c + 0), (r + h - 2, c + w - 2))
            before = len(sheet.charts)
            kind = _render_type(plan, variant - 1)
            res = create_excel_chart(
                data_range=source, chart_type=kind, title=plan["title"],
                position_cell=_a1(r + 1, c), width=int(inner.width) - 12, height=int(inner.height) - 8,
                workbook_name=book.name, sheet_name=sheet.name)
            if not res.get("success"):
                raise ExcelToolError(f"Chart '{plan['title']}' failed: {res.get('error')}")

            if len(sheet.charts) > before:                 # force exact tile geometry
                chart = sheet.charts[-1]
                chart.left, chart.top = inner.left + 6, inner.top + 4
                chart.width, chart.height = inner.width - 12, inner.height - 8
                n_series = len(header) - 1
                vary = n_series == 1 and plan["type"] in {"pie", "bar", "hist"}
                style_report[plan["title"]] = style_chart(
                    chart, theme, n_series, n_points=len(rows), vary=vary,
                    legend=(plan["type"] == "pie" or n_series > 1), offset=i)
            created.append(plan["title"])

        payload = {"source_sheet": source_name, "theme": theme, "layout": layout, "variant": variant,
                   "kpis": [k["label"] for k in kpis], "charts": created, **(meta or {})}
        sheet.range((1, META_COL)).value = META_PREFIX + json.dumps(payload)
        mask_range(sheet, 1, META_COL, 1, META_COL)
        mask_range(sheet, 1, BACK_COL, back_row + 2, BACK_COL + back_width)   # helper data invisible

        warnings += check_sheet(sheet, len(charts))
        hide_gridlines(sheet)
        sheet.activate()
    except Exception as exc:
        book.app.display_alerts = False
        try:
            sheet.delete()
        finally:
            book.app.display_alerts = True
        return {"success": False, "error": f"Dashboard build failed: {type(exc).__name__}: {exc}"}

    return {"success": True, "sheet": sheet.name, "title": title, "theme": theme, "layout": layout,
            "variant": variant, "kpis": [k["label"] for k in kpis], "charts": created,
            "insights": [x["headline"] for x in insights], "chart_style": style_report,
            "validation": summarize(warnings), "notes": notes}


# ----------------------------------------------------------------------------------
# IMAGE engine (DATA = plan_builder, DESIGN = renderer)
# ----------------------------------------------------------------------------------
def _matches(text: str, patterns: list[str] | None) -> bool:
    return any(p.lower() in str(text).lower() for p in (patterns or []) if str(p).strip())


def _log_image_error() -> None:
    import traceback
    from pathlib import Path
    try:
        Path(__file__).with_name("image_error.log").write_text(traceback.format_exc(), encoding="utf-8")
    except Exception:
        pass


def _apply_excludes(plan: dict[str, Any], exclude_charts, exclude_kpis) -> None:
    """Remove charts/KPIs the user dropped, then re-pack the grid."""
    if not (exclude_charts or exclude_kpis):
        return
    from dashboard import plan_builder as pb
    plan["charts"] = [c for c in plan["charts"] if not _matches(c["title"], exclude_charts)]
    plan["kpis"] = [k for k in plan["kpis"] if not _matches(k["label"], exclude_kpis)]
    cols = pb.GRID_COLS
    k = max(len(plan["kpis"]), 1)
    w = cols // k
    for i, kp in enumerate(plan["kpis"]):
        kp["pos"] = {"row": pb.BANNER_H, "col": i * w, "h": pb.KPI_H,
                     "w": w if i < k - 1 else cols - w * (k - 1)}
    start = pb.BANNER_H + (pb.KPI_H if plan["kpis"] else 0)
    end = pb._pack(plan["charts"], start)
    plan["insights_pos"] = ({"row": end, "col": 0, "w": cols, "h": pb.INSIGHT_H}
                            if plan["insights"] else None)


def _build_image_dashboard(data: pd.DataFrame, source, domain: dict[str, Any],
                           design: dict[str, Any], filters_text: str,
                           exclude_charts, exclude_kpis) -> dict[str, Any]:
    """Raises on any failure so the caller can fall back to the native engine."""
    from dashboard.plan_builder import build_plan
    from dashboard.renderer import THEMES, render_dashboard

    book = source.book
    wanted = design.get("theme")
    if wanted not in THEMES:                       # pick_design theme not drawn by renderer yet
        names = list(THEMES)
        wanted = names[(int(design.get("variant") or 1) - 1) % len(names)]

    plan = build_plan(data, title=domain["title"], domain=str(domain["domain"]).lower(),
                      style_request={"theme": wanted, "variant": None, "color": None})
    _apply_excludes(plan, exclude_charts, exclude_kpis)
    if not plan["kpis"] and not plan["charts"]:
        raise ExcelToolError("plan has no KPIs or charts")
    if filters_text:
        plan["subtitle"] = plan.get("subtitle", "") + f" · Filter: {filters_text}"

    sheet = _new_sheet(book)
    try:
        res = render_dashboard(plan, workbook_name=book.name, sheet_name=sheet.name)
        payload = {"source_sheet": source.name, "theme": res.get("theme"),
                   "layout": design.get("layout"), "variant": design.get("variant"),
                   "kpis": [k["label"] for k in plan["kpis"]],
                   "charts": [c["title"] for c in plan["charts"]],
                   "engine": "image",
                   "excluded_charts": list(exclude_charts or []),
                   "excluded_kpis": list(exclude_kpis or [])}
        sheet.range((1, META_COL)).value = META_PREFIX + json.dumps(payload)
        mask_range(sheet, 1, META_COL, 1, META_COL)
        sheet.activate()
    except Exception:
        book.app.display_alerts = False
        try:
            sheet.delete()
        finally:
            book.app.display_alerts = True
        raise

    return {"success": True, "sheet": sheet.name, "title": plan["title"], "engine": "image",
            "theme": res.get("theme"), "kpis": [k["label"] for k in plan["kpis"]],
            "charts": [c["title"] for c in plan["charts"]],
            "chart_why": {c["title"]: c["why"] for c in plan["charts"]},
            "insights": plan["insights"], "dropped_columns": plan["dropped_columns"],
            "validation": plan["validation"], "warnings": plan["warnings"],
            "pictures": res.get("pictures")}


# ----------------------------------------------------------------------------------
# Public entry point
# ----------------------------------------------------------------------------------
def auto_dashboard(workbook_name: str | None = None, sheet_name: str | None = None,
                   filters: list[dict[str, Any]] | None = None, theme: str | None = None,
                   layout: str | None = None, regenerate: bool = False,
                   variant: int | None = None, exclude_charts: list[str] | None = None,
                   exclude_kpis: list[str] | None = None) -> dict[str, Any]:
    """theme/layout optional (forced). regenerate=True -> new theme, layout, KPIs and charts.
    variant/exclude_* are used by refresh/modify to rebuild the same dashboard."""
    try:
        source = _pick_source(workbook_name, sheet_name)
        df, _ = load_dataframe(source)
        df = parse_date_columns(df)
        if not analyse_dataframe(df)["success"]:
            return {"success": False, "error": "Dataset is empty."}
        filtered = apply_filters(df, filters)
        if not filtered["success"]:
            return filtered

        data = filtered["df"]
        analysis = analyse_dataframe(data)
        domain = detect_domain(list(data.columns))
        design = pick_design(data, domain["domain"], bool(analysis.get("date")),
                             theme, layout, regenerate)
        if variant:
            design["variant"] = int(variant)           # keep the original KPI/chart choice

        if USE_IMAGE_DASHBOARD:
            try:
                return _build_image_dashboard(data, source, domain, design,
                                              "; ".join(filtered["applied"]),
                                              exclude_charts, exclude_kpis)
            except Exception:
                _log_image_error()                     # fall through to the native dashboard

        v = design["variant"] - 1                      # 0 = first look; grows with every "another"
        kpis = [k for k in auto_kpis(data, analysis, v, domain["domain"])
                if not _matches(k["label"], exclude_kpis)]
        charts = [c for c in auto_charts(data, analysis, v)
                  if not _matches(c["title"], exclude_charts)]
        if not kpis and not charts:
            return {"success": False, "error": "No valid KPIs or charts found for this dataset."}
        meta = {"excluded_charts": list(exclude_charts or []), "excluded_kpis": list(exclude_kpis or [])}
        return build_dashboard(data, kpis, charts, source.book, source.name,
                               "; ".join(filtered["applied"]), domain["title"], design, meta)
    except Exception as exc:
        return {"success": False, "error": f"{type(exc).__name__}: {exc}"}


if __name__ == "__main__":
    print(auto_dashboard())