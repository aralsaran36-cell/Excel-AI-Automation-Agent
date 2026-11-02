"""Auto Dashboard: analyse -> domain -> KPIs -> charts -> filter -> style -> build a new sheet.

Reads the source sheet, writes ONLY to a NEW 'Dashboard' sheet. Refuses if it already exists.
Guided Dashboard reuses build_dashboard() with its own kpis/charts/title/theme.
Test (Excel open): python -m dashboard.auto_dashboard
"""
from __future__ import annotations

import math
from typing import Any

import pandas as pd

from dashboard.chart_selector import auto_charts
from dashboard.dashboard_filter import apply_filters
from dashboard.dashboard_styler import (
    hide_gridlines, set_column_widths, style_kpi_card, style_page, style_section_heading, style_title,
)
from dashboard.dataset_analyser import analyse_dataframe, parse_date_columns
from dashboard.domain_detector import detect_domain
from dashboard.kpi_selector import auto_kpis, format_value
from excel_tools.charts import create_excel_chart
from excel_tools.workbook_manager import ExcelToolError, get_sheet, load_dataframe

DASHBOARD_SHEET = "Dashboard"
GRID_COLS = 12            # dashboard area width (columns A:L)
COL_WIDTH = 14
TABLE_COL = 14            # column N: backing tables for charts
CHART_H = 270
CHART_ROWS = 20           # rows reserved per chart row (title row + chart)
SCATTER_MAX_POINTS = 500


def _a1(row: int, col: int) -> str:
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def _py(value: Any) -> Any:
    """numpy/pandas value -> plain Python value Excel can take."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    if hasattr(value, "item"):
        value = value.item()
    return round(value, 2) if isinstance(value, float) else value


def _table(df: pd.DataFrame, plan: dict[str, Any]) -> tuple[list[str], list[list[Any]]]:
    """Backing table (header + rows) for one chart plan."""
    kind, cat, meas, agg = plan["type"], plan["category"], plan["measure"], plan["agg"]

    if kind == "scatter":
        sub = df[[meas, plan["measure2"]]].dropna()
        if len(sub) > SCATTER_MAX_POINTS:
            sub = sub.sample(SCATTER_MAX_POINTS, random_state=1)
        return [meas, plan["measure2"]], [[_py(a), _py(b)] for a, b in sub.values.tolist()]

    if kind == "stacked":
        pv = df.pivot_table(index=cat, columns=plan["series"], values=meas, aggfunc=agg, fill_value=0)
        pv = pv.loc[pv.sum(axis=1).sort_values(ascending=False).index]
        header = [cat] + [str(c) for c in pv.columns]
        return header, [[str(idx)] + [_py(v) for v in row] for idx, row in zip(pv.index, pv.values.tolist())]

    if kind == "line":
        key = df[cat]
        if pd.api.types.is_datetime64_any_dtype(key):
            many = key.nunique() > 24
            key = key.dt.to_period("M").dt.to_timestamp() if many else key
            grouped = df.groupby(key)[meas].agg(agg).sort_index()
            fmt = "%b %Y" if many else "%Y-%m-%d"
            return [cat, meas], [[k.strftime(fmt), _py(v)] for k, v in grouped.items()]
        grouped = df.groupby(key)[meas].agg(agg).sort_index()
        return [cat, meas], [[str(k), _py(v)] for k, v in grouped.items()]

    grouped = df.groupby(cat)[meas].agg(agg).sort_values(ascending=False)
    if plan.get("top_n"):
        grouped = grouped.head(plan["top_n"])
    return [cat, meas], [[str(k), _py(v)] for k, v in grouped.items()]


def _kpi_layout(count: int) -> tuple[int, int]:
    """(cards per row, card width in columns) so cards fill the 12-column grid."""
    per_row = max(1, min(count, 4)) if count <= 4 else 3
    return per_row, GRID_COLS // per_row - 1


def build_dashboard(df: pd.DataFrame, kpis: list[dict[str, Any]], charts: list[dict[str, Any]],
                    book, source_name: str, filters_text: str = "",
                    title: str = "Data Dashboard", theme: str = "blue") -> dict[str, Any]:
    """Build the dashboard sheet from ready plans. Returns a result dict."""
    if DASHBOARD_SHEET in [s.name for s in book.sheets]:
        return {"success": False, "error": f"'{DASHBOARD_SHEET}' sheet already irukku. Adha rename/delete pannitu thirumba try pannunga."}

    per_row, card_w = _kpi_layout(len(kpis))
    kpi_rows = ((len(kpis) - 1) // per_row + 1) * 3 if kpis else 0
    chart_head = 5 + kpi_rows + 1
    chart_lines = (len(charts) + 1) // 2
    total_rows = chart_head + 2 + chart_lines * CHART_ROWS + 2

    sheet = book.sheets.add(DASHBOARD_SHEET, after=book.sheets[-1])
    created: list[str] = []
    notes: list[str] = []
    try:
        set_column_widths(sheet, 1, GRID_COLS, COL_WIDTH)
        style_page(sheet, rows=total_rows, cols=GRID_COLS, theme=theme)
        subtitle = f"Source: {source_name} | {len(df)} rows" + (f" | Filters: {filters_text}" if filters_text else "")
        style_title(sheet, 1, 1, GRID_COLS, title, subtitle, theme)
        sheet.range((1, 1)).row_height = 36
        sheet.range((2, 1)).row_height = 20

        style_section_heading(sheet, 4, 1, GRID_COLS, "Key Metrics", theme)
        for i, kpi in enumerate(kpis):
            row = 5 + (i // per_row) * 3
            col = 1 + (i % per_row) * (card_w + 1)
            style_kpi_card(sheet, row, col, card_w, kpi["label"], format_value(kpi), theme)
            sheet.range((row + 1, 1)).row_height = 34

        style_section_heading(sheet, chart_head, 1, GRID_COLS, "Charts", theme)

        table_row = 2
        sheet.range((1, TABLE_COL)).value = "Chart data (backing tables)"
        for idx, plan in enumerate(charts):
            header, rows = _table(df, plan)
            if not rows:
                notes.append(f"'{plan['title']}' skip: data illa")
                continue
            sheet.range((table_row, TABLE_COL)).value = header
            sheet.range((table_row + 1, TABLE_COL)).value = rows
            last = table_row + len(rows)
            source = f"{_a1(table_row, TABLE_COL)}:{_a1(last, TABLE_COL + len(header) - 1)}"

            chart_row = chart_head + 2 + (idx // 2) * CHART_ROWS
            chart_col = 1 if idx % 2 == 0 else 7
            half = sheet.range((chart_row, chart_col), (chart_row, chart_col + 5))
            sheet.range((chart_row, chart_col)).value = plan["title"]
            sheet.range((chart_row, chart_col)).font.bold = True

            kind = plan["type"]
            tool_type = {"bar": "column", "stacked": "column"}.get(kind, kind)
            res = create_excel_chart(
                data_range=source, chart_type=tool_type, title=plan["title"],
                position_cell=_a1(chart_row + 1, chart_col),
                width=int(half.width) - 10, height=CHART_H,
                workbook_name=book.name, sheet_name=DASHBOARD_SHEET)
            if not res.get("success"):
                raise ExcelToolError(f"Chart '{plan['title']}' fail: {res.get('error')}")
            if kind == "stacked":
                try:
                    sheet.charts[res["name"]].chart_type = "column_stacked"
                except Exception as exc:
                    notes.append(f"stacked support aagala, clustered column vachirukken ({type(exc).__name__})")
            created.append(plan["title"])
            table_row = last + 3

        if not hide_gridlines(sheet):
            notes.append("gridlines hide aagala (Mac limitation)")
        sheet.activate()
    except Exception as exc:
        book.app.display_alerts = False
        try:
            sheet.delete()
        finally:
            book.app.display_alerts = True
        return {"success": False, "error": f"Dashboard build fail: {type(exc).__name__}: {exc}. Half-built sheet remove aachu."}

    return {"success": True, "sheet": DASHBOARD_SHEET, "title": title, "theme": theme,
            "kpis": [k["label"] for k in kpis], "charts": created, "notes": notes}


def auto_dashboard(workbook_name: str | None = None, sheet_name: str | None = None,
                   filters: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    """Full Auto Dashboard flow on the active workbook/sheet."""
    try:
        source = get_sheet(workbook_name, sheet_name)
        df, _ = load_dataframe(source)
        df = parse_date_columns(df)   # text dates -> real dates (in memory only, Excel data unchanged)
        if not analyse_dataframe(df)["success"]:
            return {"success": False, "error": "Dataset empty-a irukku."}
        filtered = apply_filters(df, filters)
        if not filtered["success"]:
            return filtered
        data = filtered["df"]
        analysis = analyse_dataframe(data)
        domain = detect_domain(list(data.columns))
        kpis, charts = auto_kpis(data, analysis), auto_charts(data, analysis)
        if not kpis and not charts:
            return {"success": False, "error": "Indha dataset-ku dashboard-ku useful KPI/chart kidaikkala."}
        return build_dashboard(data, kpis, charts, source.book, source.name,
                               "; ".join(filtered["applied"]),
                               title=domain["title"], theme=domain["theme"])
    except Exception as exc:
        return {"success": False, "error": f"{type(exc).__name__}: {exc}"}


if __name__ == "__main__":
    print(auto_dashboard())