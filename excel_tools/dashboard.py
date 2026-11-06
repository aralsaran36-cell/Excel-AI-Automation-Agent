"""Dashboard tools: create, list, refresh, modify and delete dashboards (thin wrappers).

All logic lives in the dashboard/ package. Dashboards are always separate sheets named
'Dashboard', 'Dashboard 2', ... and the source data sheet is never changed.
All tools return {"success": bool, ...} and never raise (via @safe_tool).
"""
from excel_tools.workbook_manager import safe_tool


@safe_tool
def auto_dashboard(
    workbook_name: str | None = None,
    sheet_name: str | None = None,
    theme: str | None = None,
    layout: str | None = None,
    regenerate: bool = False,
) -> dict:
    """
    PURPOSE:
        Build a premium dashboard (KPIs, charts, insights) in a NEW sheet. The look is chosen from the data. regenerate=true makes a different dashboard.

    PARAMETERS:
        workbook_name: Workbook name (optional).
        sheet_name: Source data sheet (optional).
        theme: neon_purple, cyber_teal, sunset, ocean_blue, emerald, midnight_gold, rose_noir, aurora.
        layout: command, hero, sidebar, magazine, split or tiles (optional).
        regenerate: true for another dashboard with a new look.

    RETURNS:
        success, sheet, theme, layout, variant, kpis, charts, insights.
    """
    from dashboard.auto_dashboard import auto_dashboard as _auto

    return _auto(
        workbook_name=workbook_name,
        sheet_name=sheet_name,
        theme=theme,
        layout=layout,
        regenerate=bool(regenerate),
    )


@safe_tool
def list_dashboards(workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        List the dashboard sheets in the workbook with their theme, layout and source sheet.

    PARAMETERS:
        workbook_name: Workbook name (optional).

    RETURNS:
        success, count, dashboards (sheet, theme, layout, variant, source_sheet, charts, kpis).
    """
    from dashboard.dashboard_manager import list_dashboards as _list

    return _list(workbook_name)


@safe_tool
def refresh_dashboard(
    dashboard_name: str | None = None,
    workbook_name: str | None = None,
) -> dict:
    """
    PURPOSE:
        Rebuild a dashboard from the latest source data with the same look, KPIs and charts.

    PARAMETERS:
        dashboard_name: Dashboard sheet name or number (optional, default latest).
        workbook_name: Workbook name (optional).

    RETURNS:
        success, sheet, theme, layout, kpis, charts, backup.
    """
    from dashboard.dashboard_manager import refresh_dashboard as _refresh

    return _refresh(workbook_name, dashboard_name)


@safe_tool
def modify_dashboard(
    action: str,
    value: str | None = None,
    dashboard_name: str | None = None,
    workbook_name: str | None = None,
) -> dict:
    """
    PURPOSE:
        Change an existing dashboard: theme, layout, remove one chart or KPI, or restore removed ones.

    PARAMETERS:
        action: theme, layout, remove_chart, remove_kpi or restore.
        value: Theme or layout name, or part of the chart title / KPI label to remove.
        dashboard_name: Dashboard sheet name or number (optional, default latest).
        workbook_name: Workbook name (optional).

    RETURNS:
        success, sheet, action, theme, layout, kpis, charts, backup.
    """
    from dashboard.dashboard_manager import modify_dashboard as _modify

    return _modify(workbook_name, dashboard_name, action, value)


@safe_tool
def delete_dashboard(
    dashboard_name: str,
    confirm: bool = False,
    workbook_name: str | None = None,
) -> dict:
    """
    PURPOSE:
        Delete a dashboard sheet. Without confirm=True it only previews. Data sheets are never touched.

    PARAMETERS:
        dashboard_name: Dashboard sheet name or number, for example 'Dashboard 2'.
        confirm: False = preview. True = delete. Set True only after the user confirmed.
        workbook_name: Workbook name (optional).

    RETURNS:
        Preview (needs_confirm) or deleted sheet name and backup path.
    """
    from dashboard.dashboard_manager import delete_dashboard as _delete

    return _delete(workbook_name, dashboard_name, bool(confirm))