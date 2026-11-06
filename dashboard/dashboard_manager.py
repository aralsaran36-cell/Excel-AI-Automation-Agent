"""Dashboard manager: list, refresh, modify and delete EXISTING dashboards.

Settings are read from the masked meta cell that auto_dashboard writes on every dashboard sheet.
A rebuild first creates the new sheet; the old one is deleted only after that succeeded.
What a modify request means is decided in dashboard_customizer.py.
"""
from __future__ import annotations

import json
import re
from typing import Any

from dashboard.auto_dashboard import auto_dashboard
from dashboard.dashboard_config import META_COL, META_PREFIX
from dashboard.dashboard_customizer import plan_change
from excel_tools.workbook_manager import ExcelToolError, backup_workbook, get_sheet

NAME_RE = re.compile(r"^dashboard(?: \d+)?$", re.IGNORECASE)
NO_SETTINGS = ("This dashboard has no saved settings (made before this update). "
               "Create a new one with auto_dashboard.")


def _book(workbook_name: str | None):
    return get_sheet(workbook_name, None).book


def is_dashboard(name: str) -> bool:
    return bool(NAME_RE.match(str(name).strip()))


def _number(name: str) -> int:
    digits = re.findall(r"\d+", str(name))
    return int(digits[0]) if digits else 1


def _dash_sheets(book) -> list[Any]:
    return [s for s in book.sheets if is_dashboard(s.name)]


def _find(book, name: str | None):
    sheets = _dash_sheets(book)
    if not sheets:
        raise ExcelToolError("No dashboard sheet found. Create one with auto_dashboard first.")
    text = str(name).strip() if name is not None else ""
    if not text:
        try:
            active = book.sheets.active
            if is_dashboard(active.name):
                return active
        except Exception:
            pass
        return max(sheets, key=lambda s: _number(s.name))
    if text.isdigit():
        text = "Dashboard" if int(text) == 1 else f"Dashboard {int(text)}"
    for s in sheets:
        if s.name.lower() == text.lower():
            return s
    raise ExcelToolError(f"Dashboard '{name}' not found. Available: {[s.name for s in sheets]}")


def read_meta(sheet) -> dict[str, Any] | None:
    try:
        raw = sheet.range((1, META_COL)).value
    except Exception:
        return None
    if not isinstance(raw, str) or not raw.startswith(META_PREFIX):
        return None
    try:
        return json.loads(raw[len(META_PREFIX):])
    except ValueError:
        return None


def list_dashboards(workbook_name: str | None = None) -> dict[str, Any]:
    book = _book(workbook_name)
    items = []
    for s in sorted(_dash_sheets(book), key=lambda x: _number(x.name)):
        meta = read_meta(s)
        items.append({
            "sheet": s.name,
            "theme": meta.get("theme") if meta else None,
            "layout": meta.get("layout") if meta else None,
            "variant": meta.get("variant") if meta else None,
            "source_sheet": meta.get("source_sheet") if meta else None,
            "charts": len(meta.get("charts", [])) if meta else None,
            "kpis": len(meta.get("kpis", [])) if meta else None,
            "has_settings": bool(meta)})
    return {"success": True, "workbook": book.name, "count": len(items), "dashboards": items,
            "message": f"{len(items)} dashboard(s) found." if items else "No dashboards in this workbook."}


def _rebuild(book, sheet, meta: dict[str, Any], theme: str | None = None, layout: str | None = None,
             exclude_charts: list[str] | None = None, exclude_kpis: list[str] | None = None) -> dict[str, Any]:
    old_name = sheet.name
    backup = backup_workbook(book)
    result = auto_dashboard(
        workbook_name=book.name, sheet_name=meta["source_sheet"],
        theme=theme or meta["theme"], layout=layout or meta["layout"],
        variant=meta.get("variant"),
        exclude_charts=exclude_charts, exclude_kpis=exclude_kpis)
    if not result.get("success"):
        return result                                   # old dashboard is untouched
    new_name = result["sheet"]
    book.app.display_alerts = False
    try:
        sheet.delete()
    finally:
        book.app.display_alerts = True
    book.sheets[new_name].name = old_name
    result["sheet"] = old_name
    result["backup"] = backup
    return result


def refresh_dashboard(workbook_name: str | None = None, dashboard_name: str | None = None) -> dict[str, Any]:
    book = _book(workbook_name)
    sheet = _find(book, dashboard_name)
    meta = read_meta(sheet)
    if not meta:
        raise ExcelToolError(NO_SETTINGS)
    result = _rebuild(book, sheet, meta,
                      exclude_charts=meta.get("excluded_charts"), exclude_kpis=meta.get("excluded_kpis"))
    if result.get("success"):
        result["message"] = f"{result['sheet']} refreshed from '{meta['source_sheet']}' with the same look."
    return result


def modify_dashboard(workbook_name: str | None = None, dashboard_name: str | None = None,
                     action: str = "", value: str | None = None) -> dict[str, Any]:
    book = _book(workbook_name)
    sheet = _find(book, dashboard_name)
    meta = read_meta(sheet)
    if not meta:
        raise ExcelToolError(NO_SETTINGS)
    change = plan_change(meta, action, value)           # validates action/value, no Excel access
    result = _rebuild(book, sheet, meta, change["theme"], change["layout"],
                      change["exclude_charts"], change["exclude_kpis"])
    if result.get("success"):
        result["action"] = change["action"]
        result["message"] = f"{result['sheet']} updated: {change['summary']}."
    return result


def delete_dashboard(workbook_name: str | None = None, dashboard_name: str | None = None,
                     confirm: bool = False) -> dict[str, Any]:
    if not dashboard_name or not str(dashboard_name).strip():
        raise ExcelToolError("dashboard_name is required to delete a dashboard.")
    book = _book(workbook_name)
    sheet = _find(book, dashboard_name)
    name = sheet.name
    if not confirm:
        return {"success": True, "deleted": False, "needs_confirm": True, "sheet": name,
                "message": f"'{name}' will be deleted. Call again with confirm=True to delete it."}
    backup = backup_workbook(book)
    book.app.display_alerts = False
    try:
        sheet.delete()
    finally:
        book.app.display_alerts = True
    return {"success": True, "deleted": True, "sheet": name, "backup": backup,
            "message": f"Deleted dashboard '{name}'."}