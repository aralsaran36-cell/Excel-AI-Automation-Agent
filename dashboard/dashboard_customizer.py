"""Dashboard customizer: turns a change request into new dashboard settings. No Excel access.

Used by dashboard_manager.modify_dashboard. Supported actions (aliases in brackets):
  theme (colour, color, style), layout, remove_chart, remove_kpi,
  restore (clears all removals, or only the items matching `value`),
  restore_chart (add_chart), restore_kpi (add_kpi).
"""
from __future__ import annotations

from typing import Any

from dashboard.dashboard_layout import LAYOUTS
from dashboard.dashboard_styler import THEMES
from excel_tools.workbook_manager import ExcelToolError

ACTIONS = ("theme", "layout", "remove_chart", "remove_kpi", "restore", "restore_chart", "restore_kpi")
ALIASES = {
    "colour": "theme", "color": "theme", "style": "theme", "colours": "theme", "colors": "theme",
    "delete_chart": "remove_chart", "delete_kpi": "remove_kpi",
    "add_chart": "restore_chart", "add_kpi": "restore_kpi", "reset": "restore",
}
COLOUR_WORDS = {
    "purple": "neon_purple", "violet": "neon_purple", "neon": "neon_purple",
    "teal": "cyber_teal", "cyan": "cyber_teal", "cyber": "cyber_teal",
    "orange": "sunset", "red": "sunset", "sunset": "sunset",
    "blue": "ocean_blue", "ocean": "ocean_blue",
    "green": "emerald", "emerald": "emerald",
    "gold": "midnight_gold", "yellow": "midnight_gold", "black": "midnight_gold", "midnight": "midnight_gold",
    "pink": "rose_noir", "rose": "rose_noir", "noir": "rose_noir",
    "aurora": "aurora", "mint": "aurora",
}


def normalise_action(action: Any) -> str:
    act = str(action or "").strip().lower().replace(" ", "_").replace("-", "_")
    act = ALIASES.get(act, act)
    if act not in ACTIONS:
        raise ExcelToolError(f"action must be one of {list(ACTIONS)} (colour = theme).")
    return act


def resolve_theme(text: str) -> str:
    key = text.strip().lower().replace(" ", "_")
    if key in THEMES:
        return key
    for word in text.lower().replace("_", " ").split():
        if word in THEMES:
            return word
        if word in COLOUR_WORDS:
            return COLOUR_WORDS[word]
    raise ExcelToolError(f"Theme '{text}' not found. Available: {sorted(THEMES)}")


def resolve_layout(text: str) -> str:
    for word in text.lower().replace("_", " ").split():
        if word in LAYOUTS:
            return word
    raise ExcelToolError(f"Layout '{text}' not found. Available: {LAYOUTS}")


def _match(text: str, items: list[str]) -> list[str]:
    needle = text.strip().lower()
    return [i for i in items if needle and needle in str(i).lower()]


def plan_change(meta: dict[str, Any], action: Any, value: str | None) -> dict[str, Any]:
    """Returns theme, layout (None = keep) and the FULL exclusion lists for the rebuild."""
    act = normalise_action(action)
    text = str(value).strip() if value is not None else ""
    if act not in {"restore", "restore_chart", "restore_kpi"} and not text:
        raise ExcelToolError("value is required for this action.")
    ex_c = list(meta.get("excluded_charts", []))
    ex_k = list(meta.get("excluded_kpis", []))
    theme = layout = None
    summary = act

    if act == "theme":
        theme = resolve_theme(text)
        summary = f"theme -> {theme}"
    elif act == "layout":
        layout = resolve_layout(text)
        summary = f"layout -> {layout}"
    elif act == "remove_chart":
        hits = _match(text, list(meta.get("charts", [])))
        if not hits:
            raise ExcelToolError(f"No chart matches '{text}'. Charts: {meta.get('charts', [])}")
        ex_c += [h for h in hits if h not in ex_c]
        summary = f"removed chart: {', '.join(hits)}"
    elif act == "remove_kpi":
        hits = _match(text, list(meta.get("kpis", [])))
        if not hits:
            raise ExcelToolError(f"No KPI matches '{text}'. KPIs: {meta.get('kpis', [])}")
        ex_k += [h for h in hits if h not in ex_k]
        summary = f"removed KPI: {', '.join(hits)}"
    else:                                                   # restore / restore_chart / restore_kpi
        if not text:
            if act in {"restore", "restore_chart"}:
                ex_c = []
            if act in {"restore", "restore_kpi"}:
                ex_k = []
            summary = "restored all removed items"
        else:
            back_c = _match(text, ex_c) if act != "restore_kpi" else []
            back_k = _match(text, ex_k) if act != "restore_chart" else []
            if not back_c and not back_k:
                raise ExcelToolError(
                    f"Nothing removed matches '{text}'. Removed charts: {ex_c}, removed KPIs: {ex_k}")
            ex_c = [c for c in ex_c if c not in back_c]
            ex_k = [k for k in ex_k if k not in back_k]
            summary = f"restored: {', '.join(back_c + back_k)}"
    return {"action": act, "theme": theme, "layout": layout,
            "exclude_charts": ex_c, "exclude_kpis": ex_k, "summary": summary}