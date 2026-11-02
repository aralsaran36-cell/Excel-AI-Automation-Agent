"""Dashboard styler: dark/light themes + Excel styling helpers for dashboard sheets.

Theme name: "purple" (dark, default) or "purple-light". The style_* helpers write
ONLY to the sheet you pass in (never to the source data sheet).
Test (Excel open, makes then deletes temp sheet 'Style_Test'):
    python -m dashboard.dashboard_styler
"""
from __future__ import annotations

import time
from typing import Any

from excel_tools.workbook_manager import get_workbook, hex_to_rgb

_BASE: dict[str, Any] = {
    "white": "#FFFFFF",
    "font_title": 22, "font_subtitle": 11, "font_heading": 13,
    "font_card_label": 10, "font_card_value": 20,
}

ACCENTS: dict[str, dict[str, Any]] = {
    "blue":   {"accent": "#4FA3FF", "dark_bg": "#0F1724", "dark_card": "#1B2638",
               "light_primary": "#1F4E78", "light_card": "#EAF1F8",
               "palette": ["#4FA3FF", "#FF8A3D", "#4CD6A0", "#FFD23F", "#B07CFF", "#FF5D7A", "#7F8FA6", "#2E75B6"]},
    "purple": {"accent": "#B07CFF", "dark_bg": "#150F26", "dark_card": "#241A3D",
               "light_primary": "#4B2E83", "light_card": "#F0EBF8",
               "palette": ["#B07CFF", "#FF5DA2", "#2EE6D6", "#FFB84D", "#7DDC6B", "#5DA9FF", "#FF7A59", "#8E7CC3"]},
    "orange": {"accent": "#FF9A4D", "dark_bg": "#1C130C", "dark_card": "#2E2016",
               "light_primary": "#C0501A", "light_card": "#FDF0E6",
               "palette": ["#FF9A4D", "#4FA3FF", "#6FD66B", "#FFD23F", "#FF5D7A", "#B07CFF", "#7F8FA6", "#C0501A"]},
    "green":  {"accent": "#4CD68A", "dark_bg": "#0C1A13", "dark_card": "#16291F",
               "light_primary": "#1E6B3A", "light_card": "#E8F5EC",
               "palette": ["#4CD68A", "#FFB84D", "#4FA3FF", "#FF7A59", "#FFD23F", "#B07CFF", "#7F8FA6", "#1E6B3A"]},
    "teal":   {"accent": "#2EE6D6", "dark_bg": "#0A1A1C", "dark_card": "#142B2E",
               "light_primary": "#0F6B73", "light_card": "#E6F5F6",
               "palette": ["#2EE6D6", "#FF7A59", "#B8E04A", "#FFD23F", "#4FA3FF", "#FF5DA2", "#7F8FA6", "#0F6B73"]},
    "slate":  {"accent": "#9FB3C8", "dark_bg": "#12171C", "dark_card": "#1F272F",
               "light_primary": "#37474F", "light_card": "#ECEFF1",
               "palette": ["#9FB3C8", "#FF7043", "#66BB6A", "#FFCA28", "#5C6BC0", "#8D6E63", "#26A6B0", "#607D8B"]},
}
THEMES = ACCENTS  # theme names list (kept for tests/imports)


def get_theme(name: str | None = "blue") -> dict[str, Any]:
    """Theme dict. 'purple' = dark, 'purple-light' = light. Unknown name -> blue dark."""
    base, _, mode = (name or "blue").lower().partition("-")
    a = ACCENTS.get(base, ACCENTS["blue"])
    if mode == "light":
        t = {"banner": a["light_primary"], "title_text": "#FFFFFF", "accent": a["light_primary"],
             "heading_text": "#FFFFFF", "page_bg": "#F7F9FB", "card_bg": a["light_card"],
             "card_label": "#5A6B7B", "card_value": a["light_primary"], "text": "#1F2933",
             "chart_bg": "#FFFFFF", "palette": a["palette"], "mode": "light"}
    else:
        t = {"banner": a["dark_card"], "title_text": "#FFFFFF", "accent": a["accent"],
             "heading_text": a["dark_bg"], "page_bg": a["dark_bg"], "card_bg": a["dark_card"],
             "card_label": "#9AA7B8", "card_value": a["accent"], "text": "#E6EAF0",
             "chart_bg": a["dark_card"], "palette": a["palette"], "mode": "dark"}
    return {**_BASE, **t}


def chart_colors(n: int, theme: str | None = "blue") -> list[str]:
    palette = get_theme(theme)["palette"]
    return [palette[i % len(palette)] for i in range(max(n, 0))]


def _block(sheet, row: int, col: int, width: int = 1, height: int = 1):
    return sheet.range((row, col), (row + height - 1, col + width - 1))


def style_page(sheet, rows: int = 60, cols: int = 20, theme: str | None = "blue") -> None:
    t = get_theme(theme)
    area = _block(sheet, 1, 1, cols, rows)
    area.color = hex_to_rgb(t["page_bg"])
    area.font.color = hex_to_rgb(t["text"])


def style_title(sheet, row: int, col: int, width: int, text: str, subtitle: str = "",
                theme: str | None = "blue") -> None:
    t = get_theme(theme)
    banner = _block(sheet, row, col, width, 2)
    banner.color = hex_to_rgb(t["banner"])
    banner.font.color = hex_to_rgb(t["title_text"])
    title = sheet.range((row, col))
    title.value = text
    title.font.bold = True
    title.font.size = t["font_title"]
    sub = sheet.range((row + 1, col))
    sub.value = subtitle
    sub.font.size = t["font_subtitle"]
    sub.font.italic = True
    sub.font.color = hex_to_rgb(t["card_label"] if t["mode"] == "dark" else t["white"])


def style_section_heading(sheet, row: int, col: int, width: int, text: str,
                          theme: str | None = "blue") -> None:
    t = get_theme(theme)
    bar = _block(sheet, row, col, width, 1)
    bar.color = hex_to_rgb(t["accent"])
    bar.font.color = hex_to_rgb(t["heading_text"])
    cell = sheet.range((row, col))
    cell.value = text
    cell.font.bold = True
    cell.font.size = t["font_heading"]


def style_kpi_card(sheet, row: int, col: int, width: int, label: str, value_text: str,
                   theme: str | None = "blue") -> None:
    t = get_theme(theme)
    card = _block(sheet, row, col, width, 2)
    card.color = hex_to_rgb(t["card_bg"])
    lab = sheet.range((row, col))
    lab.value = label
    lab.font.size = t["font_card_label"]
    lab.font.color = hex_to_rgb(t["card_label"])
    val = sheet.range((row + 1, col))
    val.number_format = "@"          # keep as text, else Excel turns "119,231.42" into a number and shows ####
    val.value = value_text
    val.font.bold = True
    val.font.size = t["font_card_value"]
    val.font.color = hex_to_rgb(t["card_value"])
    val.api.horizontal_alignment.set(-4131) if False else None


def set_column_widths(sheet, first_col: int, last_col: int, width: float) -> None:
    for c in range(first_col, last_col + 1):
        sheet.range((1, c)).column_width = width


def hide_gridlines(sheet) -> bool:
    """Hide gridlines of this sheet's window. Returns True if it worked (Mac: unverified)."""
    try:
        sheet.activate()
        window = sheet.book.app.api.active_window
        window.display_gridlines.set(False)
        return True
    except Exception:
        return False


def style_chart(chart, theme: str | None = "blue") -> dict[str, bool]:
    """Try dark background + light text on an xlwings chart. Never raises.
    Returns which parts worked, e.g. {'background': True, 'text': False}."""
    t = get_theme(theme)
    rgb = hex_to_rgb(t["chart_bg"])
    text_rgb = hex_to_rgb(t["text"])
    result = {"background": False, "text": False}
    try:
        obj = chart.api[1] if isinstance(chart.api, tuple) else chart.api
    except Exception:
        return result
    try:
        obj.chart_area.interior.color.set(rgb)
        result["background"] = True
    except Exception:
        pass
    try:
        obj.chart_area.font_object.color.set(text_rgb)
        result["text"] = True
    except Exception:
        pass
    return result


def apply_demo(sheet) -> None:
    """Three dark themes + one light, stacked vertically (8 rows each)."""
    set_column_widths(sheet, 1, 12, 14)
    for i, name in enumerate(["blue", "purple", "green", "orange-light"]):
        top = 1 + i * 8
        style_page(sheet, rows=8, cols=12, theme=name)
        style_title(sheet, top, 1, 12, f"{name} Theme", "Theme preview", name)
        style_section_heading(sheet, top + 2, 1, 12, "Key Metrics", name)
        style_kpi_card(sheet, top + 3, 1, 3, "Total Revenue", "119,231.42", name)
        style_kpi_card(sheet, top + 3, 5, 3, "Total Units", "1,957", name)
        style_kpi_card(sheet, top + 3, 9, 3, "Total Records", "100", name)


if __name__ == "__main__":
    book = get_workbook()
    name = "Style_Test"
    if name in [s.name for s in book.sheets]:
        print(f"STOP: '{name}' sheet already irukku. Adha delete pannitu run pannunga.")
    else:
        sheet = book.sheets.add(name, after=book.sheets[-1])
        try:
            apply_demo(sheet)
            sheet.range("N2").value = [["x", "y"], ["a", 3], ["b", 5], ["c", 4]]
            chart = sheet.charts.add(left=sheet.range("N8").left, top=sheet.range("N8").top, width=300, height=180)
            chart.set_source_data(sheet.range("N2:O5"))
            print("gridlines hidden :", "OK" if hide_gridlines(sheet) else "FAIL")
            print("chart style      :", style_chart(chart, "purple"))
            sheet.activate()
            print(f"Excel-la '{name}' sheet paarunga (15 seconds). Aprom auto-delete aagum...")
            time.sleep(15)
        finally:
            book.app.display_alerts = False
            try:
                sheet.delete()
            finally:
                book.app.display_alerts = True
        print("Sheets now:", [s.name for s in book.sheets])
        print("themes:", list(THEMES))
        print("chart_colors(3, 'purple'):", chart_colors(3, "purple"))