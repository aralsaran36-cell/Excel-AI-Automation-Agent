"""Dashboard styler: 8 premium dark themes + tile/card helpers.

Pure formatting. No decisions about WHICH theme/layout: that is dashboard_designer.py.
"""
from __future__ import annotations

from typing import Any

from excel_tools.workbook_manager import hex_to_rgb

_BASE: dict[str, Any] = {
    "font_title": 22, "font_subtitle": 10, "font_tile_title": 11,
    "font_card_label": 9, "font_card_value": 22, "font_body": 10,
}

THEMES: dict[str, dict[str, Any]] = {
    "neon_purple": {
        "page_bg": "#120A21", "card_bg": "#1E1333", "side_bg": "#0D0718",
        "accent": "#B07CFF", "accent2": "#00F0FF", "text": "#F1F5F9", "muted": "#9A8FB5",
        "palette": ["#B07CFF", "#00F0FF", "#FF5DA2", "#2EE6D6", "#FFB84D", "#7DDC6B", "#5DA9FF"]},
    "cyber_teal": {
        "page_bg": "#071518", "card_bg": "#0F2A2F", "side_bg": "#050F11",
        "accent": "#2EE6D6", "accent2": "#B8E04A", "text": "#EAF7F7", "muted": "#7FA8AC",
        "palette": ["#2EE6D6", "#FF7A59", "#B8E04A", "#FFD23F", "#4FA3FF", "#FF5DA2", "#7F8FA6"]},
    "sunset": {
        "page_bg": "#1A0F12", "card_bg": "#2B1819", "side_bg": "#12090B",
        "accent": "#FF7A59", "accent2": "#FFD23F", "text": "#FFF1EA", "muted": "#B39088",
        "palette": ["#FF7A59", "#FFD23F", "#FF5DA2", "#B07CFF", "#2EE6D6", "#4FA3FF", "#7DDC6B"]},
    "ocean_blue": {
        "page_bg": "#08111F", "card_bg": "#12223A", "side_bg": "#050B15",
        "accent": "#4FA3FF", "accent2": "#2EE6D6", "text": "#E8F0FB", "muted": "#8A9DB8",
        "palette": ["#4FA3FF", "#2EE6D6", "#FFB84D", "#B07CFF", "#FF5D7A", "#7DDC6B", "#FFD23F"]},
    "emerald": {
        "page_bg": "#08170F", "card_bg": "#112A1D", "side_bg": "#050F09",
        "accent": "#3DDC84", "accent2": "#F5D547", "text": "#EAF7EF", "muted": "#7FA68D",
        "palette": ["#3DDC84", "#F5D547", "#4FA3FF", "#FF7A59", "#B07CFF", "#2EE6D6", "#FF5DA2"]},
    "midnight_gold": {
        "page_bg": "#0C0C10", "card_bg": "#1A1A22", "side_bg": "#07070A",
        "accent": "#F5C451", "accent2": "#FF8A3D", "text": "#F5F3EE", "muted": "#98958A",
        "palette": ["#F5C451", "#FF8A3D", "#4FA3FF", "#B07CFF", "#2EE6D6", "#FF5D7A", "#7DDC6B"]},
    "rose_noir": {
        "page_bg": "#16090F", "card_bg": "#2A1220", "side_bg": "#0F060A",
        "accent": "#FF5DA2", "accent2": "#FFB3D1", "text": "#FFEFF6", "muted": "#B58AA0",
        "palette": ["#FF5DA2", "#B07CFF", "#FFB84D", "#2EE6D6", "#4FA3FF", "#7DDC6B", "#FF7A59"]},
    "aurora": {
        "page_bg": "#0A0F1F", "card_bg": "#171F3A", "side_bg": "#060913",
        "accent": "#7DF9C6", "accent2": "#B07CFF", "text": "#EEF2FF", "muted": "#8E98C0",
        "palette": ["#7DF9C6", "#B07CFF", "#5DA9FF", "#FF5DA2", "#FFD23F", "#2EE6D6", "#FF8A3D"]},
}
ACCENTS = THEMES


def get_theme(name: str | None = "neon_purple") -> dict[str, Any]:
    return {**_BASE, **THEMES.get((name or "").lower(), THEMES["neon_purple"])}


def chart_colors(n: int, theme: str | None = None) -> list[str]:
    palette = get_theme(theme)["palette"]
    return [palette[i % len(palette)] for i in range(max(n, 0))]


def _block(sheet, row: int, col: int, width: int = 1, height: int = 1):
    return sheet.range((row, col), (row + height - 1, col + width - 1))


def _rgb(hex_color: str):
    return hex_to_rgb(hex_color)


def _com(hex_color: str) -> int:
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


def set_column_widths(sheet, first_col: int, last_col: int, width: float) -> None:
    for c in range(first_col, last_col + 1):
        sheet.range((1, c)).column_width = width


def style_page(sheet, rows: int, cols: int, theme: str | None = None) -> None:
    t = get_theme(theme)
    area = _block(sheet, 1, 1, cols, rows)
    area.color = _rgb(t["page_bg"])
    area.font.color = _rgb(t["text"])


def style_banner(sheet, cols: int, title: str, subtitle: str, theme: str | None = None) -> None:
    t = get_theme(theme)
    _block(sheet, 1, 1, cols, 2).color = _rgb(t["side_bg"])
    cell = sheet.range((1, 2))
    cell.value = title
    cell.font.bold = True
    cell.font.size = t["font_title"]
    cell.font.color = _rgb(t["accent"])
    sub = sheet.range((2, 2))
    sub.value = subtitle
    sub.font.size = t["font_subtitle"]
    sub.font.italic = True
    sub.font.color = _rgb(t["muted"])


def style_sidebar_bg(sheet, row: int, col: int, width: int, height: int, theme: str | None = None) -> None:
    _block(sheet, row, col, width, height).color = _rgb(get_theme(theme)["side_bg"])


def style_tile(sheet, row: int, col: int, width: int, height: int, title: str = "",
               theme: str | None = None) -> None:
    """Card background. The last row/col of the slot stay empty = visible gutter."""
    t = get_theme(theme)
    _block(sheet, row, col, max(width - 1, 1), max(height - 1, 1)).color = _rgb(t["card_bg"])
    if title:
        cell = sheet.range((row, col + 1))
        cell.value = title
        cell.font.bold = True
        cell.font.size = t["font_tile_title"]
        cell.font.color = _rgb(t["text"])


def style_kpi_card(sheet, row: int, col: int, width: int, height: int, label: str,
                   value_text: str, theme: str | None = None, idx: int = 0,
                   pct: float | None = None) -> None:
    t = get_theme(theme)
    color = t["palette"][idx % len(t["palette"])]
    style_tile(sheet, row, col, width, height, theme=theme)
    _block(sheet, row, col, 1, max(height - 1, 1)).color = _rgb(color)      # accent stripe
    lab = sheet.range((row, col + 1))
    lab.value = label.upper()
    lab.font.size = t["font_card_label"]
    lab.font.color = _rgb(t["muted"])
    val = sheet.range((row + 1, col + 1))
    val.number_format = "@"
    val.value = value_text
    val.font.bold = True
    val.font.size = t["font_card_value"]
    val.font.color = _rgb(color)
    if pct is not None and height >= 4:                                      # progress bar for rates
        filled = max(0, min(12, round(float(pct) * 12)))
        bar = sheet.range((row + 2, col + 1))
        bar.value = "█" * filled + "░" * (12 - filled)
        bar.font.size = 9
        bar.font.color = _rgb(color)


def style_insight_card(sheet, row: int, col: int, width: int, height: int,
                       insight: dict[str, Any], theme: str | None = None, idx: int = 0) -> None:
    t = get_theme(theme)
    style_tile(sheet, row, col, width, height, theme=theme)
    head = sheet.range((row, col + 1))
    head.value = insight["title"].upper()
    head.font.size = t["font_card_label"]
    head.font.bold = True
    head.font.color = _rgb(t["accent2"])
    big = sheet.range((row + 1, col + 1))
    big.value = insight["headline"]
    big.font.bold = True
    big.font.size = 13
    big.font.color = _rgb(t["text"])
    det = sheet.range((row + 2, col + 1))
    det.value = insight["detail"]
    det.font.size = t["font_card_label"]
    det.font.color = _rgb(t["muted"])
    if insight.get("pct") is not None:
        filled = max(0, min(14, round(insight["pct"] * 14)))
        bar = sheet.range((row + 3, col + 1))
        bar.value = "█" * filled + "░" * (14 - filled)
        bar.font.color = _rgb(t["palette"][idx % len(t["palette"])])


def mask_range(sheet, row1: int, col1: int, row2: int, col2: int, hex_color: str = "#FFFFFF") -> None:
    """Make helper data invisible (font = sheet background). Cells are NOT hidden,
    because Excel charts do not plot hidden cells."""
    try:
        sheet.range((row1, col1), (row2, col2)).font.color = _rgb(hex_color)
    except Exception:
        pass


def hide_gridlines(sheet) -> bool:
    try:
        sheet.activate()
        sheet.book.app.api.active_window.display_gridlines.set(False)
        return True
    except Exception:
        return False


# ---------------- chart styling (Mac appscript first, Windows COM second) ----------------
def _chart_objs(chart) -> list[Any]:
    """Candidate API objects. Mac xlwings may give a ChartObject or the Chart itself."""
    try:
        api = chart.api
    except Exception:
        return []
    base = list(reversed(api)) if isinstance(api, tuple) else [api]
    out: list[Any] = []
    for obj in base:
        out.append(obj)
        for attr in ("chart", "Chart"):
            try:
                out.append(getattr(obj, attr))
            except Exception:
                pass
    return out


def _try(cache: dict, key: str, objs: list[Any], makers: list) -> Any:
    """Run the first (object, setter) pair that works. Remember it. Returns True or an error text."""
    order = [(o, m) for o in range(len(objs)) for m in range(len(makers))]
    if key in cache and cache[key] in order:
        order.remove(cache[key])
        order.insert(0, cache[key])
    last = "no chart api"
    for o, m in order:
        try:
            makers[m](objs[o])
            cache[key] = (o, m)
            return True
        except Exception as exc:
            last = f"{type(exc).__name__}: {exc}"[:90]
    return last


def style_chart(chart, theme: str | None = None, n_series: int = 1, n_points: int = 0,
                vary: bool = False, legend: bool | None = None, offset: int = 0) -> dict[str, Any]:
    """Dark chart styling. Each value in the result is True or the error text of the last attempt.
    vary=True gives every slice/bar its own colour. offset shifts the palette start per chart."""
    t = get_theme(theme)
    card, text = _rgb(t["card_bg"]), _rgb(t["text"])
    card_i, text_i = _com(t["card_bg"]), _com(t["text"])
    show_legend = (n_series > 1) if legend is None else legend
    objs = _chart_objs(chart)
    if not objs:
        return {"styled": "no chart api"}
    cache: dict = {}
    res: dict[str, Any] = {
        "background": _try(cache, "bg", objs, [
            lambda o: o.chart_area.interior.color.set(card),
            lambda o: o.chart_area.interior.color.set(card_i),
            lambda o: o.chart_area.format.fill.fore_color.set(card),
            lambda o: setattr(o.ChartArea.Format.Fill.ForeColor, "RGB", card_i),
            lambda o: setattr(o.ChartArea.Interior, "Color", card_i)]),
        "plot_bg": _try(cache, "plot", objs, [
            lambda o: o.plot_area.interior.color.set(card),
            lambda o: o.plot_area.interior.color.set(card_i),
            lambda o: setattr(o.PlotArea.Format.Fill.ForeColor, "RGB", card_i),
            lambda o: setattr(o.PlotArea.Interior, "Color", card_i)]),
        "text": _try(cache, "text", objs, [
            lambda o: o.chart_area.font_object.color.set(text),
            lambda o: o.chart_area.font_object.color.set(text_i),
            lambda o: setattr(o.ChartArea.Format.TextFrame2.TextRange.Font.Fill.ForeColor, "RGB", text_i),
            lambda o: setattr(o.ChartArea.Font, "Color", text_i)]),
        "no_title": _try(cache, "title", objs, [
            lambda o: o.has_title.set(False),
            lambda o: setattr(o, "HasTitle", False)]),
        "legend": _try(cache, "legend", objs, [
            lambda o: o.has_legend.set(show_legend),
            lambda o: setattr(o, "HasLegend", show_legend)]),
    }
    palette = t["palette"]
    if vary and n_points:
        bad = None
        for j in range(n_points):
            hex_color = palette[(j + offset) % len(palette)]
            color, ci = _rgb(hex_color), _com(hex_color)
            out = _try(cache, "point", objs, [
                lambda o, j=j, c=color: o.series[1].points[j + 1].interior.color.set(c),
                lambda o, j=j, c=ci: o.series[1].points[j + 1].interior.color.set(c),
                lambda o, j=j, c=color: o.series[1].points[j + 1].format.fill.fore_color.set(c),
                lambda o, j=j, c=ci: setattr(o.SeriesCollection(1).Points(j + 1).Format.Fill.ForeColor, "RGB", c),
                lambda o, j=j, c=ci: setattr(o.SeriesCollection(1).Points(j + 1).Interior, "Color", c)])
            if out is not True:
                bad = out
        res["per_point_colors"] = True if bad is None else bad
    else:
        for i in range(n_series):
            hex_color = palette[(i + offset) % len(palette)]
            color, ci = _rgb(hex_color), _com(hex_color)
            k = i + 1
            res[f"series{k}"] = _try(cache, "series", objs, [
                lambda o, k=k, c=color: o.series[k].interior.color.set(c),
                lambda o, k=k, c=ci: o.series[k].interior.color.set(c),
                lambda o, k=k, c=color: o.series[k].format.fill.fore_color.set(c),
                lambda o, k=k, c=ci: setattr(o.SeriesCollection(k).Format.Fill.ForeColor, "RGB", c),
                lambda o, k=k, c=ci: setattr(o.SeriesCollection(k).Interior, "Color", c)])
            res[f"series{k}_line"] = _try(cache, "line", objs, [      # line charts use the border colour
                lambda o, k=k, c=color: o.series[k].border.color.set(c),
                lambda o, k=k, c=ci: o.series[k].border.color.set(c),
                lambda o, k=k, c=ci: setattr(o.SeriesCollection(k).Format.Line.ForeColor, "RGB", c),
                lambda o, k=k, c=ci: setattr(o.SeriesCollection(k).Border, "Color", c)])
    return res