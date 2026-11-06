"""DESIGN: plan JSON -> matplotlib PNGs -> placed on an Excel sheet.
Never decides data. Only draws what the plan says.

Charts are drawn by the dashboard.Charts package (one file per chart type).
If a chart file fails, the traceback is saved to dashboard/chart_error.log and the
legacy drawing (_draw_chart_legacy) is used, so the dashboard never breaks.
"""
import math
import os
import platform
import tempfile
import traceback
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
from matplotlib.ticker import FuncFormatter

from dashboard.colors import resolve_color

try:
    import xlwings as xw
except ImportError:  # image export still works
    xw = None

CELL_W, ROW_H, GAP, LEFT0, TOP0, DPI = 90, 56, 10, 12, 12, 160
CHART_LOG = Path(__file__).with_name("chart_error.log")

# TEMPORARY themes. Move to colors.py later.
THEMES = {
    "ocean_blue": {"bg": "#0B1220", "card": "#121C2F", "text": "#E8EEF9", "muted": "#8A9BB8",
                   "up": "#22C55E", "down": "#EF4444",
                   "palette": ["#38BDF8", "#3B82F6", "#22D3EE", "#6366F1", "#14B8A6", "#A78BFA", "#60A5FA", "#2DD4BF"]},
    "neon_purple": {"bg": "#0F0A1F", "card": "#1A1233", "text": "#F3EEFF", "muted": "#A79BC8",
                    "up": "#34D399", "down": "#F87171",
                    "palette": ["#A855F7", "#EC4899", "#8B5CF6", "#22D3EE", "#F472B6", "#6366F1", "#C084FC", "#38BDF8"]},
    "sunset": {"bg": "#1A0F0F", "card": "#2A1818", "text": "#FFF1E8", "muted": "#C9A593",
               "up": "#4ADE80", "down": "#F87171",
               "palette": ["#F97316", "#EF4444", "#FACC15", "#EC4899", "#FB923C", "#F43F5E", "#FDE047", "#FDBA74"]},
    "emerald": {"bg": "#07140F", "card": "#0E2219", "text": "#E7FBF0", "muted": "#8DB5A0",
                "up": "#4ADE80", "down": "#F87171",
                "palette": ["#10B981", "#14B8A6", "#84CC16", "#22C55E", "#2DD4BF", "#A3E635", "#34D399", "#5EEAD4"]},
}


def fmt_compact(v):
    if v is None:
        return "-"
    a = abs(v)
    for lim, suf in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if a >= lim:
            return f"{v / lim:.1f}".rstrip("0").rstrip(".") + suf
    return f"{v:,.0f}" if a >= 100 else f"{v:,.2f}".rstrip("0").rstrip(".")


def _theme(plan, theme):
    sr = plan.get("style_request") or {}
    name = theme or sr.get("theme") or plan.get("theme_hint") or "ocean_blue"
    T = dict(THEMES.get(name, THEMES["ocean_blue"]))
    T["name"] = name if name in THEMES else "ocean_blue"
    custom = resolve_color(sr.get("color"))
    if custom:
        T["palette"] = [custom] + T["palette"][1:]
    T["variant"] = (sr.get("variant") or "").lower()
    return T


def _short(s, n=14):
    s = str(s)
    return s if len(s) <= n else s[: n - 1] + "…"


def _thin(ax, labels, max_n=8, rotate=0):
    n = len(labels)
    step = max(1, math.ceil(n / max_n))
    idx = list(range(0, n, step))
    ax.set_xticks(idx)
    ax.set_xticklabels([_short(labels[i]) for i in idx], rotation=rotate,
                       ha="right" if rotate else "center")


def _style_ax(ax, T, grid="y"):
    ax.set_facecolor(T["card"])
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.tick_params(colors=T["muted"], labelsize=8, length=0)
    ax.grid(axis=grid, color=T["muted"], alpha=0.15, linewidth=0.8)
    ax.set_axisbelow(True)


def _yfmt(ax, axis="y"):
    f = FuncFormatter(lambda v, _: fmt_compact(v))
    (ax.yaxis if axis == "y" else ax.xaxis).set_major_formatter(f)


# ----------------------------------------------------------------------------------
# Legacy chart drawing (fallback only)
# ----------------------------------------------------------------------------------
def _draw_chart_legacy(c, T, w_pt, h_pt, path):
    pal = T["palette"]
    fig = plt.figure(figsize=(w_pt / 72, h_pt / 72), dpi=DPI, facecolor=T["card"])
    fig.text(0.04, 0.93, c["title"], color=T["text"], fontsize=11, fontweight="bold", va="top")
    if c.get("note"):
        fig.text(0.96, 0.93, c["note"], color=T["muted"], fontsize=7.5, va="top", ha="right")
    cats = c["categories"]
    name, vals = next(iter(c["series"].items()))
    vals = [v if v is not None else 0 for v in vals]
    ctype = c["type"]
    n = len(vals)

    if ctype in ("donut", "pie"):
        ax = fig.add_axes([0.03, 0.05, 0.48, 0.78])
        ax.set_facecolor(T["card"])
        total = sum(vals) or 1
        colors = [pal[i % len(pal)] for i in range(n)]
        as_pie = T["variant"] in ("pie", "exploded") or ctype == "pie"
        explode = [0.06] + [0] * (n - 1) if T["variant"] == "exploded" else None
        wedges, _ = ax.pie(vals, colors=colors, startangle=90, counterclock=False, explode=explode,
                           wedgeprops=dict(width=1 if as_pie else 0.38, edgecolor=T["card"], linewidth=2))
        if not as_pie and c.get("center"):
            ax.text(0, 0.06, fmt_compact(c["center"]["value"]), ha="center", va="center",
                    color=T["text"], fontsize=15, fontweight="bold")
            ax.text(0, -0.2, c["center"]["label"], ha="center", va="center", color=T["muted"], fontsize=8)
        ax.legend(wedges, [f"{_short(k, 16)}  {v / total * 100:.0f}%" for k, v in zip(cats, vals)],
                  loc="center left", bbox_to_anchor=(1.08, 0.5), frameon=False,
                  labelcolor=T["text"], fontsize=9)
    elif ctype == "hbar":
        ax = fig.add_axes([0.24, 0.08, 0.70, 0.76])
        _style_ax(ax, T, grid="x")
        ypos = list(range(n))
        bars = ax.barh(ypos, vals[::-1], color=pal[0], height=0.62)
        ax.set_yticks(ypos)
        ax.set_yticklabels([_short(k, 20) for k in cats[::-1]])
        ax.bar_label(bars, labels=[fmt_compact(v) for v in vals[::-1]], color=T["text"], fontsize=8, padding=3)
        _yfmt(ax, "x")
    elif ctype == "scatter":
        ax = fig.add_axes([0.12, 0.17, 0.84, 0.65])
        _style_ax(ax, T, grid="both")
        ax.scatter(cats, vals, s=14, color=pal[0], alpha=0.65, edgecolors="none")
        try:
            import numpy as np
            m, b = np.polyfit(cats, vals, 1)
            xs = [min(cats), max(cats)]
            ax.plot(xs, [m * x + b for x in xs], color=pal[2 % len(pal)], linewidth=1.6)
        except Exception:
            pass
        ax.set_xlabel(c.get("x_name", ""), color=T["muted"], fontsize=8)
        ax.set_ylabel(name, color=T["muted"], fontsize=8)
        _yfmt(ax, "y")
        _yfmt(ax, "x")
    elif ctype == "line":
        ax = fig.add_axes([0.10, 0.16, 0.86, 0.66])
        _style_ax(ax, T)
        x = list(range(n))
        ax.plot(x, vals, color=pal[0], linewidth=2.4, solid_capstyle="round")
        ax.fill_between(x, vals, color=pal[0], alpha=0.15)
        ax.scatter([x[-1]], [vals[-1]], color=pal[0], s=28, zorder=3)
        ax.annotate(fmt_compact(vals[-1]), (x[-1], vals[-1]), textcoords="offset points",
                    xytext=(0, 8), ha="center", color=T["text"], fontsize=8)
        ax.set_ylim(bottom=min(0, min(vals)))
        ax.margins(x=0.03)
        _thin(ax, cats)
        _yfmt(ax)
    else:  # bar, histogram (fallback for unknown types)
        ax = fig.add_axes([0.10, 0.20 if ctype == "bar" else 0.16, 0.86, 0.62])
        _style_ax(ax, T)
        x = list(range(n))
        hist = ctype == "histogram"
        colors = [pal[0]] * n if hist else [pal[0]] + [pal[0]] * (n - 1)
        bars = ax.bar(x, vals, color=colors, width=0.92 if hist else 0.62)
        if not hist:
            for b in bars[1:]:
                b.set_alpha(0.55)
            ax.bar_label(bars, labels=[fmt_compact(v) for v in vals], color=T["text"], fontsize=8, padding=2)
        rotate = 25 if (not hist and max(len(str(k)) for k in cats) > 7 and n > 4) else 0
        _thin(ax, cats, max_n=8 if hist else 10, rotate=rotate)
        _yfmt(ax)
    fig.savefig(path, dpi=DPI, facecolor=T["card"])
    plt.close(fig)


# ----------------------------------------------------------------------------------
# Chart drawing via the Charts package (with fallback)
# ----------------------------------------------------------------------------------
def _draw_chart(c, T, w_pt, h_pt, path):
    """Use dashboard.Charts.render. On any error: log it and use the legacy drawing."""
    try:
        from dashboard.Charts import render
        render(c, T, w_pt, h_pt, path, T.get("variant") or None)
        if not os.path.exists(path):
            raise RuntimeError(f"Charts.render did not create {path}")
    except Exception:
        plt.close("all")
        try:
            with CHART_LOG.open("a", encoding="utf-8") as f:
                f.write(f"--- chart '{c.get('title')}' type={c.get('type')} ---\n")
                f.write(traceback.format_exc() + "\n")
        except Exception:
            pass
        _draw_chart_legacy(c, T, w_pt, h_pt, path)


def _fmt_kpi(k):
    v, f = k["value"], k["format"]
    if v is None:
        return "-"
    if f == "int":
        return f"{int(round(v)):,}"
    if f == "percent":
        return f"{v:.1f}%"
    return fmt_compact(v)


def _draw_kpi(k, T, w_pt, h_pt, path):
    fig = plt.figure(figsize=(w_pt / 72, h_pt / 72), dpi=DPI, facecolor=T["card"])
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    ax.add_patch(Rectangle((0, 0.18), 0.012, 0.64, transform=ax.transAxes, color=T["palette"][0]))
    ax.text(0.07, 0.80, k["label"], color=T["muted"], fontsize=9, va="center", transform=ax.transAxes)
    val = _fmt_kpi(k)
    ax.text(0.07, 0.48, val, color=T["text"], fontsize=20 if len(val) <= 7 else 15,
            fontweight="bold", va="center", transform=ax.transAxes)
    if k.get("delta") is not None:
        up = k["delta"] >= 0
        ax.text(0.07, 0.15, f"{'▲' if up else '▼'} {abs(k['delta']):.1f}%  {k.get('delta_label') or ''}".strip(),
                color=T["up"] if up else T["down"], fontsize=7.5, va="center", transform=ax.transAxes)
    elif k.get("sub"):
        ax.text(0.07, 0.15, _short(k["sub"], 24), color=T["muted"], fontsize=8, va="center",
                transform=ax.transAxes)
    fig.savefig(path, dpi=DPI, facecolor=T["card"])
    plt.close(fig)


def _draw_banner(plan, T, w_pt, h_pt, path):
    fig = plt.figure(figsize=(w_pt / 72, h_pt / 72), dpi=DPI, facecolor=T["card"])
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    ax.add_patch(Rectangle((0, 0), 1, 0.06, transform=ax.transAxes, color=T["palette"][0]))
    ax.text(0.02, 0.60, plan["title"], color=T["text"], fontsize=20, fontweight="bold",
            va="center", transform=ax.transAxes)
    ax.text(0.02, 0.28, f"{plan.get('subtitle', '')}  ·  {date.today():%d %b %Y}", color=T["muted"],
            fontsize=9, va="center", transform=ax.transAxes)
    fig.savefig(path, dpi=DPI, facecolor=T["card"])
    plt.close(fig)


def _draw_insights(plan, T, w_pt, h_pt, path):
    fig = plt.figure(figsize=(w_pt / 72, h_pt / 72), dpi=DPI, facecolor=T["card"])
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    ax.add_patch(Rectangle((0, 0.1), 0.006, 0.8, transform=ax.transAxes, color=T["palette"][0]))
    ax.text(0.02, 0.84, "Key insights", color=T["text"], fontsize=11, fontweight="bold",
            va="center", transform=ax.transAxes)
    for i, line in enumerate(plan["insights"][:4]):
        ax.text(0.02, 0.62 - i * 0.17, f"•  {line}", color=T["muted"], fontsize=9,
                va="center", transform=ax.transAxes)
    fig.savefig(path, dpi=DPI, facecolor=T["card"])
    plt.close(fig)


def _box(pos):
    return (LEFT0 + pos["col"] * CELL_W + GAP / 2, TOP0 + pos["row"] * ROW_H + GAP / 2,
            pos["w"] * CELL_W - GAP, pos["h"] * ROW_H - GAP)


def _make_images(plan, T, out_dir):
    os.makedirs(out_dir, exist_ok=True)
    try:
        CHART_LOG.unlink()          # stale log would mislead; only errors of THIS run stay
    except FileNotFoundError:
        pass
    except Exception:
        pass
    specs = []

    def spec(name, pos, fn, *args):
        left, top, w, h = _box(pos)
        path = os.path.join(out_dir, f"{name}.png")
        fn(*args, w, h, path)
        specs.append({"name": name, "path": path, "left": left, "top": top, "width": w, "height": h})

    spec("banner", plan["banner"], _draw_banner, plan, T)
    for i, k in enumerate(plan["kpis"]):
        spec(f"kpi_{i + 1}", k["pos"], _draw_kpi, k, T)
    for i, c in enumerate(plan["charts"]):
        spec(f"chart_{i + 1}", c["pos"], _draw_chart, c, T)
    if plan.get("insights") and plan.get("insights_pos"):
        spec("insights", plan["insights_pos"], _draw_insights, plan, T)
    return specs


def render_to_folder(plan, out_dir, theme=None):
    """Test mode: only saves PNGs (no Excel)."""
    T = _theme(plan, theme)
    return [s["path"] for s in _make_images(plan, T, out_dir)]


def _rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def render_dashboard(plan, workbook_name=None, sheet_name=None, theme=None):
    """Draw the plan on an Excel sheet as pictures."""
    if xw is None:
        raise RuntimeError("xlwings not installed")
    T = _theme(plan, theme)
    specs = _make_images(plan, T, tempfile.mkdtemp(prefix="dash_"))

    book = None
    if workbook_name:
        for b in xw.books:
            if workbook_name in (b.name, b.fullname):
                book = b
                break
    book = book or xw.books.active
    sname = sheet_name or plan.get("sheet_name") or "Dashboard"
    names = [s.name for s in book.sheets]
    sheet = book.sheets[sname] if sname in names else book.sheets.add(sname, after=book.sheets[-1])

    for coll in (sheet.pictures, sheet.charts):
        try:
            for obj in list(coll):
                obj.delete()
        except Exception:
            pass
    try:
        sheet.cells.clear()
        sheet.cells.color = _rgb(T["bg"])
    except Exception:
        sheet.range("A1:AZ300").color = _rgb(T["bg"])
    try:
        sheet.activate()
        if platform.system() == "Darwin":
            book.app.api.active_window.display_gridlines.set(False)
        else:
            book.app.api.ActiveWindow.DisplayGridlines = False
    except Exception:
        pass

    for s in specs:
        sheet.pictures.add(s["path"], name=s["name"], left=s["left"], top=s["top"],
                           width=s["width"], height=s["height"])
    return {"sheet": sname, "pictures": len(specs), "theme": T["name"]}