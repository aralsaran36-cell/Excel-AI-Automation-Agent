"""Shared drawing helpers for every chart file. No data decisions here."""
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap, to_rgb
from matplotlib.ticker import FuncFormatter

DPI = 160


def fmt_compact(v):
    if v is None:
        return "-"
    a = abs(v)
    for lim, suf in ((1e9, "B"), (1e6, "M"), (1e3, "K")):
        if a >= lim:
            return f"{v / lim:.1f}".rstrip("0").rstrip(".") + suf
    return f"{v:,.0f}" if a >= 100 else f"{v:,.2f}".rstrip("0").rstrip(".")


def short(s, n=14):
    s = str(s)
    return s if len(s) <= n else s[: n - 1] + "…"


def lighten(color, amt=0.3):
    r, g, b = to_rgb(color)
    return (r + (1 - r) * amt, g + (1 - g) * amt, b + (1 - b) * amt)


def darken(color, amt=0.3):
    r, g, b = to_rgb(color)
    return (r * (1 - amt), g * (1 - amt), b * (1 - amt))


def pal(T, i):
    return T["palette"][i % len(T["palette"])]


def series_items(c):
    """[(name, values)] with None -> 0."""
    return [(str(k), [v if v is not None else 0 for v in vals]) for k, vals in c["series"].items()]


def new_fig(c, T, w_pt, h_pt):
    fig = plt.figure(figsize=(w_pt / 72, h_pt / 72), dpi=DPI, facecolor=T["card"])
    fig.text(0.04, 0.93, c["title"], color=T["text"], fontsize=11, fontweight="bold", va="top")
    if c.get("note"):
        fig.text(0.96, 0.93, c["note"], color=T["muted"], fontsize=7.5, va="top", ha="right")
    return fig


def save(fig, path, T):
    fig.savefig(path, dpi=DPI, facecolor=T["card"])
    plt.close(fig)
    return path


def style_ax(ax, T, grid="y"):
    ax.set_facecolor(T["card"])
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.tick_params(colors=T["muted"], labelsize=8, length=0)
    ax.grid(axis=grid, color=T["muted"], alpha=0.15, linewidth=0.8)
    ax.set_axisbelow(True)


def compact_axis(ax, axis="y"):
    f = FuncFormatter(lambda v, _: fmt_compact(v))
    (ax.yaxis if axis == "y" else ax.xaxis).set_major_formatter(f)


def thin_labels(ax, labels, max_n=8, rotate=0):
    n = len(labels)
    step = max(1, -(-n // max_n))
    idx = list(range(0, n, step))
    ax.set_xticks(idx)
    ax.set_xticklabels([short(labels[i]) for i in idx], rotation=rotate,
                       ha="right" if rotate else "center")


def gradient_bars(ax, pos, sizes, width, color, horizontal=False):
    """Bars with a dark-to-light gradient. Caller must set axis limits afterwards."""
    cmap = LinearSegmentedColormap.from_list("g", [darken(color, 0.30), lighten(color, 0.22)])
    for p, s in zip(pos, sizes):
        if s == 0:
            continue
        if horizontal:
            grad = np.linspace(0, 1, 128).reshape(1, -1)
            ax.imshow(grad, extent=(0, s, p - width / 2, p + width / 2), origin="lower",
                      aspect="auto", cmap=cmap, zorder=2)
        else:
            grad = np.linspace(0, 1, 128).reshape(-1, 1)
            ax.imshow(grad, extent=(p - width / 2, p + width / 2, 0, s), origin="lower",
                      aspect="auto", cmap=cmap, zorder=2)


def y_limits(vals, pad=1.15):
    lo, hi = min(0, min(vals)), max(0, max(vals))
    return lo * pad if lo < 0 else 0, hi * pad if hi > 0 else 1