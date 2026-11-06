import numpy as np
from .base import style_ax, compact_axis, series_items, pal


def draw(fig, c, T, variant=""):
    xs = [float(x) for x in c["categories"]]
    name, ys = series_items(c)[0]
    ax = fig.add_axes([0.12, 0.17, 0.84, 0.65])
    style_ax(ax, T, grid="both")
    ax.scatter(xs, ys, s=14, color=pal(T, 0), alpha=0.65, edgecolors="none", zorder=2)
    try:
        m, b = np.polyfit(xs, ys, 1)
        lx = [min(xs), max(xs)]
        ax.plot(lx, [m * x + b for x in lx], color=pal(T, 2), linewidth=1.6, zorder=3)
    except Exception:
        pass
    ax.set_xlabel(c.get("x_name", ""), color=T["muted"], fontsize=8)
    ax.set_ylabel(name, color=T["muted"], fontsize=8)
    compact_axis(ax, "x")
    compact_axis(ax, "y")