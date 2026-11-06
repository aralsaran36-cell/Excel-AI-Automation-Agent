from matplotlib.patches import Wedge
from .base import series_items, pal, fmt_compact


def draw(fig, c, T, variant=""):
    value = c.get("value")
    if value is None:
        value = series_items(c)[0][1][0]
    top = c.get("max") or max(value * 1.25, 1)
    frac = max(0.0, min(value / top, 1.0))
    ax = fig.add_axes([0.08, 0.05, 0.84, 0.78])
    ax.set_facecolor(T["card"])
    ax.add_patch(Wedge((0, 0), 1, 0, 180, width=0.26, color=T["muted"], alpha=0.2))
    ax.add_patch(Wedge((0, 0), 1, 180 - 180 * frac, 180, width=0.26, color=pal(T, 0)))
    ax.text(0, 0.18, fmt_compact(value), ha="center", va="center", color=T["text"],
            fontsize=20, fontweight="bold")
    ax.text(0, -0.05, f"{frac * 100:.0f}% of {fmt_compact(top)}", ha="center", va="center",
            color=T["muted"], fontsize=8)
    ax.set_xlim(-1.15, 1.15)
    ax.set_ylim(-0.2, 1.1)
    ax.set_aspect("equal")
    ax.axis("off")