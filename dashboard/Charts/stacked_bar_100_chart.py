from .base import style_ax, thin_labels, series_items, pal
from . import column_chart


def draw(fig, c, T, variant=""):
    items = series_items(c)
    if len(items) < 2:
        return column_chart.draw(fig, c, T, variant)
    cats = c["categories"]
    n = len(cats)
    totals = [sum(v[i] for _, v in items) or 1 for i in range(n)]
    ax = fig.add_axes([0.10, 0.20, 0.86, 0.62])
    style_ax(ax, T)
    bottom = [0] * n
    for j, (name, vals) in enumerate(items):
        pct = [v / t * 100 for v, t in zip(vals, totals)]
        ax.bar(range(n), pct, 0.62, bottom=bottom, color=pal(T, j), label=name, zorder=2)
        bottom = [b + p for b, p in zip(bottom, pct)]
    ax.set_ylim(0, 100)
    ax.set_yticks([0, 25, 50, 75, 100])
    ax.set_yticklabels(["0%", "25%", "50%", "75%", "100%"])
    ax.legend(frameon=False, labelcolor=T["text"], fontsize=8, ncol=len(items), loc="upper right")
    thin_labels(ax, cats, max_n=10)