from .base import style_ax, compact_axis, thin_labels, series_items, fmt_compact, pal


def draw(fig, c, T, variant=""):
    cats, items = c["categories"], series_items(c)
    n = len(cats)
    ax = fig.add_axes([0.10, 0.16, 0.86, 0.66])
    style_ax(ax, T)
    x = list(range(n))
    allv = []
    for j, (name, vals) in enumerate(items):
        col = pal(T, j)
        ax.plot(x, vals, color=col, linewidth=2.4, solid_capstyle="round", label=name,
                marker="o" if variant == "markers" else None, markersize=4)
        if len(items) == 1:
            ax.fill_between(x, vals, color=col, alpha=0.15)
            ax.scatter([x[-1]], [vals[-1]], color=col, s=28, zorder=3)
            ax.annotate(fmt_compact(vals[-1]), (x[-1], vals[-1]), textcoords="offset points",
                        xytext=(0, 8), ha="center", color=T["text"], fontsize=8)
        allv += vals
    if len(items) > 1:
        ax.legend(frameon=False, labelcolor=T["text"], fontsize=8, ncol=len(items), loc="upper left")
    ax.set_ylim(bottom=min(0, min(allv)))
    ax.margins(x=0.03)
    thin_labels(ax, cats)
    compact_axis(ax)