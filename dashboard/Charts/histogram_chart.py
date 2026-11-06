from .base import style_ax, compact_axis, thin_labels, gradient_bars, series_items, pal


def draw(fig, c, T, variant=""):
    cats = c["categories"]
    vals = series_items(c)[0][1]
    n = len(vals)
    ax = fig.add_axes([0.10, 0.16, 0.86, 0.66])
    style_ax(ax, T)
    gradient_bars(ax, range(n), vals, 0.94, pal(T, 0))
    ax.set_xlim(-0.6, n - 0.4)
    ax.set_ylim(0, max(max(vals), 1) * 1.12)
    thin_labels(ax, cats, max_n=6)
    compact_axis(ax)