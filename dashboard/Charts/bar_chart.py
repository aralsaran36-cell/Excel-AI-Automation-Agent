from .base import style_ax, compact_axis, gradient_bars, series_items, fmt_compact, pal, short


def draw(fig, c, T, variant=""):
    cats = c["categories"]
    vals = series_items(c)[0][1]
    n = len(vals)
    ax = fig.add_axes([0.24, 0.08, 0.70, 0.76])
    style_ax(ax, T, grid="x")
    ypos = list(range(n))
    rv = vals[::-1]
    gradient_bars(ax, ypos, rv, 0.62, pal(T, 0), horizontal=True)
    for y, v in zip(ypos, rv):
        ax.text(v, y, " " + fmt_compact(v), va="center", color=T["text"], fontsize=8)
    ax.set_yticks(ypos)
    ax.set_yticklabels([short(k, 20) for k in cats[::-1]])
    ax.set_xlim(0, max(max(vals), 1) * 1.18)
    ax.set_ylim(-0.6, n - 0.4)
    compact_axis(ax, "x")