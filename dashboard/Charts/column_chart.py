from .base import (style_ax, compact_axis, thin_labels, gradient_bars, series_items,
                   fmt_compact, pal, y_limits)


def draw(fig, c, T, variant=""):
    cats, items = c["categories"], series_items(c)
    n = len(cats)
    ax = fig.add_axes([0.10, 0.20, 0.86, 0.62])
    style_ax(ax, T)
    if len(items) == 1:
        vals = items[0][1]
        gradient_bars(ax, range(n), vals, 0.62, pal(T, 0))
        for i, v in enumerate(vals):
            ax.text(i, v, fmt_compact(v), ha="center", va="bottom" if v >= 0 else "top",
                    color=T["text"], fontsize=8)
        lo, hi = y_limits(vals)
    elif variant == "stacked":
        bottom = [0] * n
        for j, (name, vals) in enumerate(items):
            ax.bar(range(n), vals, 0.62, bottom=bottom, color=pal(T, j), label=name, zorder=2)
            bottom = [b + v for b, v in zip(bottom, vals)]
        lo, hi = 0, max(bottom) * 1.12
        ax.legend(frameon=False, labelcolor=T["text"], fontsize=8, ncol=len(items), loc="upper right")
    else:  # clustered
        k = len(items)
        w = 0.8 / k
        allv = []
        for j, (name, vals) in enumerate(items):
            ax.bar([i - 0.4 + w * (j + 0.5) for i in range(n)], vals, w * 0.9,
                   color=pal(T, j), label=name, zorder=2)
            allv += vals
        lo, hi = y_limits(allv)
        ax.legend(frameon=False, labelcolor=T["text"], fontsize=8, ncol=k, loc="upper right")
    ax.set_xlim(-0.6, n - 0.4)
    ax.set_ylim(lo, hi)
    rotate = 25 if (max(len(str(k)) for k in cats) > 7 and n > 4) else 0
    thin_labels(ax, cats, max_n=10, rotate=rotate)
    compact_axis(ax)