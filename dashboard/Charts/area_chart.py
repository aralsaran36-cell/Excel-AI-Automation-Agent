from .base import style_ax, compact_axis, thin_labels, series_items, pal


def draw(fig, c, T, variant=""):
    cats, items = c["categories"], series_items(c)
    n = len(cats)
    ax = fig.add_axes([0.10, 0.16, 0.86, 0.66])
    style_ax(ax, T)
    x = list(range(n))
    if variant == "stacked" and len(items) > 1:
        ax.stackplot(x, [v for _, v in items], colors=[pal(T, j) for j in range(len(items))],
                     alpha=0.85, labels=[k for k, _ in items])
    else:
        for j, (name, vals) in enumerate(items):
            ax.fill_between(x, vals, color=pal(T, j), alpha=0.35, label=name)
            ax.plot(x, vals, color=pal(T, j), linewidth=2.2)
    if len(items) > 1:
        ax.legend(frameon=False, labelcolor=T["text"], fontsize=8, ncol=len(items), loc="upper left")
    ax.set_ylim(bottom=0)
    ax.margins(x=0.02)
    thin_labels(ax, cats)
    compact_axis(ax)