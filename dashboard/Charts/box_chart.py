from .base import style_ax, compact_axis, series_items, pal, short


def draw(fig, c, T, variant=""):
    groups = c.get("groups") or {k: v for k, v in series_items(c)}
    names = list(groups)
    ax = fig.add_axes([0.10, 0.20, 0.86, 0.62])
    style_ax(ax, T)
    bp = ax.boxplot([groups[k] for k in names], patch_artist=True, widths=0.5,
                    medianprops=dict(color=T["text"], linewidth=1.6),
                    whiskerprops=dict(color=T["muted"]), capprops=dict(color=T["muted"]),
                    flierprops=dict(marker="o", markersize=3, markerfacecolor=T["muted"],
                                    markeredgecolor="none", alpha=0.6))
    for i, box in enumerate(bp["boxes"]):
        box.set(facecolor=pal(T, i), alpha=0.75, edgecolor=pal(T, i))
    ax.set_xticklabels([short(k, 12) for k in names])
    compact_axis(ax)