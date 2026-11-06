from .base import series_items, pal, short, fmt_compact


def draw(fig, c, T, variant=""):
    pairs = sorted(zip(c["categories"], series_items(c)[0][1]), key=lambda p: -p[1])[:8]
    names = [p[0] for p in pairs]
    vals = [p[1] for p in pairs]
    top = vals[0] or 1
    ax = fig.add_axes([0.04, 0.05, 0.92, 0.78])
    ax.set_facecolor(T["card"])
    for i, v in enumerate(vals):
        ax.barh(i, v, left=-v / 2, height=0.78, color=pal(T, 0), alpha=1 - i * 0.09)
        ax.text(0, i, f"{short(names[i], 18)}  {fmt_compact(v)}  ({v / top * 100:.0f}%)",
                ha="center", va="center", color=T["text"], fontsize=8, fontweight="bold")
    ax.set_ylim(len(vals) - 0.4, -0.6)
    ax.set_xlim(-top * 0.55, top * 0.55)
    ax.axis("off")