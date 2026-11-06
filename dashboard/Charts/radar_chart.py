import math
from .base import series_items, pal, short, fmt_compact


def draw(fig, c, T, variant=""):
    cats, items = c["categories"], series_items(c)
    n = len(cats)
    ang = [i / n * 2 * math.pi for i in range(n)] + [0]
    ax = fig.add_axes([0.18, 0.08, 0.64, 0.74], polar=True)
    ax.set_facecolor(T["card"])
    ax.set_theta_offset(math.pi / 2)
    ax.set_theta_direction(-1)
    top = max(max(v) for _, v in items) or 1
    for j, (name, vals) in enumerate(items):
        v = vals + vals[:1]
        ax.plot(ang, v, color=pal(T, j), linewidth=2, label=name)
        ax.fill(ang, v, color=pal(T, j), alpha=0.2)
    ax.set_xticks(ang[:-1])
    ax.set_xticklabels([short(k, 12) for k in cats], color=T["text"], fontsize=8)
    ax.set_ylim(0, top * 1.1)
    ax.set_yticks([top * f for f in (0.25, 0.5, 0.75, 1)])
    ax.set_yticklabels([fmt_compact(top * f) for f in (0.25, 0.5, 0.75, 1)], color=T["muted"], fontsize=6)
    ax.grid(color=T["muted"], alpha=0.25)
    ax.spines["polar"].set_color(T["muted"])
    ax.spines["polar"].set_alpha(0.25)
    if len(items) > 1:
        ax.legend(frameon=False, labelcolor=T["text"], fontsize=8, loc="upper right",
                  bbox_to_anchor=(1.35, 1.1))