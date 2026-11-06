from matplotlib.patches import Rectangle
from .base import series_items, pal, short, fmt_compact


def _split(items, x, y, w, h, out):
    if len(items) == 1:
        out.append((items[0], x, y, w, h))
        return
    total = sum(v for _, v in items)
    acc, k = 0, 1
    for i, (_, v) in enumerate(items):
        acc += v
        if acc >= total / 2:
            k = i + 1
            break
    k = min(max(k, 1), len(items) - 1)
    a, b = items[:k], items[k:]
    fa = sum(v for _, v in a) / total
    if w >= h:
        _split(a, x, y, w * fa, h, out)
        _split(b, x + w * fa, y, w * (1 - fa), h, out)
    else:
        _split(a, x, y, w, h * fa, out)
        _split(b, x, y + h * fa, w, h * (1 - fa), out)


def draw(fig, c, T, variant=""):
    pairs = [(str(k), v) for k, v in zip(c["categories"], series_items(c)[0][1]) if v and v > 0]
    pairs = sorted(pairs, key=lambda p: -p[1])[:12]
    ax = fig.add_axes([0.03, 0.04, 0.94, 0.80])
    ax.set_facecolor(T["card"])
    ax.axis("off")
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    if not pairs:
        return
    out = []
    _split(pairs, 0, 0, 100, 100, out)
    total = sum(v for _, v in pairs)
    for i, ((name, v), x, y, w, h) in enumerate(out):
        ax.add_patch(Rectangle((x, y), w, h, facecolor=pal(T, i), edgecolor=T["card"], linewidth=2.5))
        if w > 12 and h > 12:
            ax.text(x + 2, y + h - 3, short(name, int(w * 0.35)), color="white", fontsize=8,
                    fontweight="bold", va="top")
            ax.text(x + 2, y + 3, f"{fmt_compact(v)}  {v / total * 100:.0f}%", color="white",
                    fontsize=7.5, va="bottom")