from .base import series_items, pal, short, darken, fmt_compact

EDGE = dict(linewidth=2)


def _legend(ax, wedges, cats, vals, T):
    total = sum(vals) or 1
    ax.legend(wedges, [f"{short(k, 16)}  {v / total * 100:.0f}%" for k, v in zip(cats, vals)],
              loc="center left", bbox_to_anchor=(1.08, 0.5), frameon=False,
              labelcolor=T["text"], fontsize=9)


def _of_pie(fig, cats, vals, T, bar):
    order = sorted(range(len(vals)), key=lambda i: -vals[i])
    k = min(3, len(vals) // 2)
    small, big = order[-k:], order[:-k]
    mv = [vals[i] for i in big] + [sum(vals[i] for i in small)]
    mc = [cats[i] for i in big] + ["Other"]
    cols = [pal(T, i) for i in range(len(mv))]
    ax1 = fig.add_axes([0.02, 0.08, 0.42, 0.74])
    w, _ = ax1.pie(mv, colors=cols, startangle=90, counterclock=False,
                   wedgeprops=dict(edgecolor=T["card"], **EDGE))
    _legend(ax1, w, mc, mv, T)
    ax2 = fig.add_axes([0.62, 0.08, 0.34, 0.74])
    sv = [vals[i] for i in small]
    sc = [pal(T, len(big) + j) for j in range(len(sv))]
    if bar:
        ax2.set_facecolor(T["card"])
        bottom = 0
        tot = sum(sv) or 1
        for v, col, i in zip(sv, sc, small):
            ax2.bar(0, v / tot, 0.5, bottom=bottom, color=col, edgecolor=T["card"])
            ax2.text(0.32, bottom + v / tot / 2, short(cats[i], 12), color=T["text"], fontsize=8, va="center")
            bottom += v / tot
        ax2.set_xlim(-0.4, 1.2)
        ax2.axis("off")
    else:
        ax2.pie(sv, colors=sc, startangle=90, counterclock=False,
                labels=[short(cats[i], 10) for i in small], textprops=dict(color=T["text"], fontsize=8),
                wedgeprops=dict(edgecolor=T["card"], **EDGE))


def draw(fig, c, T, variant=""):
    cats = c["categories"]
    vals = series_items(c)[0][1]
    n = len(vals)
    if variant in ("pie_of_pie", "bar_of_pie") and n >= 4:
        return _of_pie(fig, cats, vals, T, bar=variant == "bar_of_pie")
    cols = [pal(T, i) for i in range(n)]
    ax = fig.add_axes([0.03, 0.05, 0.48, 0.78])
    ax.set_facecolor(T["card"])
    if variant == "pie3d":
        for k in range(9, 0, -1):
            ax.pie(vals, colors=[darken(x, 0.5) for x in cols], startangle=90, counterclock=False,
                   center=(0, -k * 0.03), wedgeprops=dict(edgecolor=T["card"], linewidth=0.5))
        w, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                      wedgeprops=dict(edgecolor=T["card"], **EDGE))
        ax.set_xlim(-1.25, 1.25)
        ax.set_ylim(-1.2, 1.2)
        ax.set_aspect(0.62)
    else:
        explode = [0.07] * n if variant == "exploded" else None
        w, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False, explode=explode,
                      wedgeprops=dict(edgecolor=T["card"], **EDGE))
    _legend(ax, w, cats, vals, T)