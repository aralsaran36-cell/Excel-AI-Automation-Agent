from .base import series_items, pal, short, fmt_compact


def draw(fig, c, T, variant=""):
    cats = c["categories"]
    vals = series_items(c)[0][1]
    n = len(vals)
    total = sum(vals) or 1
    cols = [pal(T, i) for i in range(n)]
    ax = fig.add_axes([0.03, 0.05, 0.48, 0.78])
    ax.set_facecolor(T["card"])
    width = 0.22 if variant == "thin" else 0.38
    if variant == "semi":
        w, _ = ax.pie(vals + [total], colors=cols + ["none"], startangle=180, counterclock=False,
                      wedgeprops=dict(width=width, edgecolor=T["card"], linewidth=2))
        w = w[:n]
        ax.set_ylim(-0.1, 1.1)
        cy = 0.1
    else:
        w, _ = ax.pie(vals, colors=cols, startangle=90, counterclock=False,
                      wedgeprops=dict(width=width, edgecolor=T["card"], linewidth=2))
        cy = 0.06
    ctr = c.get("center") or {"value": total, "label": "Total"}
    ax.text(0, cy, fmt_compact(ctr["value"]), ha="center", va="center", color=T["text"],
            fontsize=15, fontweight="bold")
    ax.text(0, cy - 0.26, ctr["label"], ha="center", va="center", color=T["muted"], fontsize=8)
    ax.legend(w, [f"{short(k, 16)}  {v / total * 100:.0f}%" for k, v in zip(cats, vals)],
              loc="center left", bbox_to_anchor=(1.08, 0.5), frameon=False,
              labelcolor=T["text"], fontsize=9)