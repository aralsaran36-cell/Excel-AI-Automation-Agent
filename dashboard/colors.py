"""Single source of truth for all dashboard colours."""

SHADES = ("very_light", "light", "base", "dark", "very_dark")

FAMILIES = {
    "red":    ["#FEE2E2", "#FCA5A5", "#EF4444", "#B91C1C", "#7F1D1D"],
    "orange": ["#FFEDD5", "#FDBA74", "#F97316", "#C2410C", "#7C2D12"],
    "yellow": ["#FEF9C3", "#FDE047", "#EAB308", "#A16207", "#713F12"],
    "green":  ["#DCFCE7", "#86EFAC", "#22C55E", "#15803D", "#14532D"],
    "teal":   ["#CCFBF1", "#5EEAD4", "#14B8A6", "#0F766E", "#134E4A"],
    "blue":   ["#DBEAFE", "#93C5FD", "#3B82F6", "#1D4ED8", "#1E3A8A"],
    "purple": ["#F3E8FF", "#D8B4FE", "#A855F7", "#7E22CE", "#581C87"],
    "pink":   ["#FCE7F3", "#F9A8D4", "#EC4899", "#BE185D", "#831843"],
    "brown":  ["#F5EBE0", "#D6B79A", "#A47148", "#6F4518", "#3E2723"],
    "gray":   ["#F3F4F6", "#D1D5DB", "#6B7280", "#374151", "#111827"],
}

ALIASES = {
    "grey": "gray", "violet": "purple", "cyan": "teal",
    "magenta": "pink", "gold": "yellow", "navy": "blue",
}

_PREFIXES = (("very light", "very_light"), ("very dark", "very_dark"),
             ("light", "light"), ("dark", "dark"))


def shade(family, level="base"):
    """shade('red', 'dark') -> '#B91C1C'"""
    family = ALIASES.get(family, family)
    return FAMILIES[family][SHADES.index(level)]


def family_gradient(family, n=5):
    """n colours from light to dark of one family (for bars, heatmaps)."""
    family = ALIASES.get(family, family)
    cols = FAMILIES[family]
    if n >= len(cols):
        return cols[:]
    step = (len(cols) - 1) / max(n - 1, 1)
    return [cols[round(i * step)] for i in range(n)]


def resolve_color(value, default=None):
    """'dark red' / 'light_blue' / 'red' / '#abc' -> '#RRGGBB'. Unknown -> default."""
    if not value:
        return default
    s = str(value).strip()
    if s.startswith("#"):
        h = s[1:]
        if len(h) == 3:
            h = "".join(c * 2 for c in h)
        if len(h) == 6 and all(c in "0123456789abcdefABCDEF" for c in h):
            return "#" + h.upper()
        return default
    s = " ".join(s.lower().replace("_", " ").replace("-", " ").split())
    level = "base"
    for prefix, lvl in _PREFIXES:
        if s.startswith(prefix + " "):
            level = lvl
            s = s[len(prefix) + 1:]
            break
    s = ALIASES.get(s, s)
    if s in FAMILIES:
        return FAMILIES[s][SHADES.index(level)]
    return default