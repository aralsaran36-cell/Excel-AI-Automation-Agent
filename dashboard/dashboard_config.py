"""Dashboard config: every constant in one place. No logic, no imports from other dashboard files."""
from pathlib import Path

# ---- grid / geometry ----
GRID = 24                 # half-columns; every tile keeps 1 unit as gutter
GRID_LEFT = 2             # col 1 = left margin
TOP_ROW = 4               # rows 1-2 banner, 3 gap
COL_WIDTH = 6.5
MARGIN_WIDTH = 2
ROW_H = 18
KPI_VALUE_ROW_H = 32
BANNER_ROW_H = 34
MIN_CHART_W = 5           # smallest readable chart tile (units)
MIN_CHART_H = 8           # rows

# ---- helper data on the dashboard sheet ----
BACK_COL = 40             # chart source data, masked (white font)
META_COL = 60             # saved dashboard settings, masked
META_PREFIX = "AUTO_DASHBOARD_META:"
SCATTER_MAX_POINTS = 500
HIST_BINS = 8

# ---- sheets that are never a data source ----
GENERATED_PREFIXES = ("dashboard", "filtered", "pivot")

# ---- chart selection ----
MAX_CHARTS = 6
MIN_TARGET = 3
MIN_CHARTS = 1
MIN_SCORE = 50
PIE_MAX_CATEGORIES = 5
BAR_MAX_CATEGORIES = 15
TOP_N = 10
SCATTER_ENABLED = False   # scatter plotted wrong on Mac (x = row number)

# ---- insights / history ----
MAX_INSIGHTS = 3
HISTORY_FILE = Path.home() / ".auto_dashboard_history.json"