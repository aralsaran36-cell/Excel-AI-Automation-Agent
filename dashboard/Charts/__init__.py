"""Chart registry: chart type -> draw function. render() is the only entry point."""
from . import (area_chart, bar_chart, box_chart, column_chart, doughnut_chart, funnel_chart,
               gauge_chart, histogram_chart, line_chart, pie_chart, radar_chart, scatter_chart,
               stacked_bar_100_chart, treemap_chart)
from .base import new_fig, save

REGISTRY = {
    "bar": column_chart.draw, "column": column_chart.draw, "hbar": bar_chart.draw,
    "line": line_chart.draw, "area": area_chart.draw,
    "pie": pie_chart.draw, "donut": doughnut_chart.draw, "doughnut": doughnut_chart.draw,
    "histogram": histogram_chart.draw, "scatter": scatter_chart.draw, "box": box_chart.draw,
    "radar": radar_chart.draw, "funnel": funnel_chart.draw, "gauge": gauge_chart.draw,
    "treemap": treemap_chart.draw, "stacked_bar_100": stacked_bar_100_chart.draw,
}

VARIANTS = {
    "bar": ["clustered", "stacked"], "column": ["clustered", "stacked"], "line": ["markers"],
    "area": ["stacked"], "pie": ["exploded", "pie3d", "pie_of_pie", "bar_of_pie"],
    "donut": ["thin", "semi"],
}


def render(chart, T, w_pt, h_pt, path, variant=None):
    fn = REGISTRY.get(chart["type"], column_chart.draw)
    fig = new_fig(chart, T, w_pt, h_pt)
    fn(fig, chart, T, (variant or T.get("variant") or "").lower())
    return save(fig, path, T)