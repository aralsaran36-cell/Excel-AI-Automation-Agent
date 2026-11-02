"""MCP server: registers the 18 verified Excel tools.

Run (stdio): python -m mcp_layer.server
Never print() to stdout here: stdout is the MCP protocol channel.
"""
from mcp.server.fastmcp import FastMCP

from excel_tools.analysis import analyze_excel_column, calculate_excel_total
from excel_tools.charts import create_excel_chart
from excel_tools.cleaning import remove_excel_blank_rows, remove_excel_duplicates
from excel_tools.conditional_formatting import highlight_excel_values
from excel_tools.filter_sort import clear_excel_filter, filter_excel_data
from excel_tools.formatting import autofit_excel_columns, format_excel_cells
from excel_tools.formulas import apply_excel_formula, apply_excel_formula_range
from excel_tools.pivot_tables import create_excel_pivot_summary
from excel_tools.reader import get_excel_sheet_names, read_excel_data
from excel_tools.validator import validate_excel_data
from excel_tools.writer import append_excel_data, write_excel_data

mcp = FastMCP("Excel AI Agent")

TOOLS = [
    read_excel_data,
    get_excel_sheet_names,
    write_excel_data,
    append_excel_data,
    remove_excel_duplicates,
    remove_excel_blank_rows,
    apply_excel_formula,
    apply_excel_formula_range,
    format_excel_cells,
    autofit_excel_columns,
    highlight_excel_values,
    filter_excel_data,
    clear_excel_filter,
    create_excel_chart,
    analyze_excel_column,
    calculate_excel_total,
    validate_excel_data,
    create_excel_pivot_summary,
]

for _tool in TOOLS:
    mcp.tool()(_tool)

if __name__ == "__main__":
    mcp.run()