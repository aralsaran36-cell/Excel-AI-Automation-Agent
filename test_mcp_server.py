"""Level 3a: checks the MCP server registered 18 tools with correct names and schemas.
No Excel needed. Run: python test_mcp_server.py
"""
import asyncio

from mcp_layer.server import mcp

EXPECTED = {
    "read_excel_data", "get_excel_sheet_names", "write_excel_data", "append_excel_data",
    "remove_excel_duplicates", "remove_excel_blank_rows", "apply_excel_formula",
    "apply_excel_formula_range", "format_excel_cells", "autofit_excel_columns",
    "highlight_excel_values", "filter_excel_data", "clear_excel_filter",
    "create_excel_chart", "analyze_excel_column", "calculate_excel_total",
    "validate_excel_data", "create_excel_pivot_summary",
}


async def main() -> None:
    tools = await mcp.list_tools()
    names = {t.name for t in tools}
    for t in sorted(tools, key=lambda x: x.name):
        params = list(t.inputSchema.get("properties", {}))
        print(f"{t.name:<28} params={params}")
    print()
    print("count == 18      :", len(tools) == 18)
    print("names match      :", names == EXPECTED, "| missing:", EXPECTED - names, "| extra:", names - EXPECTED)
    bad = [t.name for t in tools if set(t.inputSchema.get("properties", {})) <= {"args", "kwargs"}]
    print("bad schemas      :", bad or "none")
    empty = [t.name for t in tools if not (t.description or "").strip()]
    print("no description   : ", empty or "none")


asyncio.run(main())