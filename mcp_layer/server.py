from mcp.server.fastmcp import FastMCP

from excel_tools.analysis import (
    analyze_excel_column,
    calculate_excel_total,
    compare_excel_datasets,
    identify_excel_outliers,
    value_counts_excel_column,
)

from excel_tools.dashboard import (
    auto_dashboard,
    delete_dashboard,
    list_dashboards,
    modify_dashboard,
    refresh_dashboard,
)
from excel_tools.automation import (
    list_workflows, 
    save_workflow, 
    delete_workflow, 
    run_workflow)

from excel_tools.charts import (
    change_excel_chart_data,
    create_excel_chart,
    delete_excel_chart,
    list_excel_charts,
    modify_excel_chart,
)
from excel_tools.cleaning import (
    clean_excel_text,
    fill_missing_excel_values,
    find_replace_excel_data,
    remove_excel_blank_rows,
    remove_excel_duplicates,
)
from excel_tools.conditional_formatting import (
    apply_conditional_formatting,
    apply_data_validation,
    clear_conditional_formatting,
    highlight_excel_values,
    list_conditional_formatting,
)
from excel_tools.filter_sort import (
    apply_excel_autofilter,
    clear_excel_filter,
    filter_excel_data,
    filter_excel_multi,
    sort_excel_data,
)
from excel_tools.formatting import (
    apply_excel_borders,
    autofit_excel_columns,
    format_excel_cells,
    set_excel_alignment,
    set_excel_size,
)
from excel_tools.formulas import (
    add_excel_formula_column,
    apply_excel_formula,
    apply_excel_formula_range,
    convert_excel_formulas_to_values,
    get_excel_formula,
)
from excel_tools.pivot_tables import (
    create_excel_native_pivot,
    create_excel_pivot_summary,
    delete_excel_pivot,
    list_excel_pivots,
    modify_excel_pivot,
)
from excel_tools.range_manager import (
    clear_excel_formatting,
    copy_excel_format,
    copy_excel_range,
    merge_excel_cells,
    unmerge_excel_cells,
)

from excel_tools.reader import (
    get_excel_sheet_names, 
    read_excel_data, 
    get_excel_column_info,
    read_excel_range, 
    search_excel_value
)
from excel_tools.row_column_manager import (
    delete_excel_columns,
    delete_excel_rows,
    hide_unhide_excel_rows_columns,
    insert_excel_columns,
    insert_excel_rows,
)

from excel_tools.validator import validate_excel_data

from excel_tools.worksheet import (
    copy_worksheet, create_worksheet, delete_worksheet, freeze_panes,
    get_all_worksheets, get_worksheet_info, hide_unhide_worksheet,
    move_worksheet, protect_unprotect_worksheet, rename_worksheet,
    set_print_setup, set_sheet_tab_color, set_worksheet_zoom_view,
)
from excel_tools.writer import (
    append_excel_data, 
    clear_excel_range, 
    fill_excel_series,
    update_excel_cell, 
    write_excel_data,
)

mcp = FastMCP("Excel AI Agent")

TOOLS_ANALYSIS = [
    analyze_excel_column, 
    calculate_excel_total, 
    compare_excel_datasets,
    value_counts_excel_column, 
    identify_excel_outliers,
]
TOOLS_AUTOMATION = [
    list_workflows, 
    save_workflow, 
    delete_workflow, 
    run_workflow]

TOOLS_DASHBOARD = [
    auto_dashboard,
    list_dashboards,
    refresh_dashboard,
    modify_dashboard,
    delete_dashboard,
]

TOOLS_CHARTS = [
    change_excel_chart_data,
    create_excel_chart,
    delete_excel_chart,
    list_excel_charts,
    modify_excel_chart,
]
TOOLS_CLEANING = [
    clean_excel_text,
    fill_missing_excel_values,
    find_replace_excel_data,
    remove_excel_blank_rows,
    remove_excel_duplicates,
]
TOOLS_CONDITIONAL_FORMATTING = [
    apply_conditional_formatting,
    apply_data_validation,
    clear_conditional_formatting,
    highlight_excel_values,
    list_conditional_formatting,
]
TOOLS_FILTER_SORT = [
    apply_excel_autofilter,
    clear_excel_filter, 
    filter_excel_data, 
    filter_excel_multi, 
    sort_excel_data]

TOOLS_FORMATTING = [
    apply_excel_borders,
    autofit_excel_columns,
    format_excel_cells,
    set_excel_alignment,
    set_excel_size,
]
TOOLS_FORMULAS = [
    add_excel_formula_column, 
    apply_excel_formula, 
    apply_excel_formula_range, 
    convert_excel_formulas_to_values, 
    get_excel_formula]

TOOLS_PIVOT = [
    create_excel_native_pivot,
    create_excel_pivot_summary, 
    delete_excel_pivot, 
    list_excel_pivots, 
    modify_excel_pivot]

TOOLS_RANGE = [
    clear_excel_formatting,
    copy_excel_format,
    copy_excel_range,
    merge_excel_cells,
    unmerge_excel_cells,
]

TOOLS_READER = [
    get_excel_sheet_names, 
    read_excel_data, 
    get_excel_column_info, 
    read_excel_range, 
    search_excel_value]

TOOLS_ROW_COLUMN = [
    delete_excel_columns,
    delete_excel_rows,
    hide_unhide_excel_rows_columns,
    insert_excel_columns,
    insert_excel_rows,
]

TOOLS_VALIDATOR = [validate_excel_data]

TOOLS_WORKSHEET = [
    copy_worksheet, create_worksheet, delete_worksheet, freeze_panes,
    get_all_worksheets, get_worksheet_info, hide_unhide_worksheet,
    move_worksheet, protect_unprotect_worksheet, rename_worksheet,
    set_print_setup, set_sheet_tab_color, set_worksheet_zoom_view,
]
TOOLS_WRITER = [
    append_excel_data, 
    clear_excel_range, 
    fill_excel_series,
    update_excel_cell, 
    write_excel_data]

TOOLS = (
    TOOLS_ANALYSIS + TOOLS_AUTOMATION + TOOLS_DASHBOARD
    + TOOLS_CHARTS + TOOLS_CLEANING
    + TOOLS_CONDITIONAL_FORMATTING + TOOLS_FILTER_SORT
    + TOOLS_FORMATTING + TOOLS_FORMULAS + TOOLS_PIVOT
    + TOOLS_RANGE + TOOLS_READER + TOOLS_ROW_COLUMN
    + TOOLS_VALIDATOR + TOOLS_WORKSHEET + TOOLS_WRITER
)

for _tool in TOOLS:
    mcp.tool()(_tool)

if __name__ == "__main__":
    mcp.run()