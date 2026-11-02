import pandas as pd

from excel_tools.workbook_manager import (
    ExcelToolError, get_sheet, load_dataframe, safe_tool, write_dataframe_to_sheet,
)

AGGREGATIONS = ("sum", "mean", "count", "min", "max")
PIVOT_PREFIX = "Pivot"


@safe_tool
def create_excel_pivot_summary(group_by: str, value_column: str, aggregation: str = "sum",
                               output_sheet: str = "Pivot_Summary",
                               workbook_name: str | None = None,
                               sheet_name: str | None = None) -> dict:
    """Group by one column and aggregate another (sum, mean, count, min, max).
    Writes a summary table to a new sheet (Pandas-based, not a native PivotTable)."""
    if aggregation not in AGGREGATIONS:
        raise ExcelToolError(f"Invalid aggregation. Use one of: {list(AGGREGATIONS)}")
    sheet = get_sheet(workbook_name, sheet_name)
    if output_sheet == sheet.name:
        raise ExcelToolError("output_sheet source sheet-oda same-a irukka koodathu.")
    df, _ = load_dataframe(sheet)
    for col in (group_by, value_column):
        if col not in df.columns:
            raise ExcelToolError(f"Column '{col}' illa. Available: {list(df.columns)}")
    if group_by == value_column:
        raise ExcelToolError("group_by-um value_column-um vera vera columns-a irukkanum.")
    values = df[value_column]
    if aggregation != "count":
        values = pd.to_numeric(values, errors="coerce")
        if values.notna().sum() == 0:
            raise ExcelToolError(f"'{value_column}' numeric column illa.")
    summary = (pd.DataFrame({group_by: df[group_by], value_column: values})
               .groupby(group_by, dropna=True)[value_column].agg(aggregation).reset_index())
    summary.columns = [group_by, f"{aggregation.capitalize()} of {value_column}"]
    summary = summary.sort_values(summary.columns[1], ascending=False)
    write_dataframe_to_sheet(sheet.book, output_sheet, summary, PIVOT_PREFIX)
    return {"success": True, "output_sheet": output_sheet, "groups": len(summary),
            "top_rows": summary.head(5).to_dict("records")}