from excel_tools.workbook_manager import (
    EXCEL_ERRORS, ExcelToolError, backup_workbook, get_sheet, safe_tool, validate_address,
)


def _check_formula(formula: str) -> None:
    if not isinstance(formula, str) or not formula.startswith("=") or len(formula) < 2:
        raise ExcelToolError("Formula '=' la start aagi irukkanum. Example: =SUM(B2:B10)")


def _has_data(rng) -> bool:
    return any(v is not None for row in rng.options(ndim=2).value for v in row)


def _formula_is_error(target) -> bool | None:
    """Ask Excel itself via ISERROR, because xlwings on macOS returns None for error cells.
    Uses an empty neighbour cell temporarily and clears it. Returns None if no free cell."""
    candidates = [target.offset(0, 1)]
    if target.row > 1:
        candidates.append(target.offset(-1, 0))
    for probe in candidates:
        if probe.value is None and not probe.formula:
            try:
                probe.formula = f"=ISERROR({target.address.replace('$', '')})"
                return bool(probe.value)
            finally:
                probe.clear()
    return None


@safe_tool
def apply_excel_formula(cell: str, formula: str, overwrite: bool = False,
                        workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """Put a formula in one cell and return the calculated result."""
    _check_formula(formula)
    address = validate_address(cell)
    if ":" in address:
        raise ExcelToolError("Oru cell kudunga. Range-ku apply_excel_formula_range use pannunga.")
    sheet = get_sheet(workbook_name, sheet_name)
    target = sheet.range(address)
    backup = None
    if target.value is not None:
        if not overwrite:
            raise ExcelToolError(f"{address}-la already value irukku. overwrite=True kudunga.")
        backup = backup_workbook(sheet.book)
    target.formula = formula
    result = target.value
    if isinstance(result, str) and result in EXCEL_ERRORS:
        return {"success": False, "error": f"Formula error {result} in {address}", "formula": formula}
    if result is None:
        is_error = _formula_is_error(target)
        if is_error:
            return {"success": False,
                    "error": f"Formula error in {address} (Excel returned an error value).",
                    "formula": formula}
        if is_error is None:
            return {"success": True, "cell": address, "formula": formula, "result": None,
                    "backup": backup,
                    "warning": "Result empty; could not verify error state (no free neighbour cell)."}
    return {"success": True, "cell": address, "formula": formula, "result": result, "backup": backup}


@safe_tool
def apply_excel_formula_range(target_range: str, formula: str, overwrite: bool = False,
                              workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """Fill a range with a formula written for its first cell; references adjust per row/column."""
    _check_formula(formula)
    address = validate_address(target_range)
    sheet = get_sheet(workbook_name, sheet_name)
    target = sheet.range(address)
    backup = None
    if _has_data(target):
        if not overwrite:
            raise ExcelToolError(f"{address}-la already data irukku. overwrite=True kudunga.")
        backup = backup_workbook(sheet.book)
    target.formula = formula
    values = [v for row in target.options(ndim=2).value for v in row]
    error_count = sum(1 for v in values if isinstance(v, str) and v in EXCEL_ERRORS)
    return {"success": error_count == 0, "range": address, "formula": formula,
            "cells_filled": len(values), "error_cells": error_count,
            "sample_results": values[:5], "backup": backup}