import sys

from excel_tools.workbook_manager import (
    ExcelToolError,
    backup_workbook,
    get_sheet,
    safe_tool,
    validate_address,
)

IS_WINDOWS = sys.platform.startswith("win")
BACKUP_NOTE = "Backup is a copy of the last saved file. Unsaved edits are not in it."


# ---------- helpers ----------

def _clean(address: str) -> str:
    return str(address).replace("$", "")


def _grid(rng) -> list:
    return rng.options(ndim=2).value


def _filled(value) -> bool:
    return value not in (None, "")


def _count_data(rng) -> int:
    return sum(1 for row in _grid(rng) for c in row if _filled(c))


def _overlaps(a, b) -> bool:
    return not (a.last_cell.row < b.row or b.last_cell.row < a.row or
                a.last_cell.column < b.column or b.last_cell.column < a.column)


def _center(rng) -> None:
    if IS_WINDOWS:
        rng.api.HorizontalAlignment = -4108
    else:
        from appscript import k  # Mac (untested: send the traceback if it fails)
        rng.api.horizontal_alignment.set(k.horizontal_align_center)


def _clear_clipboard(book) -> None:
    try:
        book.app.api.CutCopyMode = False
    except Exception:
        pass


# ---------- tools ----------

@safe_tool
def copy_excel_range(source_range: str, destination_cell: str, move: bool = False,
                     values_only: bool = False, overwrite: bool = False,
                     workbook_name: str | None = None, sheet_name: str | None = None,
                     destination_sheet: str | None = None) -> dict:
    """
    PURPOSE:
        Copy or move a range to another place (same sheet or another sheet).

    PARAMETERS:
        source_range: Range to copy, example A1:D10.
        destination_cell: Top-left cell of the target, example H2.
        move: true to move (the source is cleared).
        values_only: true to copy only values, without formatting (not with move).
        overwrite: true to replace existing data at the destination (a backup is made).
        workbook_name: Workbook name (optional).
        sheet_name: Source sheet name (optional).
        destination_sheet: Target sheet name (optional, default same sheet).

    RETURNS:
        success, action, from, to, to_sheet, backup, note.
    """
    src_addr = validate_address(source_range)
    dst_addr = validate_address(destination_cell)
    if ":" in dst_addr:
        raise ExcelToolError("destination_cell must be a single cell (example: H2).")
    if move and values_only:
        raise ExcelToolError("values_only cannot be used with move.")

    src_sheet = get_sheet(workbook_name, sheet_name)
    dst_sheet = get_sheet(workbook_name, destination_sheet) if destination_sheet else src_sheet
    src = src_sheet.range(src_addr)
    dst_top = dst_sheet.range(dst_addr)
    dst = dst_top.resize(src.rows.count, src.columns.count)
    same_sheet = src_sheet.name == dst_sheet.name

    if same_sheet and src.address == dst.address:
        raise ExcelToolError("Source and destination are the same.")
    if same_sheet and _overlaps(src, dst):
        raise ExcelToolError("Source and destination overlap. Choose another destination.")

    existing = _count_data(dst)
    if existing and not overwrite:
        raise ExcelToolError(
            f"{_clean(dst.address)} already has {existing} data cell(s). Use overwrite=True to replace them.")

    backup = backup_workbook(src_sheet.book) if (existing or move) else None

    if values_only:
        dst.value = _grid(src)
    else:
        src.copy(destination=dst_top)
        _clear_clipboard(src_sheet.book)
    if move:
        src.clear()

    out = {"success": True, "action": "moved" if move else "copied", "from": _clean(src.address),
           "to": _clean(dst.address), "to_sheet": dst_sheet.name, "backup": backup}
    if move:
        out["note"] = ("Move = copy + clear. Formulas that point to the moved cells are not updated. "
                       + BACKUP_NOTE)
    elif backup:
        out["note"] = BACKUP_NOTE
    return out


@safe_tool
def merge_excel_cells(cell_range: str, across: bool = False, center: bool = False,
                      confirm: bool = False, workbook_name: str | None = None,
                      sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Merge cells into one. Only the top-left value is kept. If other values would be lost, the first call shows a preview and the second call with confirm=true merges.

    PARAMETERS:
        cell_range: Range, example A1:D1.
        across: true to merge each row separately (needs 2 or more columns).
        center: true to center the text after merging (good for titles).
        confirm: Must be true when values would be lost.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, merged, range, values_lost, backup.
    """
    addr = validate_address(cell_range)
    if ":" not in addr:
        raise ExcelToolError("Merge needs a range (example: A1:C1).")
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(addr)
    if rng.count < 2:
        raise ExcelToolError("At least 2 cells are needed to merge.")
    if across and rng.columns.count < 2:
        raise ExcelToolError("'across' needs a range with 2 or more columns.")

    grid = _grid(rng)
    total = sum(1 for row in grid for c in row if _filled(c))
    kept_cells = [row[0] for row in grid] if across else [grid[0][0]]
    lost = total - sum(1 for c in kept_cells if _filled(c))

    if lost > 0 and not confirm:
        return {"success": True, "merged": False, "needs_confirm": True,
                "range": _clean(rng.address), "values_would_be_lost": lost,
                "message": f"Merging will lose {lost} value(s) (only the top-left value is kept). Confirm to merge."}

    backup = backup_workbook(sheet.book) if lost > 0 else None

    app = sheet.book.app
    try:
        app.display_alerts = False
        if IS_WINDOWS:
            rng.api.Merge(bool(across))
        elif across:
            rng.api.merge(across=True)  # Mac (untested)
        else:
            rng.api.merge()
    finally:
        app.display_alerts = True
    if center:
        _center(rng)

    out = {"success": True, "merged": True, "range": _clean(rng.address),
           "across": bool(across), "values_lost": max(lost, 0), "backup": backup}
    if backup:
        out["note"] = BACKUP_NOTE
    return out


@safe_tool
def unmerge_excel_cells(cell_range: str, workbook_name: str | None = None,
                        sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Unmerge merged cells in a range. The value stays in the top-left cell.

    PARAMETERS:
        cell_range: Range, example A1:D1.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, unmerged.
    """
    addr = validate_address(cell_range)
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(addr)
    if IS_WINDOWS:
        rng.api.UnMerge()
    else:
        rng.api.unmerge()
    return {"success": True, "range": _clean(rng.address), "unmerged": True}


@safe_tool
def copy_excel_format(source_range: str, destination_range: str,
                      workbook_name: str | None = None, sheet_name: str | None = None,
                      destination_sheet: str | None = None) -> dict:
    """
    PURPOSE:
        Copy only the formatting (font, colors, borders, number format) from one range to another, like Format Painter. Values are not changed.

    PARAMETERS:
        source_range: Range to copy the format from, example A1:D1.
        destination_range: Range or top-left cell to apply it to, example A10:D10.
        workbook_name: Workbook name (optional).
        sheet_name: Source sheet name (optional).
        destination_sheet: Target sheet name (optional, default same sheet).

    RETURNS:
        success, from, to, to_sheet.
    """
    src_addr = validate_address(source_range)
    dst_addr = validate_address(destination_range)
    src_sheet = get_sheet(workbook_name, sheet_name)
    dst_sheet = get_sheet(workbook_name, destination_sheet) if destination_sheet else src_sheet
    src = src_sheet.range(src_addr)
    dst = dst_sheet.range(dst_addr)
    if src_sheet.name == dst_sheet.name and src.address == dst.address:
        raise ExcelToolError("Source and destination are the same.")
    src.copy()
    try:
        dst.paste(paste="formats")
    finally:
        _clear_clipboard(src_sheet.book)
    return {"success": True, "from": _clean(src.address), "to": _clean(dst.address),
            "to_sheet": dst_sheet.name}


@safe_tool
def clear_excel_formatting(cell_range: str, workbook_name: str | None = None,
                           sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Remove all formatting (font, colors, borders, number format) from a range. Values and formulas are kept. Applies directly with a backup.

    PARAMETERS:
        cell_range: Range, example A1:D10.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, backup, note.
    """
    addr = validate_address(cell_range)
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(addr)
    backup = backup_workbook(sheet.book)
    rng.clear_formats()
    return {"success": True, "range": _clean(rng.address), "backup": backup,
            "note": "Number formats were reset to General too. " + BACKUP_NOTE}