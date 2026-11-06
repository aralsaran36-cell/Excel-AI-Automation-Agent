"""Worksheet management tools for Excel automation.

Tools:
    create_worksheet, rename_worksheet, delete_worksheet (preview + confirm),
    copy_worksheet, get_all_worksheets, move_worksheet, hide_unhide_worksheet,
    set_sheet_tab_color, protect_unprotect_worksheet, freeze_panes,
    set_worksheet_zoom_view, set_print_setup, get_worksheet_info

All tools return {"success": bool, ...} and never raise (via @safe_tool).
delete_worksheet only previews by default; it needs confirm=True to really delete.
Tools that use sheet.api (COM) need Excel for Windows.
"""
import re
from contextlib import contextmanager

from excel_tools.workbook_manager import ExcelToolError, get_workbook, safe_tool

INVALID_CHARS = set('[]:*?/\\')
MAX_NAME_LEN = 31

_STATE_NAMES = {0: "hidden", 2: "very_hidden"}
_HEX_RE = re.compile(r"^#?([0-9a-fA-F]{6})$")


# ---------------------------------------------------------
# SHARED HELPERS
# ---------------------------------------------------------

def _validate_name(name: str) -> str:
    """Validate an Excel worksheet name against Excel's naming rules."""
    if not isinstance(name, str) or not name.strip():
        raise ExcelToolError("Sheet name cannot be empty.")
    name = name.strip()
    if len(name) > MAX_NAME_LEN:
        raise ExcelToolError(f"Sheet name cannot exceed {MAX_NAME_LEN} characters.")
    bad = sorted(set(name) & INVALID_CHARS)
    if bad:
        raise ExcelToolError(f"Sheet name contains invalid characters: {bad}")
    if name.startswith("'") or name.endswith("'"):
        raise ExcelToolError("Sheet name cannot start or end with an apostrophe.")
    return name


def _names(book) -> list[str]:
    """Return all worksheet names in a workbook."""
    return [s.name for s in book.sheets]


def _find(book, name: str) -> str:
    """Case-insensitive match to a real sheet name."""
    if not isinstance(name, str) or not name.strip():
        raise ExcelToolError("Sheet name cannot be empty.")
    wanted = name.strip().lower()
    for existing in _names(book):
        if existing.lower() == wanted:
            return existing
    raise ExcelToolError(f"Sheet '{name}' does not exist. Available sheets: {_names(book)}")


def _ensure_unique(book, name: str, ignore: str | None = None) -> None:
    """Ensure a worksheet name is not already used (case-insensitive)."""
    taken = [n.lower() for n in _names(book) if n != ignore]
    if name.lower() in taken:
        raise ExcelToolError(f"Sheet '{name}' already exists.")


def _sheet_state(sheet) -> str:
    """Return 'visible', 'hidden' or 'very_hidden'."""
    return _STATE_NAMES.get(int(sheet.api.Visible), "visible")


def _visible_count(book) -> int:
    """Count worksheets that are currently visible."""
    return sum(1 for s in book.sheets if _sheet_state(s) == "visible")


def _require_visible(sheet) -> None:
    """Window-level settings need a visible sheet."""
    if _sheet_state(sheet) != "visible":
        raise ExcelToolError(
            f"Sheet '{sheet.name}' is hidden. Unhide it first with hide_unhide_worksheet.")


def _parse_hex_color(color: str) -> int:
    """Convert '#RRGGBB' to the BGR integer Excel's COM API expects."""
    if not isinstance(color, str) or not _HEX_RE.match(color.strip()):
        raise ExcelToolError("Color must be a hex value like '#FF0000'.")
    h = _HEX_RE.match(color.strip()).group(1)
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


def _bgr_to_hex(value: int) -> str:
    """Convert Excel's BGR integer back to '#RRGGBB'."""
    value = int(value)
    return "#{:02X}{:02X}{:02X}".format(value & 255, (value >> 8) & 255, (value >> 16) & 255)


def _check_int(value, label: str, minimum: int = 0) -> int:
    """Validate that a value is an integer >= minimum (bool rejected)."""
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise ExcelToolError(f"{label} must be a whole number >= {minimum}.")
    return value


@contextmanager
def _activated(book, sheet):
    """Temporarily activate a sheet (needed for window settings), then restore."""
    previous = None
    try:
        book.activate()
        previous = book.sheets.active
    except Exception:
        previous = None
    sheet.activate()
    try:
        yield book.app.api.ActiveWindow
    finally:
        if previous is not None and previous.name != sheet.name:
            try:
                previous.activate()
            except Exception:
                pass


# ---------------------------------------------------------
# 1. CREATE WORKSHEET
# ---------------------------------------------------------

@safe_tool
def create_worksheet(sheet_name: str,
                     workbook_name: str | None = None,
                     after_sheet: str | None = None) -> dict:
    """
    PURPOSE:
        Create a new worksheet inside an existing Excel workbook.

    WHEN TO USE:
        - User asks to add a new worksheet.
        - User wants a separate sheet for new data.
        - User wants a worksheet created after a specific sheet.

    WHEN NOT TO USE:
        - User wants to create an entirely new workbook.
        - User wants to copy an existing worksheet (use copy_worksheet).

    PARAMETERS:
        sheet_name: Name of the new worksheet.
        workbook_name: Target workbook. Uses the active workbook if omitted.
        after_sheet: Optional existing worksheet name. The new worksheet is
            inserted after this sheet. If omitted, it is added at the end.

    RETURNS:
        Created worksheet name and the updated list of sheet names.

    NOTES:
        - Does not overwrite existing worksheets.
        - Worksheet names must follow Excel naming rules.
    """
    name = _validate_name(sheet_name)
    book = get_workbook(workbook_name)
    _ensure_unique(book, name)

    if after_sheet is not None:
        anchor = book.sheets[_find(book, after_sheet)]
    else:
        anchor = book.sheets[-1]

    book.sheets.add(name=name, after=anchor)
    return {"success": True, "workbook": book.name, "created": name,
            "sheets": _names(book)}


# ---------------------------------------------------------
# 2. RENAME WORKSHEET
# ---------------------------------------------------------

@safe_tool
def rename_worksheet(old_sheet_name: str,
                     new_sheet_name: str,
                     workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Rename an existing worksheet.

    WHEN TO USE:
        - User asks to change a worksheet name.
        - User wants a more meaningful sheet name.
        - User wants to correct a worksheet name.

    WHEN NOT TO USE:
        - User wants to create another worksheet (use create_worksheet).
        - User wants to copy a worksheet (use copy_worksheet).

    PARAMETERS:
        old_sheet_name: Current worksheet name (case-insensitive).
        new_sheet_name: New worksheet name.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Old name, new name and the updated list of sheet names.

    NOTES:
        - Existing worksheet data is preserved.
        - New name must be unique within the workbook.
    """
    book = get_workbook(workbook_name)
    old = _find(book, old_sheet_name)
    new = _validate_name(new_sheet_name)
    _ensure_unique(book, new, ignore=old)

    book.sheets[old].name = new
    return {"success": True, "workbook": book.name, "old_name": old,
            "new_name": new, "sheets": _names(book)}


# ---------------------------------------------------------
# 3. DELETE WORKSHEET (preview + confirm)
# ---------------------------------------------------------

@safe_tool
def delete_worksheet(sheet_name: str,
                     confirm: bool = False,
                     workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Permanently delete an existing worksheet from a workbook.
        Without confirm=True it only previews and deletes nothing.

    WHEN TO USE:
        - User explicitly asks to delete or remove a worksheet.

    WHEN NOT TO USE:
        - User wants to clear cell contents but keep the worksheet.
        - User wants to hide a worksheet temporarily.

    PARAMETERS:
        sheet_name: Name of the worksheet to delete (case-insensitive).
        confirm: False (default) = preview only. True = really delete.
            Set True only after the user has clearly confirmed.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Preview (deleted=False) with the size of the data area, or the
        deleted sheet name and remaining sheets (deleted=True).

    NOTES:
        - Deletion is permanent unless the user uses Excel Undo or a backup.
        - The last remaining worksheet cannot be deleted.
    """
    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    if len(book.sheets) <= 1:
        raise ExcelToolError("Cannot delete the last remaining worksheet.")

    sheet = book.sheets[target]
    if _sheet_state(sheet) == "visible" and _visible_count(book) <= 1:
        raise ExcelToolError("Cannot delete the last visible worksheet. Unhide another sheet first.")

    if not confirm:
        used = sheet.used_range
        cells = 0 if (used.count == 1 and used.value is None) else int(used.count)
        return {"success": True, "deleted": False, "sheet": target,
                "cells_in_used_range": cells,
                "message": f"'{target}' has {cells} cells in its used range. "
                           f"Call again with confirm=True to delete it."}

    app = book.app
    app.display_alerts = False
    try:
        sheet.delete()
    finally:
        app.display_alerts = True
    return {"success": True, "deleted": True, "sheet": target,
            "sheets": _names(book)}


# ---------------------------------------------------------
# 4. COPY WORKSHEET
# ---------------------------------------------------------

@safe_tool
def copy_worksheet(source_sheet_name: str,
                   new_sheet_name: str,
                   workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Create a copy of an existing worksheet in the same workbook.

    WHEN TO USE:
        - User wants to duplicate a worksheet.
        - User wants a backup copy before making changes.
        - User wants another sheet with the same layout and data.

    WHEN NOT TO USE:
        - User wants to move the original worksheet (use move_worksheet).
        - User wants to copy only a cell range.

    PARAMETERS:
        source_sheet_name: Name of the worksheet to copy (case-insensitive).
        new_sheet_name: Name for the copied worksheet.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Source sheet, copied sheet and the updated list of sheet names.

    NOTES:
        - The copy is placed right after the original.
        - The original worksheet remains unchanged.
    """
    book = get_workbook(workbook_name)
    source = _find(book, source_sheet_name)
    new = _validate_name(new_sheet_name)
    _ensure_unique(book, new)

    src_sheet = book.sheets[source]
    copied = src_sheet.copy(after=src_sheet)
    copied.name = new
    return {"success": True, "workbook": book.name, "source_sheet": source,
            "copied_sheet": copied.name, "sheets": _names(book)}


# ---------------------------------------------------------
# 5. GET ALL WORKSHEETS
# ---------------------------------------------------------

@safe_tool
def get_all_worksheets(workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        List all worksheets in a workbook with position, visibility and
        used range.

    WHEN TO USE:
        - User asks which sheets exist in the workbook.
        - Before renaming, copying, moving or deleting, to check exact sheet names.

    WHEN NOT TO USE:
        - User wants to read cell data (use a range-reading tool).
        - User wants full details of one sheet (use get_worksheet_info).

    PARAMETERS:
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Workbook name, total sheet count and one entry per sheet:
        name, position (starts at 1), visible, state, used_range.

    NOTES:
        - Read-only. Nothing in the workbook is changed.
    """
    book = get_workbook(workbook_name)
    sheets = []
    for s in book.sheets:
        state = _sheet_state(s)
        sheets.append({
            "name": s.name,
            "position": s.index,
            "visible": state == "visible",
            "state": state,
            "used_range": s.used_range.address,
        })
    return {"success": True, "workbook": book.name,
            "total_sheets": len(sheets), "sheets": sheets}


# ---------------------------------------------------------
# 6. MOVE WORKSHEET
# ---------------------------------------------------------

@safe_tool
def move_worksheet(sheet_name: str,
                   position: int | None = None,
                   before_sheet: str | None = None,
                   after_sheet: str | None = None,
                   workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Change the position of a worksheet inside the workbook.

    WHEN TO USE:
        - User asks to move a sheet to the first/last/Nth position.
        - User wants a sheet placed before or after another sheet.

    WHEN NOT TO USE:
        - User wants a duplicate of the sheet (use copy_worksheet).
        - User wants to move a cell range (use a range tool).

    PARAMETERS:
        sheet_name: Worksheet to move (case-insensitive).
        position: Final 1-based position (1 = first sheet).
        before_sheet: Place the sheet before this sheet.
        after_sheet: Place the sheet after this sheet.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Moved sheet, its new position and the updated list of sheet names.

    NOTES:
        - Provide exactly ONE of position, before_sheet, after_sheet.
        - Sheet data is not changed.
    """
    given = [x for x in (position, before_sheet, after_sheet) if x is not None]
    if len(given) != 1:
        raise ExcelToolError("Provide exactly one of position, before_sheet or after_sheet.")

    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]
    others = [s for s in book.sheets if s.name != target]

    if position is not None:
        _check_int(position, "position", minimum=1)
        if position > len(book.sheets):
            raise ExcelToolError(f"position cannot exceed {len(book.sheets)}.")
        if position <= len(others):
            sheet.api.Move(Before=others[position - 1].api)
        else:
            sheet.api.Move(After=others[-1].api)
    elif before_sheet is not None:
        anchor = _find(book, before_sheet)
        if anchor == target:
            raise ExcelToolError("Cannot move a sheet relative to itself.")
        sheet.api.Move(Before=book.sheets[anchor].api)
    else:
        anchor = _find(book, after_sheet)
        if anchor == target:
            raise ExcelToolError("Cannot move a sheet relative to itself.")
        sheet.api.Move(After=book.sheets[anchor].api)

    return {"success": True, "workbook": book.name, "moved": target,
            "new_position": book.sheets[target].index, "sheets": _names(book)}


# ---------------------------------------------------------
# 7. HIDE / UNHIDE WORKSHEET
# ---------------------------------------------------------

@safe_tool
def hide_unhide_worksheet(sheet_name: str,
                          action: str = "hide",
                          workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Hide, very-hide or show a worksheet.

    WHEN TO USE:
        - User wants to hide a sheet without deleting it.
        - User wants to unhide a hidden sheet.

    WHEN NOT TO USE:
        - User wants to remove the sheet permanently (use delete_worksheet).

    PARAMETERS:
        sheet_name: Worksheet name (case-insensitive).
        action: 'hide', 'very_hidden' or 'show'.
            'very_hidden' sheets cannot be unhidden from Excel's normal menu.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Sheet name, new state and the visible sheet count.

    NOTES:
        - The last visible worksheet cannot be hidden.
        - Data in hidden sheets is preserved.
    """
    act = str(action).strip().lower()
    if act not in ("hide", "very_hidden", "show"):
        raise ExcelToolError("action must be 'hide', 'very_hidden' or 'show'.")

    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]

    if act == "show":
        sheet.api.Visible = -1
    else:
        if _sheet_state(sheet) == "visible" and _visible_count(book) <= 1:
            raise ExcelToolError("Cannot hide the last visible worksheet.")
        sheet.api.Visible = 0 if act == "hide" else 2

    return {"success": True, "workbook": book.name, "sheet": target,
            "state": _sheet_state(sheet), "visible_sheets": _visible_count(book)}


# ---------------------------------------------------------
# 8. SET SHEET TAB COLOR
# ---------------------------------------------------------

@safe_tool
def set_sheet_tab_color(sheet_name: str,
                        color: str | None = None,
                        workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Set or clear the color of a worksheet tab.

    WHEN TO USE:
        - User wants to color-code sheet tabs.
        - User wants to remove a tab color.

    WHEN NOT TO USE:
        - User wants to color cells (use a formatting tool).

    PARAMETERS:
        sheet_name: Worksheet name (case-insensitive).
        color: Hex color like '#FF0000'. Omit or None to clear the color.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Sheet name and the resulting tab color (None if cleared).

    NOTES:
        - Only the tab changes; sheet data is untouched.
    """
    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]

    if color is None or (isinstance(color, str) and not color.strip()):
        sheet.api.Tab.ColorIndex = -4142  # xlColorIndexNone
        return {"success": True, "workbook": book.name, "sheet": target,
                "tab_color": None}

    sheet.api.Tab.Color = _parse_hex_color(color)
    return {"success": True, "workbook": book.name, "sheet": target,
            "tab_color": _bgr_to_hex(sheet.api.Tab.Color)}


# ---------------------------------------------------------
# 9. PROTECT / UNPROTECT WORKSHEET
# ---------------------------------------------------------

@safe_tool
def protect_unprotect_worksheet(sheet_name: str,
                                action: str = "protect",
                                password: str | None = None,
                                workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Protect a worksheet from edits, or remove its protection.

    WHEN TO USE:
        - User wants to lock a sheet so cells cannot be changed.
        - User wants to remove sheet protection.

    WHEN NOT TO USE:
        - User wants to protect the whole workbook structure.

    PARAMETERS:
        sheet_name: Worksheet name (case-insensitive).
        action: 'protect' or 'unprotect'.
        password: Optional password. Needed to unprotect if one was set.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Sheet name and whether it is now protected.

    NOTES:
        - A forgotten password cannot be recovered by this tool.
        - Passwords are never returned in the result.
    """
    act = str(action).strip().lower()
    if act not in ("protect", "unprotect"):
        raise ExcelToolError("action must be 'protect' or 'unprotect'.")

    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]
    already = bool(sheet.api.ProtectContents)

    if act == "protect":
        if already:
            raise ExcelToolError(f"Sheet '{target}' is already protected.")
        if password:
            sheet.api.Protect(Password=str(password))
        else:
            sheet.api.Protect()
    else:
        if not already:
            raise ExcelToolError(f"Sheet '{target}' is not protected.")
        try:
            if password:
                sheet.api.Unprotect(Password=str(password))
            else:
                sheet.api.Unprotect()
        except Exception:
            raise ExcelToolError("Could not unprotect the sheet. The password is missing or wrong.")

    return {"success": True, "workbook": book.name, "sheet": target,
            "protected": bool(sheet.api.ProtectContents)}


# ---------------------------------------------------------
# 10. FREEZE PANES
# ---------------------------------------------------------

@safe_tool
def freeze_panes(sheet_name: str,
                 rows: int = 0,
                 columns: int = 0,
                 unfreeze: bool = False,
                 workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Freeze the top rows and/or left columns of a worksheet, or unfreeze.

    WHEN TO USE:
        - User wants the header row to stay visible while scrolling.
        - User wants the first column(s) locked.
        - User wants to remove frozen panes.

    WHEN NOT TO USE:
        - User wants to hide rows or columns.

    PARAMETERS:
        sheet_name: Worksheet name (case-insensitive).
        rows: Number of top rows to freeze (0 = none).
        columns: Number of left columns to freeze (0 = none).
        unfreeze: True removes any frozen panes (rows/columns ignored).
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Sheet name and the resulting frozen rows/columns.

    NOTES:
        - The sheet must be visible.
        - Any existing freeze is replaced.
    """
    _check_int(rows, "rows")
    _check_int(columns, "columns")
    if not unfreeze and rows == 0 and columns == 0:
        raise ExcelToolError("Give rows and/or columns to freeze, or set unfreeze=True.")

    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]
    _require_visible(sheet)

    with _activated(book, sheet) as win:
        win.FreezePanes = False
        win.SplitColumn = 0
        win.SplitRow = 0
        if unfreeze:
            frozen_rows, frozen_cols = 0, 0
        else:
            win.ScrollRow = 1
            win.ScrollColumn = 1
            win.SplitColumn = columns
            win.SplitRow = rows
            win.FreezePanes = True
            frozen_rows, frozen_cols = rows, columns

    return {"success": True, "workbook": book.name, "sheet": target,
            "frozen_rows": frozen_rows, "frozen_columns": frozen_cols}


# ---------------------------------------------------------
# 11. SET WORKSHEET ZOOM / VIEW
# ---------------------------------------------------------

@safe_tool
def set_worksheet_zoom_view(sheet_name: str,
                            zoom: int | None = None,
                            show_gridlines: bool | None = None,
                            workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Change the zoom level and gridline display of a worksheet.

    WHEN TO USE:
        - User wants to zoom in or out.
        - User wants gridlines shown or hidden.

    WHEN NOT TO USE:
        - User wants to change print layout (use set_print_setup).

    PARAMETERS:
        sheet_name: Worksheet name (case-insensitive).
        zoom: Zoom percent from 10 to 400. Omit to keep the current zoom.
        show_gridlines: True/False. Omit to keep the current setting.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Sheet name, current zoom and gridline state.

    NOTES:
        - The sheet must be visible.
        - At least one of zoom or show_gridlines is required.
    """
    if zoom is None and show_gridlines is None:
        raise ExcelToolError("Provide zoom and/or show_gridlines.")
    if zoom is not None:
        _check_int(zoom, "zoom", minimum=10)
        if zoom > 400:
            raise ExcelToolError("zoom must be between 10 and 400.")

    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]
    _require_visible(sheet)

    with _activated(book, sheet) as win:
        if zoom is not None:
            win.Zoom = zoom
        if show_gridlines is not None:
            win.DisplayGridlines = bool(show_gridlines)
        current_zoom = int(win.Zoom)
        gridlines = bool(win.DisplayGridlines)

    return {"success": True, "workbook": book.name, "sheet": target,
            "zoom": current_zoom, "show_gridlines": gridlines}


# ---------------------------------------------------------
# 12. SET PRINT SETUP
# ---------------------------------------------------------

@safe_tool
def set_print_setup(sheet_name: str,
                    orientation: str | None = None,
                    print_area: str | None = None,
                    fit_to_one_page_wide: bool | None = None,
                    workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Configure how a worksheet prints.

    WHEN TO USE:
        - User wants portrait or landscape printing.
        - User wants to set or clear a print area.
        - User wants the sheet to fit one page wide.

    WHEN NOT TO USE:
        - User wants to change on-screen zoom (use set_worksheet_zoom_view).

    PARAMETERS:
        sheet_name: Worksheet name (case-insensitive).
        orientation: 'portrait' or 'landscape'. Omit to keep the current one.
        print_area: Range like 'A1:F40'. Empty string '' clears the print area.
            Omit to keep the current one.
        fit_to_one_page_wide: True fits all columns on one page width
            (height automatic). False turns fitting off. Omit to keep it.
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Sheet name and the resulting orientation, print area and fit setting.

    NOTES:
        - At least one setting is required.
        - Needs a default printer installed in Windows for page setup to work.
    """
    if orientation is None and print_area is None and fit_to_one_page_wide is None:
        raise ExcelToolError("Provide orientation, print_area and/or fit_to_one_page_wide.")

    orient = None
    if orientation is not None:
        orient = str(orientation).strip().lower()
        if orient not in ("portrait", "landscape"):
            raise ExcelToolError("orientation must be 'portrait' or 'landscape'.")

    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]
    setup = sheet.api.PageSetup

    if orient is not None:
        setup.Orientation = 1 if orient == "portrait" else 2  # xlPortrait / xlLandscape

    if print_area is not None:
        if not isinstance(print_area, str):
            raise ExcelToolError("print_area must be text like 'A1:F40'.")
        area = print_area.strip()
        if area:
            try:
                sheet.range(area)  # validate the address
            except Exception:
                raise ExcelToolError(f"Invalid print_area '{print_area}'.")
        setup.PrintArea = area.replace("$", "") if area else ""

    if fit_to_one_page_wide is not None:
        if fit_to_one_page_wide:
            setup.Zoom = False
            setup.FitToPagesWide = 1
            setup.FitToPagesTall = False
        else:
            setup.Zoom = 100

    return {"success": True, "workbook": book.name, "sheet": target,
            "orientation": "landscape" if int(setup.Orientation) == 2 else "portrait",
            "print_area": setup.PrintArea or None,
            "fit_to_one_page_wide": setup.Zoom is False}


# ---------------------------------------------------------
# 13. GET WORKSHEET INFO
# ---------------------------------------------------------

@safe_tool
def get_worksheet_info(sheet_name: str,
                       workbook_name: str | None = None) -> dict:
    """
    PURPOSE:
        Return detailed information about one worksheet.

    WHEN TO USE:
        - User asks about a sheet's size, protection, color or freeze status.
        - Before editing a sheet, to check if it is protected or hidden.

    WHEN NOT TO USE:
        - User wants only the list of sheets (use get_all_worksheets).
        - User wants cell data (use a range-reading tool).

    PARAMETERS:
        sheet_name: Worksheet name (case-insensitive).
        workbook_name: Target workbook. Uses the active workbook if omitted.

    RETURNS:
        Name, position, state, used range, row/column count, tab color,
        protection, frozen rows/columns, zoom, gridlines, and counts of
        charts and tables.

    NOTES:
        - Read-only. The previously active sheet is restored afterwards.
        - View details (freeze, zoom, gridlines) are None for hidden sheets.
    """
    book = get_workbook(workbook_name)
    target = _find(book, sheet_name)
    sheet = book.sheets[target]
    state = _sheet_state(sheet)

    used = sheet.used_range
    empty = used.count == 1 and used.value is None

    tab_color = None
    try:
        if int(sheet.api.Tab.ColorIndex) != -4142:
            tab_color = _bgr_to_hex(sheet.api.Tab.Color)
    except Exception:
        tab_color = None

    try:
        charts = len(sheet.charts)
    except Exception:
        charts = None
    try:
        tables = int(sheet.api.ListObjects.Count)
    except Exception:
        tables = None

    frozen_rows = frozen_cols = zoom = gridlines = None
    if state == "visible":
        with _activated(book, sheet) as win:
            if bool(win.FreezePanes):
                frozen_rows, frozen_cols = int(win.SplitRow), int(win.SplitColumn)
            else:
                frozen_rows, frozen_cols = 0, 0
            zoom = int(win.Zoom)
            gridlines = bool(win.DisplayGridlines)

    return {"success": True, "workbook": book.name, "sheet": target,
            "position": sheet.index, "state": state,
            "used_range": None if empty else used.address,
            "rows": 0 if empty else used.rows.count,
            "columns": 0 if empty else used.columns.count,
            "tab_color": tab_color,
            "protected": bool(sheet.api.ProtectContents),
            "frozen_rows": frozen_rows, "frozen_columns": frozen_cols,
            "zoom": zoom, "show_gridlines": gridlines,
            "charts": charts, "tables": tables}