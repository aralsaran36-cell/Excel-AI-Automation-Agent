import sys

from excel_tools.workbook_manager import (
    ExcelToolError, get_sheet, hex_to_rgb, safe_tool, validate_address,
)

IS_WINDOWS = sys.platform.startswith("win")

COLOR_NAMES = {
    "red": "#FF0000", "green": "#00B050", "blue": "#0070C0", "yellow": "#FFFF00",
    "orange": "#FFA500", "purple": "#7030A0", "pink": "#FFC0CB", "black": "#000000",
    "white": "#FFFFFF", "gray": "#808080", "grey": "#808080",
}


def _color(value: str):
    """Accept a hex code or a basic color name and return an RGB tuple."""
    text = str(value).strip()
    return hex_to_rgb(COLOR_NAMES.get(text.lower(), text))


def _bgr_int(rgb) -> int:
    r, g, b = rgb
    return int(r) + int(g) * 256 + int(b) * 65536


def _need_com(feature: str) -> None:
    if not IS_WINDOWS:
        raise ExcelToolError(f"{feature} is supported only with Windows Excel (COM).")


def _clean(address: str) -> str:
    return str(address).replace("$", "")


@safe_tool
def format_excel_cells(cell_range: str, bold: bool | None = None, italic: bool | None = None,
                       underline: bool | None = None, strikethrough: bool | None = None,
                       font_name: str | None = None, font_size: int | None = None,
                       font_color: str | None = None, background_color: str | None = None,
                       number_format: str | None = None, workbook_name: str | None = None,
                       sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Format a range: font, size, color, fill color and number format.

    PARAMETERS:
        cell_range: Range, example A1:D1.
        bold: true or false (optional).
        italic: true or false (optional).
        underline: true or false (optional, Windows Excel only).
        strikethrough: true or false (optional, Windows Excel only).
        font_name: Font name, example Arial (optional).
        font_size: Font size 1-409 (optional).
        font_color: Hex like #FF0000 or a color name (optional).
        background_color: Hex like #FFFF00 or a color name (optional).
        number_format: Example #,##0.00 or 0% or dd-mm-yyyy (optional).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, applied.
    """
    address = validate_address(cell_range)
    options = {"bold": bold, "italic": italic, "underline": underline,
               "strikethrough": strikethrough, "font_name": font_name, "font_size": font_size,
               "font_color": font_color, "background_color": background_color,
               "number_format": number_format}
    applied = [name for name, value in options.items() if value is not None]
    if not applied:
        raise ExcelToolError("Give at least one formatting option.")
    if font_size is not None and not 1 <= font_size <= 409:
        raise ExcelToolError("font_size must be between 1 and 409.")
    if underline is not None or strikethrough is not None:
        _need_com("Underline and strikethrough")
    font_rgb = _color(font_color) if font_color is not None else None
    fill_rgb = _color(background_color) if background_color is not None else None

    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)
    if bold is not None:
        rng.font.bold = bold
    if italic is not None:
        rng.font.italic = italic
    if font_name is not None:
        rng.font.name = font_name
    if font_size is not None:
        rng.font.size = font_size
    if font_rgb is not None:
        rng.font.color = font_rgb
    if fill_rgb is not None:
        rng.color = fill_rgb
    if number_format is not None:
        rng.number_format = number_format
    if underline is not None:
        rng.api.Font.Underline = 2 if underline else -4142
    if strikethrough is not None:
        rng.api.Font.Strikethrough = bool(strikethrough)
    return {"success": True, "range": address, "applied": applied}


H_ALIGN_WIN = {"left": -4131, "center": -4108, "right": -4152}
V_ALIGN_WIN = {"top": -4160, "middle": -4108, "bottom": -4107}


@safe_tool
def set_excel_alignment(cell_range: str, horizontal: str | None = None,
                        vertical: str | None = None, wrap_text: bool | None = None,
                        indent: int | None = None, workbook_name: str | None = None,
                        sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Set text alignment, wrap text and indent for a range.

    PARAMETERS:
        cell_range: Range, example A1:D1.
        horizontal: left, center or right (optional).
        vertical: top, middle or bottom (optional).
        wrap_text: true or false (optional).
        indent: Indent level 0-15 (optional).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, applied.
    """
    address = validate_address(cell_range)
    h = str(horizontal).lower().strip() if horizontal is not None else None
    v = str(vertical).lower().strip() if vertical is not None else None
    if h == "centre":
        h = "center"
    if v in {"center", "centre"}:
        v = "middle"
    if h is not None and h not in H_ALIGN_WIN:
        raise ExcelToolError("horizontal must be left, center or right.")
    if v is not None and v not in V_ALIGN_WIN:
        raise ExcelToolError("vertical must be top, middle or bottom.")
    if indent is not None and not 0 <= int(indent) <= 15:
        raise ExcelToolError("indent must be between 0 and 15.")
    applied = [n for n, x in (("horizontal", h), ("vertical", v), ("wrap_text", wrap_text),
                              ("indent", indent)) if x is not None]
    if not applied:
        raise ExcelToolError("Give at least one alignment option.")

    sheet = get_sheet(workbook_name, sheet_name)
    api = sheet.range(address).api
    if IS_WINDOWS:
        if h is not None:
            api.HorizontalAlignment = H_ALIGN_WIN[h]
        if v is not None:
            api.VerticalAlignment = V_ALIGN_WIN[v]
        if wrap_text is not None:
            api.WrapText = bool(wrap_text)
        if indent is not None:
            api.IndentLevel = int(indent)
    else:
        from appscript import k  # Mac (untested: send the traceback if it fails)
        if h is not None:
            api.horizontal_alignment.set(getattr(k, f"horizontal_align_{h}"))
        if v is not None:
            api.vertical_alignment.set(getattr(k, f"vertical_alignment_{v}"))
        if wrap_text is not None:
            api.wrap_text.set(bool(wrap_text))
        if indent is not None:
            api.indent_level.set(int(indent))
    return {"success": True, "range": address, "applied": applied}


BORDER_SIDES = {"left": 7, "top": 8, "bottom": 9, "right": 10}
# Mac (appscript) style codes
BORDER_STYLES = {"thin": 2, "medium": -4138, "thick": 4, "dashed": -4115, "dotted": -4118,
                 "double": -4119, "none": -4142}
# Windows: style -> (LineStyle, Weight)
BORDER_WIN = {"thin": (1, 2), "medium": (1, -4138), "thick": (1, 4), "dashed": (-4115, 2),
              "dotted": (-4118, 2), "double": (-4119, 4), "none": (-4142, 2)}


@safe_tool
def apply_excel_borders(cell_range: str, style: str = "thin", color: str = "#000000",
                        sides: str = "all", workbook_name: str | None = None,
                        sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Apply or remove borders on a range.

    PARAMETERS:
        cell_range: Range, example A1:D10.
        style: thin, medium, thick, dashed, dotted, double or none (none removes borders).
        color: Hex or color name (default black).
        sides: all (outside + inner lines), outline, or a list like left,right,top,bottom.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, style, sides.
    """
    address = validate_address(cell_range)
    style = str(style).lower().strip()
    if style not in BORDER_STYLES:
        raise ExcelToolError(f"style must be one of: {list(BORDER_STYLES)}")
    side_text = str(sides).lower().strip()
    if side_text in {"all", "outline"}:
        chosen = list(BORDER_SIDES)
    else:
        chosen = [s.strip() for s in side_text.split(",")]
        bad = [s for s in chosen if s not in BORDER_SIDES]
        if bad:
            raise ExcelToolError(f"Invalid sides: {bad}. Use all, outline, or left,right,top,bottom.")
    rgb = _color(color)

    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)
    indexes = [BORDER_SIDES[s] for s in chosen]
    if side_text == "all" and rng.count > 1:
        indexes += [11, 12]  # inside vertical, inside horizontal

    for idx in indexes:
        if IS_WINDOWS:
            border = rng.api.Borders(idx)
            line, weight = BORDER_WIN[style]
            border.LineStyle = line
            if style != "none":
                border.Weight = weight
                border.Color = _bgr_int(rgb)
        else:
            border = rng.api.get_border(which_border=idx)
            border.line_style.set(BORDER_STYLES[style])
            if style != "none":
                border.color.set(rgb)
    return {"success": True, "range": address, "style": style, "sides": side_text}


@safe_tool
def autofit_excel_columns(cell_range: str | None = None, autofit_rows: bool = False,
                          workbook_name: str | None = None,
                          sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Autofit column widths (and optionally row heights) for a range or the whole used range.

    PARAMETERS:
        cell_range: Range (optional, default whole used range).
        autofit_rows: true to autofit row heights too (optional).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sheet, autofitted, rows_autofitted.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(validate_address(cell_range)) if cell_range else sheet.used_range
    rng.columns.autofit()
    if autofit_rows:
        rng.rows.autofit()
    return {"success": True, "sheet": sheet.name, "autofitted": _clean(rng.address),
            "rows_autofitted": bool(autofit_rows)}


@safe_tool
def set_excel_size(columns: str | None = None, width: float | None = None,
                   rows: str | None = None, height: float | None = None,
                   workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Set column width and/or row height.

    PARAMETERS:
        columns: Column or span, example C or C:F (use with width).
        width: Column width in characters, 0-255.
        rows: Row or span, example 2 or 2:10 (use with height).
        height: Row height in points, 1-409.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, sheet, columns, width, rows, height.
    """
    if (columns is None) != (width is None):
        raise ExcelToolError("columns and width must be given together.")
    if (rows is None) != (height is None):
        raise ExcelToolError("rows and height must be given together.")
    if columns is None and rows is None:
        raise ExcelToolError("Give columns + width and/or rows + height.")

    col_span = row_span = None
    if columns is not None:
        if not 0 <= float(width) <= 255:
            raise ExcelToolError("width must be between 0 and 255.")
        parts = str(columns).strip().upper().split(":")
        if len(parts) > 2 or not all(p.isalpha() and len(p) <= 3 for p in parts):
            raise ExcelToolError("columns format: 'C' or 'C:F'.")
        col_span = f"{parts[0]}:{parts[-1]}"
    if rows is not None:
        if not 1 <= float(height) <= 409:
            raise ExcelToolError("height must be between 1 and 409.")
        parts = str(rows).strip().split(":")
        if len(parts) > 2 or not all(p.isdigit() and int(p) >= 1 for p in parts):
            raise ExcelToolError("rows format: '2' or '2:10'.")
        nums = sorted(int(p) for p in parts)
        row_span = f"{nums[0]}:{nums[-1]}"

    sheet = get_sheet(workbook_name, sheet_name)
    if col_span:
        sheet.range(col_span).column_width = float(width)
    if row_span:
        sheet.range(row_span).row_height = float(height)
    return {"success": True, "sheet": sheet.name, "columns": col_span,
            "width": float(width) if width is not None else None,
            "rows": row_span, "height": float(height) if height is not None else None}