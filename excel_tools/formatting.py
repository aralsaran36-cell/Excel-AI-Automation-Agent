from excel_tools.workbook_manager import (
    ExcelToolError, get_sheet, hex_to_rgb, safe_tool, validate_address,
)


@safe_tool
def format_excel_cells(cell_range: str, bold: bool | None = None, italic: bool | None = None,
                       font_size: int | None = None, font_color: str | None = None,
                       background_color: str | None = None, number_format: str | None = None,
                       workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """Format a range. Colors are hex ('#FF0000'). number_format example: '#,##0.00' or '0%'."""
    address = validate_address(cell_range)
    options = {"bold": bold, "italic": italic, "font_size": font_size, "font_color": font_color,
               "background_color": background_color, "number_format": number_format}
    applied = [name for name, value in options.items() if value is not None]
    if not applied:
        raise ExcelToolError("Kuraintha oru formatting option kudunga.")
    if font_size is not None and not 1 <= font_size <= 409:
        raise ExcelToolError("font_size 1 mudhal 409 varaikkum.")
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)
    if bold is not None:
        rng.font.bold = bold
    if italic is not None:
        rng.font.italic = italic
    if font_size is not None:
        rng.font.size = font_size
    if font_color is not None:
        rng.font.color = hex_to_rgb(font_color)
    if background_color is not None:
        rng.color = hex_to_rgb(background_color)
    if number_format is not None:
        rng.number_format = number_format
    return {"success": True, "range": address, "applied": applied}


@safe_tool
def autofit_excel_columns(cell_range: str | None = None, workbook_name: str | None = None,
                          sheet_name: str | None = None) -> dict:
    """Autofit column widths for a range (or the whole used range if none given)."""
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(validate_address(cell_range)) if cell_range else sheet.used_range
    rng.columns.autofit()
    return {"success": True, "sheet": sheet.name, "autofitted": rng.address}