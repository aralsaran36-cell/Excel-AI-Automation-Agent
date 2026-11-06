import re
import sys

from excel_tools.workbook_manager import (
    ExcelToolError,
    build_condition_mask,
    contiguous_blocks,
    get_sheet,
    hex_to_rgb,
    load_dataframe,
    safe_tool,
    validate_address,
)

IS_WINDOWS = sys.platform.startswith("win")
CELL_RE = re.compile(r"^\$?[A-Z]{1,3}\$?\d+$")

RULE_OPERATORS = {
    "greater_than": 5,
    "less_than": 6,
    "greater_or_equal": 7,
    "less_or_equal": 8,
    "equal": 3,
    "not_equal": 4,
}
RULE_ALIASES = {"equals": "equal", "not_equals": "not_equal", "duplicate": "duplicates"}
OTHER_RULES = [
    "between", "top_n", "bottom_n", "data_bars", "color_scale", "custom_formula",
    "contains_text", "begins_with", "ends_with", "duplicates", "unique",
    "today", "this_week", "past_due",
]
CF_TYPES = {1: "cell_value", 2: "formula", 3: "color_scale", 4: "data_bar", 5: "top_bottom",
            6: "icon_set", 8: "unique_duplicate", 9: "text", 10: "blanks", 11: "no_blanks",
            12: "errors", 13: "no_errors", 15: "date"}


# ---------- helpers ----------

def _rgb_to_bgr_int(rgb) -> int:
    r, g, b = rgb
    return int(r) + int(g) * 256 + int(b) * 65536


def _first_cell(address: str) -> str:
    first = address.split(":")[0].replace("$", "").upper()
    if not CELL_RE.match(first):
        raise ExcelToolError("Use an explicit cell range like A2:A100 for this rule.")
    return first


def _absolute(address: str) -> str:
    parts = address.replace("$", "").upper().split(":")
    for p in parts:
        if not CELL_RE.match(p):
            raise ExcelToolError("Use an explicit cell range like A2:A100 for this rule.")
    return ":".join(re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", p) for p in parts)


def _q(text) -> str:
    """Escape double quotes for use inside an Excel formula string."""
    return str(text).replace('"', '""')


def _need_text(value, rule: str) -> str:
    if value is None or str(value) == "":
        raise ExcelToolError(f"value (text) is required for rule '{rule}'.")
    return str(value)


def _need_number(value, label: str = "value") -> float:
    if value is None:
        raise ExcelToolError(f"{label} is required for this rule.")
    try:
        return float(value)
    except (TypeError, ValueError):
        raise ExcelToolError(f"{label} must be a number.")


def _num_text(number: float) -> str:
    return str(int(number)) if float(number).is_integer() else str(number)


def _build_rule(rule: str, value, value2, address: str) -> dict:
    """Validate inputs and describe the rule. Nothing is written to Excel here."""
    if rule in RULE_OPERATORS:
        if value is None:
            raise ExcelToolError("value is required for this rule.")
        if rule in {"equal", "not_equal"}:
            try:
                f1 = "=" + _num_text(float(value))
            except (TypeError, ValueError):
                f1 = f'="{_q(value)}"'
        else:
            f1 = "=" + _num_text(_need_number(value))
        return {"kind": "cell", "operator": RULE_OPERATORS[rule], "f1": f1, "f2": None}

    if rule == "between":
        low, high = _need_number(value, "value"), _need_number(value2, "value2")
        if low > high:
            raise ExcelToolError("value must be less than or equal to value2.")
        return {"kind": "cell", "operator": 1, "f1": "=" + _num_text(low),
                "f2": "=" + _num_text(high)}

    if rule in {"top_n", "bottom_n"}:
        n = int(_need_number(value, "value (N)"))
        if not 1 <= n <= 1000:
            raise ExcelToolError("value (N) must be between 1 and 1000.")
        return {"kind": "top", "bottom": rule == "bottom_n", "n": n}

    if rule == "data_bars":
        return {"kind": "databar"}
    if rule == "color_scale":
        return {"kind": "scale"}

    if rule == "custom_formula":
        formula = _need_text(value, rule).strip()
        if not formula.startswith("="):
            formula = "=" + formula
        return {"kind": "expr", "formula": formula}

    a1 = _first_cell(address) if rule not in {"duplicates", "unique"} else None

    if rule == "contains_text":
        return {"kind": "expr", "formula": f'=ISNUMBER(SEARCH("{_q(_need_text(value, rule))}",{a1}))'}
    if rule == "begins_with":
        t = _need_text(value, rule)
        return {"kind": "expr", "formula": f'=LEFT({a1},{len(t)})="{_q(t)}"'}
    if rule == "ends_with":
        t = _need_text(value, rule)
        return {"kind": "expr", "formula": f'=RIGHT({a1},{len(t)})="{_q(t)}"'}
    if rule in {"duplicates", "unique"}:
        absolute = _absolute(address)
        a1 = _first_cell(address)
        op = ">1" if rule == "duplicates" else "=1"
        return {"kind": "expr", "formula": f"=COUNTIF({absolute},{a1}){op}"}
    if rule == "today":
        return {"kind": "expr", "formula": f"={a1}=TODAY()"}
    if rule == "past_due":
        return {"kind": "expr", "formula": f'=AND({a1}<>"",{a1}<TODAY())'}
    if rule == "this_week":
        return {"kind": "expr", "formula":
                f"=AND({a1}>=TODAY()-WEEKDAY(TODAY(),2)+1,{a1}<=TODAY()-WEEKDAY(TODAY(),2)+7)"}

    raise ExcelToolError("Invalid rule. Use: " + ", ".join(list(RULE_OPERATORS) + OTHER_RULES))


def _prepare_active_cell(sheet, address: str) -> None:
    """Relative references in rule formulas are read from the active cell, so select the top-left cell."""
    try:
        sheet.activate()
        sheet.range(address.split(":")[0].replace("$", "")).select()
    except Exception:
        pass


def _set_fill(cond, rgb) -> None:
    if IS_WINDOWS:
        cond.Interior.Color = _rgb_to_bgr_int(rgb)
    else:
        cond.interior.color.set(rgb)


def _apply_spec(rng, spec: dict, rgb) -> None:
    kind = spec["kind"]
    if IS_WINDOWS:
        fcs = rng.api.FormatConditions
        if kind == "cell":
            kw = {"Type": 1, "Operator": spec["operator"], "Formula1": spec["f1"]}
            if spec["f2"] is not None:
                kw["Formula2"] = spec["f2"]
            _set_fill(fcs.Add(**kw), rgb)
        elif kind == "expr":
            _set_fill(fcs.Add(Type=2, Formula1=spec["formula"]), rgb)
        elif kind == "top":
            cond = fcs.AddTop10()
            cond.TopBottom = 2 if spec["bottom"] else 1
            cond.Rank = spec["n"]
            _set_fill(cond, rgb)
        elif kind == "databar":
            bar = fcs.AddDatabar()
            bar.BarColor.Color = _rgb_to_bgr_int(rgb)
        elif kind == "scale":
            fcs.AddColorScale(ColorScaleType=3)
        return

    # Mac (appscript)
    fcs = rng.api.format_conditions
    if kind == "cell":
        kw = {"type": 1, "operator": spec["operator"], "formula1": spec["f1"]}
        if spec["f2"] is not None:
            kw["formula2"] = spec["f2"]
        _set_fill(fcs.add(**kw), rgb)
    elif kind == "expr":
        _set_fill(fcs.add(type=2, formula1=spec["formula"]), rgb)
    elif kind == "top":
        cond = fcs.add_top10()
        cond.top_bottom.set(2 if spec["bottom"] else 1)
        cond.rank.set(spec["n"])
        _set_fill(cond, rgb)
    elif kind == "databar":
        bar = fcs.add_databar()
        bar.bar_color.color.set(rgb)
    elif kind == "scale":
        fcs.add_color_scale(color_scale_type=3)


# ---------- colour names ----------

COLOR_NAMES = {
    "red": "#FF0000", "light red": "#FF9999", "dark red": "#8B0000",
    "crimson": "#DC143C", "maroon": "#800000", "brick red": "#B22222",
    "salmon": "#FA8072", "coral": "#FF7F50", "tomato": "#FF6347",
    "pink": "#FFC0CB", "light pink": "#FFB6C1", "hot pink": "#FF69B4",
    "deep pink": "#FF1493", "rose": "#FF007F", "magenta": "#FF00FF",
    "burgundy": "#800020", "wine": "#722F37",
    "purple": "#7030A0", "light purple": "#CBA3E3", "dark purple": "#301934",
    "violet": "#EE82EE", "lavender": "#E6E6FA", "plum": "#DDA0DD",
    "orchid": "#DA70D6", "mauve": "#E0B0FF", "indigo": "#4B0082",
    "lilac": "#C8A2C8",
    "orange": "#FFA500", "light orange": "#FFD580", "dark orange": "#FF8C00",
    "amber": "#FFBF00", "peach": "#FFDAB9", "apricot": "#FBCEB1",
    "yellow": "#FFFF00", "light yellow": "#FFFFE0", "dark yellow": "#CCCC00",
    "gold": "#FFD700", "lemon": "#FFF44F", "mustard": "#FFDB58",
    "cream": "#FFFDD0", "khaki": "#F0E68C", "beige": "#F5F5DC",
    "green": "#00B050", "light green": "#C6EFCE", "dark green": "#006400",
    "lime": "#00FF00", "lime green": "#32CD32", "olive": "#808000",
    "mint": "#98FF98", "sea green": "#2E8B57", "forest green": "#228B22",
    "emerald": "#50C878", "pale green": "#98FB98", "teal": "#008080",
    "chartreuse": "#7FFF00", "sage": "#9CAF88",
    "blue": "#0070C0", "light blue": "#ADD8E6", "dark blue": "#00008B",
    "sky blue": "#87CEEB", "navy": "#000080", "navy blue": "#000080",
    "royal blue": "#4169E1", "pure blue": "#0000FF", "cyan": "#00FFFF",
    "aqua": "#00FFFF", "light cyan": "#E0FFFF", "turquoise": "#40E0D0",
    "steel blue": "#4682B4", "cornflower blue": "#6495ED",
    "baby blue": "#89CFF0", "powder blue": "#B0E0E6", "azure": "#007FFF",
    "dodger blue": "#1E90FF", "midnight blue": "#191970", "cobalt": "#0047AB",
    "denim": "#1560BD", "periwinkle": "#CCCCFF", "slate blue": "#6A5ACD",
    "brown": "#A52A2A", "light brown": "#C4A484", "dark brown": "#654321",
    "chocolate": "#D2691E", "tan": "#D2B48C", "sienna": "#A0522D",
    "copper": "#B87333", "bronze": "#CD7F32", "sand": "#C2B280",
    "coffee": "#6F4E37",
    "white": "#FFFFFF", "black": "#000000", "grey": "#D9D9D9",
    "gray": "#D9D9D9", "light grey": "#EEEEEE", "light gray": "#EEEEEE",
    "dark grey": "#595959", "dark gray": "#595959", "silver": "#C0C0C0",
    "charcoal": "#36454F", "slate grey": "#708090", "ivory": "#FFFFF0",
    "off white": "#FAF9F6",
}


def _resolve_color(color) -> str:
    """Colour name (light blue, dark red...) -> hex. Hex passes through."""
    text = str(color or "").strip()
    if not text or text.startswith("#"):
        return text or "#FFFF00"
    key = " ".join(text.lower().replace("_", " ").replace("-", " ").split())
    return COLOR_NAMES.get(key, text)


# ---------- tools ----------

@safe_tool
def highlight_excel_values(column: str, condition: str, value: str | float | None = None,
                           color: str = "#FFFF00", workbook_name: str | None = None,
                           sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Highlight cells in a column that match a condition (static color, not a live rule).

    PARAMETERS:
        column: Column header name.
        condition: equals, not_equals, greater_than, less_than, greater_or_equal, less_or_equal, contains, is_empty, duplicate, unique.
        value: Value to compare (not needed for is_empty, duplicate, unique).
        color: Hex color, example #FFFF00 or a color name.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, column, condition, highlighted_cells, color.
    """
    rgb = hex_to_rgb(_resolve_color(color))
    sheet = get_sheet(workbook_name, sheet_name)
    df, used = load_dataframe(sheet)
    if column not in df.columns:
        raise ExcelToolError(f"Column '{column}' not found. Available: {list(df.columns)}")

    condition = str(condition).lower().strip()
    if condition in {"duplicate", "duplicates"}:
        mask = df[column].duplicated(keep=False)
    elif condition in {"unique", "unique_values"}:
        mask = ~df[column].duplicated(keep=False)
    else:
        mask = build_condition_mask(df[column], condition, value)

    matched = [used.row + 1 + int(i) for i in mask[mask].index]
    if not matched:
        return {"success": True, "column": column, "condition": condition,
                "highlighted_cells": 0, "message": "No matching cells found."}

    col_number = used.column + list(df.columns).index(column)
    for start, end in contiguous_blocks(matched):
        sheet.range((start, col_number), (end, col_number)).color = rgb

    return {"success": True, "column": column, "condition": condition,
            "highlighted_cells": len(matched), "color": color}


@safe_tool
def apply_conditional_formatting(cell_range: str, rule: str, value: float | str | None = None,
                                 value2: float | None = None, color: str = "#FFC7CE",
                                 workbook_name: str | None = None,
                                 sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Add a live conditional formatting rule: compare, between, top/bottom N, text, duplicates, dates, formula.

    PARAMETERS:
        cell_range: Range, example B2:B100.
        rule: greater_than, less_than, equal, between, top_n, contains_text, duplicates, custom_formula, today, past_due, data_bars, color_scale.
        value: Number, text, N or formula (custom_formula example =$C2>100).
        value2: Upper limit, only for between.
        color: Hex color or color name, example #FFC7CE.
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, rule.
    """
    address = validate_address(cell_range)
    rule = RULE_ALIASES.get(str(rule).lower().strip(), str(rule).lower().strip())
    spec = _build_rule(rule, value, value2, address)     # validate first
    rgb = hex_to_rgb(_resolve_color(color))
    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)
    if spec["kind"] == "expr":
        _prepare_active_cell(sheet, address)
    _apply_spec(rng, spec, rgb)
    return {"success": True, "range": address, "rule": rule}


@safe_tool
def clear_conditional_formatting(cell_range: str | None = None, workbook_name: str | None = None,
                                 sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Remove conditional formatting rules from a range, or from the whole sheet if no range is given.

    PARAMETERS:
        cell_range: Range, example B2:B100 (optional, default whole sheet).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, removed_rules, message.
    """
    sheet = get_sheet(workbook_name, sheet_name)
    if cell_range:
        address = validate_address(cell_range)
        rng = sheet.range(address)
    else:
        rng = sheet.used_range
        address = rng.address.replace("$", "")

    removed = None
    if IS_WINDOWS:
        try:
            removed = int(rng.api.FormatConditions.Count)
        except Exception:
            removed = None
        rng.api.FormatConditions.Delete()
    else:
        rng.api.format_conditions.delete()

    return {"success": True, "range": address, "removed_rules": removed,
            "message": f"Conditional formatting cleared from {address}."}


@safe_tool
def list_conditional_formatting(cell_range: str | None = None, workbook_name: str | None = None,
                                sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        List conditional formatting rules on a range or on the whole sheet (read-only).

    PARAMETERS:
        cell_range: Range, example B2:B100 (optional, default whole sheet).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, count, rules (index, type, applies_to, formula).
    """
    if not IS_WINDOWS:
        raise ExcelToolError("Listing conditional formatting rules is supported only with Windows Excel (COM).")

    sheet = get_sheet(workbook_name, sheet_name)
    if cell_range:
        address = validate_address(cell_range)
        rng = sheet.range(address)
    else:
        rng = sheet.used_range
        address = rng.address.replace("$", "")

    fcs = rng.api.FormatConditions
    count = int(fcs.Count)
    rules = []
    for i in range(1, min(count, 50) + 1):
        fc = fcs.Item(i)
        info = {"index": i}
        try:
            info["type"] = CF_TYPES.get(int(fc.Type), str(fc.Type))
        except Exception:
            info["type"] = None
        try:
            info["applies_to"] = fc.AppliesTo.Address.replace("$", "")
        except Exception:
            info["applies_to"] = None
        try:
            info["formula"] = fc.Formula1
        except Exception:
            info["formula"] = None
        rules.append(info)

    return {"success": True, "range": address, "count": count, "rules": rules}


@safe_tool
def apply_data_validation(cell_range: str, validation_type: str, items: list[str] | None = None,
                          minimum: float | None = None, maximum: float | None = None,
                          input_message: str = "", error_message: str = "",
                          workbook_name: str | None = None, sheet_name: str | None = None) -> dict:
    """
    PURPOSE:
        Apply data validation (dropdown list, number limits, text length) or clear it.

    PARAMETERS:
        cell_range: Range, example D2:D100.
        validation_type: list, whole_number, decimal, text_length or clear.
        items: Dropdown items for list (items cannot contain commas).
        minimum: Minimum value (not for list).
        maximum: Maximum value (not for list).
        input_message: Hint shown when the cell is selected (optional).
        error_message: Message shown on invalid input (optional).
        workbook_name: Workbook name (optional).
        sheet_name: Sheet name (optional).

    RETURNS:
        success, range, validation.
    """
    address = validate_address(cell_range)
    kind = str(validation_type).lower().strip()
    codes = {"whole_number": 1, "decimal": 2, "list": 3, "text_length": 6}
    if kind != "clear" and kind not in codes:
        raise ExcelToolError("validation_type must be: list, whole_number, decimal, text_length or clear.")

    formula1 = formula2 = None
    if kind == "list":
        if not items or any("," in str(i) for i in items):
            raise ExcelToolError("items is required for a list, and items cannot contain commas.")
        formula1 = ",".join(str(i) for i in items)
    elif kind != "clear":
        if minimum is None or maximum is None or minimum > maximum:
            raise ExcelToolError("minimum and maximum are required (minimum <= maximum).")
        formula1, formula2 = str(minimum), str(maximum)

    sheet = get_sheet(workbook_name, sheet_name)
    rng = sheet.range(address)

    if IS_WINDOWS:
        val = rng.api.Validation
        val.Delete()
        if kind == "clear":
            return {"success": True, "range": address, "validation": "cleared"}
        kw = {"Type": codes[kind], "AlertStyle": 1, "Operator": 1, "Formula1": formula1}
        if formula2 is not None:
            kw["Formula2"] = formula2
        val.Add(**kw)
        if input_message:
            val.InputMessage = input_message[:255]
        if error_message:
            val.ErrorMessage = error_message[:255]
    else:
        rng.api.validation.delete()
        if kind == "clear":
            return {"success": True, "range": address, "validation": "cleared"}
        kw = {"validation_type": codes[kind], "alert_style": 1, "operator": 1, "formula1": formula1}
        if formula2 is not None:
            kw["formula2"] = formula2
        rng.api.validation.add(**kw)
        if input_message:
            rng.api.validation.input_message.set(input_message[:255])
        if error_message:
            rng.api.validation.error_message.set(error_message[:255])

    return {"success": True, "range": address, "validation": kind}