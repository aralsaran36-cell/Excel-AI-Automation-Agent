"""Tool router: shrinks the tool list with an elimination pipeline.

Pure Python, no LLM, no Excel calls. Takes ~0 ms per request.

Pipeline (a stage removes a tool only when it is SURE):
  0  keyword shortlist (top INITIAL_SHORTLIST tools)
  1  mode            (chat mode keeps read-only tools only)
  2  intent          (read request drops write tools, and the reverse)
  3  object type     (sheet / row / column / range / chart / pivot ...)
  4  sheet clue      (no sheet clue drops sheet-only tools)
  5  range clue      (no range clue drops range-only tools)
  6  required params (e.g. tab color needs a color word)
  7  workbook state  (e.g. last remaining sheet cannot be deleted)
  8  negative words  ("WHEN NOT TO USE" phrases)
  9  score gap       (weak tools next to a clear winner are dropped)
  10 final rank      (top MAX_FINAL tools, tie allows MAX_FINAL_TIE)

Safety rules:
  - Heuristic stages never leave fewer than MIN_KEEP tools (stage is skipped).
  - Hard stages (1, 6, 7, 8) are deterministic facts and may remove anything.
  - Removed tools are returned so the agent can retry with them.
  - Empty result means "unclear": the agent falls back to category shortlist.
"""
import logging
import re
from typing import Any

log = logging.getLogger("excel_agent.router")

INITIAL_SHORTLIST = 15
MAX_FINAL = 5
MAX_FINAL_TIE = 6
MIN_KEEP = 3
STRONG_WEIGHT = 3
VERB_WEIGHT = 1
OBJECT_BONUS = 2
MAX_HITS = 2

SHEET_ONLY = frozenset({"sheet"})
RANGE_ONLY = frozenset({"range"})

# Clues for required parameters (matched on the lower-cased request).
COLOR_CLUE = re.compile(
    r"\b(?:red|green|blue|yellow|orange|purple|pink|black|white|gray|grey|"
    r"brown|teal|cyan|magenta|clear|reset|remove|none)\b|#[0-9a-f]{6}"
)
ZOOM_CLUE = re.compile(r"\d|gridline")
MOVE_CLUE = re.compile(
    r"\b(?:first|last|top|end|beginning|start|position|before|after)\b"
    r"|\b\d+(?:st|nd|rd|th)?\b"
)

# A1 or A1:B5 style references.
RANGE_RE = re.compile(r"(?<!\w)[a-z]{1,3}\d{1,7}(?::[a-z]{1,3}\d{1,7})?(?!\w)")

READ_WORDS = (
    "list", "display", "view", "details", "info", "information", "what",
    "which", "how many", "tell me", "show me", "give me", "read", "preview",
)
WRITE_WORDS = (
    "copy", "rename", "delete", "remove", "move", "hide", "unhide", "freeze",
    "unfreeze", "protect", "unprotect", "create", "insert", "append", "sort",
    "merge", "unmerge", "format", "highlight", "apply", "set", "clear",
    "fill", "write", "change", "update", "modify", "make", "duplicate",
    "reorder",
)


def _t(kind: str, objects: set[str], strong: tuple = (), verbs: tuple = (),
       negative: tuple = (), needs: re.Pattern | None = None) -> dict[str, Any]:
    """Tool info. strong = specific phrases, verbs = weak words."""
    return {
        "kind": kind,                      # "read" or "write"
        "objects": frozenset(objects),     # what the tool works on ("any" = generic)
        "strong": tuple(strong),
        "verbs": tuple(verbs),
        "negative": tuple(negative),
        "needs": needs,
    }


# Keys must match the real tool names registered in mcp_layer/server.py.
TOOL_INFO: dict[str, dict[str, Any]] = {
    # ================= WORKSHEET =================
    "create_worksheet": _t(
        "write", {"sheet"},
        ("create sheet", "new sheet", "add sheet", "insert sheet",
         "create worksheet", "new worksheet", "add worksheet",
         "add a new sheet", "create a new sheet", "make a new sheet",
         "add a tab", "new tab", "create tab", "create blank sheet",
         "create an empty sheet", "add another sheet", "add new worksheet",
         "sheet podu"),
        ("create", "add"),
        negative=("new workbook",),
    ),
    "rename_worksheet": _t(
        "write", {"sheet"},
        ("rename sheet", "rename worksheet", "rename tab", "rename the sheet",
         "rename the worksheet", "rename the tab", "rename this sheet",
         "change sheet name", "change tab name", "change the name of sheet",
         "update sheet name", "set sheet name", "name the sheet",
         "relabel sheet", "call the sheet", "sheet name to"),
        ("rename", "change name"),
        negative=("rename column", "rename header"),
    ),
    "delete_worksheet": _t(
        "write", {"sheet"},
        ("delete sheet", "remove sheet", "delete worksheet",
         "remove worksheet", "drop sheet", "drop worksheet",
         "delete the sheet", "remove the sheet", "delete the worksheet",
         "remove the worksheet", "delete tab", "remove tab",
         "delete this sheet", "delete the tab", "get rid of sheet",
         "erase sheet", "delete entire sheet", "remove that sheet"),
        ("delete", "remove"),
        negative=("delete row", "delete column", "clear contents",
                  "clear data", "remove duplicate"),
    ),
    "copy_worksheet": _t(
        "write", {"sheet"},
        ("copy sheet", "copy worksheet", "duplicate sheet",
         "duplicate worksheet", "clone sheet", "clone worksheet",
         "copy the sheet", "copy the worksheet", "duplicate the sheet",
         "duplicate the worksheet", "make a copy of sheet",
         "make a copy of the sheet", "create a copy of sheet",
         "copy tab", "duplicate tab", "copy this sheet", "backup sheet",
         "sheet copy"),
        ("copy", "clone"),
        negative=("another workbook", "other workbook", "another file",
                  "other file", "copy range", "copy cells"),
    ),
    "get_all_worksheets": _t(
        "read", {"sheet"},
        ("all sheets", "all worksheets", "list worksheets", "sheets in",
         "how many sheets", "sheet count", "number of sheets",
         "worksheet count", "every sheet", "all tabs",
         "worksheets in workbook", "worksheets overview",
         "workbook sheets", "worksheet summary", "sheets with details"),
        ("list",),
    ),
    "get_excel_sheet_names": _t(
        "read", {"sheet"},
        ("sheet names", "list sheets", "which sheets", "what sheets",
         "names of sheets", "sheets available", "available sheets",
         "tab names", "sheet list", "show sheets", "list of sheets",
         "sheets present", "sheet names enna", "name of the sheets"),
        (),
    ),
    "move_worksheet": _t(
        "write", {"sheet"},
        ("move sheet", "move worksheet", "reorder sheet", "first position",
         "last position", "sheet position", "move the sheet", "move tab",
         "move the tab", "reorder sheets", "reorder tabs",
         "sheet order", "change sheet order", "rearrange sheets",
         "put sheet first", "put sheet last", "move sheet to the end",
         "move sheet to the front", "move worksheet to position"),
        ("move", "reorder", "position"),
        negative=("move range", "move cells", "move data", "move column",
                  "move row"),
        needs=MOVE_CLUE,
    ),
    "hide_unhide_worksheet": _t(
        "write", {"sheet"},
        ("hide sheet", "hide worksheet", "unhide sheet", "unhide worksheet",
         "unhide tab", "show sheet", "very hidden", "hidden sheet",
         "hide the tab", "hide tab", "hide this sheet",
         "hide the worksheet", "make sheet hidden", "show hidden sheet",
         "reveal sheet", "make sheet visible", "sheet visibility"),
        ("hide",),
        negative=("hide row", "hide column", "unhide row", "unhide column"),
    ),
    "set_sheet_tab_color": _t(
        "write", {"sheet"},
        ("tab color", "tab colour", "color the tab", "colour the tab",
         "sheet tab color", "sheet tab colour", "change tab color",
         "change tab colour", "set tab color", "worksheet tab color",
         "color of the tab", "paint the tab", "tab background",
         "color the sheet tab", "sheet color"),
        ("tab",),
        negative=("cell color", "background color"),
        needs=COLOR_CLUE,
    ),
    "protect_unprotect_worksheet": _t(
        "write", {"sheet"},
        ("protect sheet", "unprotect sheet", "lock sheet", "unlock sheet",
         "password", "protect worksheet", "unprotect worksheet",
         "protect the sheet", "lock the sheet", "unlock the sheet",
         "password protect", "remove protection", "sheet protection",
         "remove password", "set password", "prevent editing",
         "make sheet read only", "enable protection"),
        ("protect", "unprotect", "lock", "unlock"),
        negative=("protect workbook", "workbook structure"),
    ),
    "freeze_panes": _t(
        "write", {"sheet"},
        ("freeze", "unfreeze", "freeze panes", "frozen", "freeze top row",
         "freeze first row", "freeze first column", "freeze header",
         "freeze the header", "freeze rows", "freeze columns",
         "unfreeze panes", "lock top row", "keep header visible",
         "keep headers visible", "pin header", "pin top row",
         "pin first column", "header stays visible"),
        ("pane",),
        negative=("freeze formula", "freeze value"),
    ),
    "set_worksheet_zoom_view": _t(
        "write", {"sheet"},
        ("zoom", "gridline", "zoom level", "zoom to", "set zoom",
         "zoom in", "zoom out", "hide gridlines", "show gridlines",
         "turn off gridlines", "turn on gridlines", "remove gridlines",
         "page layout view", "normal view", "view mode", "magnify",
         "zoom percent", "reduce zoom"),
        ("view",),
        needs=ZOOM_CLUE,
    ),
    "set_print_setup": _t(
        "write", {"sheet"},
        ("print", "landscape", "portrait", "print area", "fit one page",
         "fit to page", "page setup", "print setup", "print layout",
         "paper size", "a4", "page margins", "set margins",
         "print titles", "repeat header rows", "print orientation",
         "fit to width", "print on one page", "print scaling",
         "page orientation", "header and footer"),
        (),
    ),
    "get_worksheet_info": _t(
        "read", {"sheet"},
        ("sheet info", "worksheet info", "sheet details",
         "worksheet details", "details of sheet", "info about sheet",
         "sheet dimensions", "sheet size", "used range",
         "how many rows in sheet", "row count", "number of rows",
         "max row", "max column", "sheet properties", "sheet structure",
         "sheet overview", "dimensions of"),
        ("details", "info"),
    ),

    # ================= READER =================
    "read_excel_data": _t(
        "read", {"any"},
        ("show data", "display data", "view data", "preview",
         "first rows", "top rows", "show rows", "show records",
         "read data", "read sheet", "read the sheet", "read the data",
         "sample rows", "head of", "get data", "fetch data",
         "see the data", "show the sheet", "show the data",
         "first 10 rows"),
        ("read", "show", "display", "view", "list", "get", "fetch", "see"),
    ),
    "read_excel_range": _t(
        "read", {"range"},
        ("read range", "show range", "range values", "values in range",
         "values in cells", "cells from", "read cells", "read cell",
         "get cell value", "value of cell", "what is in cell",
         "contents of", "show cells", "read values from",
         "data in range", "cell value", "cells between", "values of cells"),
        ("range",),
    ),
    "get_excel_column_info": _t(
        "read", {"column", "data"},
        ("column info", "column type", "column types", "data type",
         "data types", "column details", "blank count", "null count",
         "missing count", "how many blanks", "column names",
         "column headers", "list columns", "what columns", "show columns",
         "header names", "columns in", "describe column"),
        ("columns",),
    ),
    "search_excel_value": _t(
        "read", {"data", "range"},
        ("search for", "find value", "find where", "where is", "locate",
         "look for", "which cell", "search value", "search the sheet",
         "find the cell", "find occurrences", "which row has",
         "which rows contain", "find text", "find the row",
         "search column for", "search data for", "locate value",
         "find records with", "find entries"),
        ("search", "find"),
        negative=("replace", "duplicate"),
    ),

    # ================= WRITER =================
    "write_excel_data": _t(
        "write", {"any"},
        ("write data", "write values", "write rows", "put value",
         "enter value", "type value", "write into", "write to cells",
         "populate cells", "fill in the values", "insert data into",
         "write table", "write list", "write the following",
         "enter data", "input data", "paste data", "save data to sheet"),
        ("write", "put", "enter"),
    ),
    "update_excel_cell": _t(
        "write", {"range"},
        ("update cell", "set cell", "change cell", "edit cell",
         "fill cell", "change value", "update value", "set value",
         "modify cell", "set cell value", "change the value of",
         "edit value", "update the value in", "correct the value",
         "replace cell value", "put in cell"),
        ("update", "change", "set"),
    ),
    "clear_excel_range": _t(
        "write", {"range", "data"},
        ("clear range", "clear cells", "clear contents", "clear data",
         "erase", "empty the cells", "wipe", "wipe out",
         "clear the contents", "delete contents", "delete cell contents",
         "blank out", "clear values", "erase data", "erase cells",
         "remove contents", "empty range", "clear cell"),
        ("clear",),
        negative=("clear filter", "clear format", "clear formatting",
                  "clear conditional", "clear rules", "clear highlight"),
    ),
    "fill_excel_series": _t(
        "write", {"range", "column"},
        ("serial number", "serial numbers", "fill series", "number series",
         "date series", "sequence of", "dates for", "autofill",
         "auto fill", "fill numbers", "number the rows", "add serial",
         "index column", "running number", "sequential numbers",
         "increment by", "month series", "generate dates", "list of dates"),
        ("series", "sequence"),
    ),
    "append_excel_data": _t(
        "write", {"any"},
        ("append", "add record", "add entry", "new record",
         "add at the end", "append row", "append data", "add to the end",
         "add below", "add after last row", "add new entry",
         "insert record", "add a new record", "add new row with",
         "bottom of the data", "next empty row", "last row of data",
         "add data at bottom"),
        ("add",),
    ),

    # ================= RANGE =================
    "copy_excel_range": _t(
        "write", {"range"},
        ("copy range", "copy cells", "move range", "move cells",
         "copy data to", "move data to", "copy values to", "paste range",
         "copy and paste", "cut and paste", "duplicate range",
         "copy the range", "paste cells", "copy cells to", "move values",
         "transfer data to", "duplicate the first", "duplicate the rows",
         "duplicate the record", "make copies of rows"),
        ("copy", "move"),
    ),
    "merge_excel_cells": _t(
        "write", {"range"},
        ("merge cells", "merge range", "merge across", "merge title",
         "merge a1", "merge and center", "merge the cells",
         "combine cells", "join cells", "merge header", "merge heading",
         "merge the title", "merge the header", "merge these cells",
         "merge into one cell", "span title across"),
        ("merge",),
    ),
    "unmerge_excel_cells": _t(
        "write", {"range"},
        ("unmerge", "un merge", "split merged", "unmerge cells",
         "unmerge range", "undo merge", "remove merge", "merged cells",
         "separate merged cells", "break merged", "unmerge header",
         "unmerge title", "demerge", "unmerge all", "unmerge the"),
        ("unmerge",),
    ),
    "copy_excel_format": _t(
        "write", {"range"},
        ("copy format", "copy formatting", "format painter",
         "same format as", "apply same format", "copy style",
         "copy the format", "copy the style", "match formatting",
         "paste format", "paste formatting", "format like",
         "make it look like", "use format from", "same style as",
         "clone formatting", "copy cell format"),
        (),
    ),
    "clear_excel_formatting": _t(
        "write", {"range"},
        ("clear formatting", "clear format", "remove formatting",
         "remove format", "reset format", "reset formatting",
         "strip formatting", "remove styles", "clear styles",
         "clear style", "remove all formatting", "default formatting",
         "unformat", "make plain", "clear cell format"),
        (),
        negative=("conditional",),
    ),

    # ================= ROWS / COLUMNS =================
    "insert_excel_rows": _t(
        "write", {"row"},
        ("insert row", "add row", "add a row", "add blank row",
         "insert blank row", "insert empty row", "add empty row",
         "new blank row", "insert row above", "insert row below",
         "insert row at", "insert row before", "insert row after",
         "row in between", "shift rows down", "new row between"),
        ("insert",),
    ),
    "insert_excel_columns": _t(
        "write", {"column"},
        ("insert column", "insert a column", "insert blank column",
         "insert empty column", "add blank column", "add empty column",
         "add a blank column", "new blank column", "new empty column",
         "insert column before", "insert column after",
         "insert column between", "insert column at",
         "column in between", "shift columns right"),
        ("insert",),
    ),
    "delete_excel_rows": _t(
        "write", {"row"},
        ("delete row", "remove row", "delete the row", "remove the row",
         "delete rows where", "delete rows with", "remove rows where",
         "drop row", "erase row", "delete entire row",
         "remove entire row", "get rid of rows", "delete rows from",
         "delete rows containing", "delete records", "remove records"),
        ("delete", "remove"),
        negative=("blank row", "empty row"),
    ),
    "delete_excel_columns": _t(
        "write", {"column"},
        ("delete column", "remove column", "delete the column",
         "remove the column", "drop column", "erase column",
         "delete entire column", "remove entire column",
         "get rid of column", "delete columns from", "delete a column",
         "remove a column", "delete column named", "remove unwanted column"),
        ("delete", "remove"),
    ),
    "hide_unhide_excel_rows_columns": _t(
        "write", {"row", "column"},
        ("hide row", "hide column", "unhide row", "unhide column",
         "show hidden", "show hidden row", "show hidden column",
         "hide the column", "hide the row", "unhide all", "reveal hidden",
         "hidden row", "hidden column", "hide these columns",
         "hide those rows", "hide specific rows"),
        ("hide", "unhide"),
    ),

    # ================= DASHBOARD =================
    "auto_dashboard": _t(
        "write", {"dashboard"},
        ("create dashboard", "build dashboard", "make dashboard",
         "auto dashboard", "generate dashboard", "new dashboard",
         "create a dashboard", "build a dashboard", "make a dashboard",
         "automatic dashboard", "dashboard from", "kpi dashboard",
         "sales dashboard", "executive dashboard", "summary dashboard",
         "dashboard with kpis", "dashboard with charts", "design dashboard"),
        ("kpi",),
    ),
    "list_dashboards": _t(
        "read", {"dashboard"},
        ("list dashboards", "show dashboards", "all dashboards",
         "which dashboards", "what dashboards", "existing dashboards",
         "dashboards available", "dashboards in workbook",
         "dashboard names", "how many dashboards", "saved dashboards",
         "show all dashboards", "dashboards created", "my dashboards",
         "dashboard list"),
        ("list",),
    ),
    "refresh_dashboard": _t(
        "write", {"dashboard"},
        ("refresh dashboard", "update dashboard data", "reload dashboard",
         "recalculate dashboard", "rebuild dashboard", "sync dashboard",
         "dashboard refresh", "refresh the dashboard", "refresh kpis",
         "refresh dashboard charts", "dashboard with latest data",
         "reload the dashboard", "renew dashboard",
         "update dashboard numbers", "dashboard is outdated"),
        ("refresh",),
    ),
    "modify_dashboard": _t(
        "write", {"dashboard"},
        ("modify dashboard", "change dashboard", "edit dashboard",
         "update dashboard", "add to dashboard", "dashboard filter",
         "add chart to dashboard", "add kpi", "remove kpi",
         "change dashboard title", "dashboard title",
         "change dashboard theme", "dashboard theme", "dashboard color",
         "dashboard layout", "add slicer to dashboard",
         "rearrange dashboard", "customize dashboard"),
        ("modify", "edit", "change"),
        negative=("update dashboard data", "refresh dashboard"),
    ),
    "delete_dashboard": _t(
        "write", {"dashboard"},
        ("delete dashboard", "remove dashboard", "drop dashboard",
         "delete the dashboard", "remove the dashboard", "erase dashboard",
         "get rid of dashboard", "delete all dashboards",
         "delete dashboard sheet", "remove dashboard sheet",
         "discard dashboard", "clear dashboard", "delete this dashboard",
         "remove that dashboard", "dashboard delete"),
        ("delete", "remove"),
    ),

    # ================= FORMATTING =================
    "format_excel_cells": _t(
        "write", {"range", "column", "row"},
        ("bold", "italic", "underline", "strikethrough", "font name",
         "font size", "font color", "font colour", "fill color",
         "fill colour", "background color", "background colour",
         "number format", "currency format", "percentage format",
         "date format", "text color", "format as currency",
         "format as percent", "format the cells", "make header bold"),
        ("color", "colour", "format"),
        negative=("copy format", "clear format", "remove format",
                  "conditional format", "same format"),
    ),
    "set_excel_alignment": _t(
        "write", {"range", "column", "row"},
        ("align", "alignment", "center", "centre", "left align",
         "right align", "wrap text", "indent", "middle align", "top align",
         "align center", "align left", "align right", "center the text",
         "text wrap", "vertical alignment", "horizontal alignment",
         "justify", "align text"),
        ("align", "center", "wrap", "indent"),
    ),
    "apply_excel_borders": _t(
        "write", {"range"},
        ("border", "outline", "add border", "apply border", "all borders",
         "outside border", "thick border", "thin border", "box around",
         "draw box", "cell borders", "border style", "remove border",
         "bottom border", "top border", "double border", "table border"),
        (),
    ),
    "autofit_excel_columns": _t(
        "write", {"column", "range"},
        ("autofit", "auto fit", "auto-fit", "fit column", "best fit",
         "auto size", "autosize", "fit to content", "fit columns to content",
         "resize to fit", "fit width to text", "make columns fit",
         "fit the columns", "column auto width", "autofit rows",
         "adjust columns to fit"),
        (),
    ),
    "set_excel_size": _t(
        "write", {"column", "row"},
        ("column width", "row height", "set width", "set height",
         "widen", "narrow", "resize column", "resize row",
         "make column wider", "make row taller", "increase column width",
         "decrease column width", "increase row height",
         "set column size", "set row size", "column size", "row size",
         "make column narrower", "make row shorter", "width of column"),
        ("width", "height", "resize"),
    ),

    # ================= CONDITIONAL FORMATTING =================
    "apply_conditional_formatting": _t(
        "write", {"range", "column"},
        ("conditional formatting", "conditional format", "color scale",
         "data bar", "icon set", "highlight cells greater than",
         "highlight cells less than", "highlight top", "highlight bottom",
         "highlight above average", "highlight below average",
         "highlight negative", "highlight rows where", "highlight if",
         "color cells based on", "color rows based on",
         "rule based formatting", "traffic light", "heat map"),
        ("rule",),
    ),
    "highlight_excel_values": _t(
        "write", {"range", "column"},
        ("highlight values", "mark values", "highlight value",
         "highlight duplicate", "mark duplicate", "show duplicate",
         "find duplicate", "flag duplicate", "highlight blanks",
         "highlight empty", "highlight missing", "highlight errors",
         "highlight cells containing", "highlight cells equal to",
         "highlight cells with", "color cells", "colour cells",
         "highlight in yellow", "highlight in red", "paint cells"),
        ("mark", "highlight"),
    ),
    "apply_data_validation": _t(
        "write", {"range", "column"},
        ("data validation", "dropdown", "drop down", "drop-down",
         "validation list", "allow only", "dropdown list",
         "drop down list", "restrict input", "restrict entries",
         "only allow", "limit input", "limit values", "allowed values",
         "input validation", "validation rule", "whole numbers only",
         "date validation", "prevent invalid", "error alert"),
        ("restrict",),
        negative=("filter dropdown", "filter arrows", "autofilter"),
    ),
    "clear_conditional_formatting": _t(
        "write", {"conditional"},
        ("clear conditional formatting", "remove conditional formatting",
         "remove highlight", "clear highlight", "clear rules",
         "remove rules", "delete conditional formatting", "delete rules",
         "remove color scale", "remove highlights", "clear highlights",
         "clear conditional rules", "reset conditional formatting",
         "remove data bars", "remove icon sets",
         "remove the highlighting", "stop highlighting",
         "delete formatting rules"),
        ("clear", "remove"),
    ),
    "list_conditional_formatting": _t(
        "read", {"conditional"},
        ("list rules", "show rules", "formatting rules",
         "conditional formatting rules", "what rules", "which rules",
         "existing rules", "list conditional formatting",
         "show conditional formatting", "what conditional formatting",
         "rules applied", "active rules", "rules on sheet",
         "conditional rules", "existing conditional formats",
         "show formatting rules", "view rules", "rules in sheet"),
        ("list", "show"),
    ),

    # ================= FORMULAS =================
    "apply_excel_formula": _t(
        "write", {"range", "column", "formula"},
        ("vlookup", "xlookup", "hlookup", "lookup", "sumif", "sumifs",
         "countif", "countifs", "averageif", "index match", "use formula",
         "apply formula", "write formula", "enter formula", "set formula",
         "put formula", "add formula", "if formula", "nested if",
         "concatenate", "text join", "use the formula"),
        ("formula", "calculate", "compute"),
        negative=("show formula", "display formula", "view formula",
                  "read formula", "what formula", "which formula",
                  "get formula", "see the formula", "check formula",
                  "formula used", "what is the formula"),
    ),
    "apply_excel_formula_range": _t(
        "write", {"range", "column", "formula"},
        ("fill formula", "formula range", "formula down",
         "formula for all rows", "formula to column",
         "apply formula to range", "apply formula to column",
         "formula to all rows", "copy formula down", "drag formula down",
         "fill formula down", "formula across rows",
         "formula for every row", "apply formula for all",
         "extend formula", "fill the formula", "formula to rows",
         "same formula for all rows", "formula to each row",
         "autofill formula"),
        (),
    ),
    "add_excel_formula_column": _t(
        "write", {"column"},
        ("add column", "add a column", "new column", "calculated column",
         "formula column", "add total column", "add a total", "add profit",
         "add computed column", "add derived column",
         "new column with formula", "add column that calculates",
         "create a new column", "create column", "add margin column",
         "add percentage column", "add tax column",
         "add discount column", "add bonus column", "column equal to"),
        ("add",),
    ),
    "get_excel_formula": _t(
        "read", {"range", "formula"},
        ("show formula", "get formula", "what formula", "which formula",
         "formula in", "formula used", "display formula", "view formula",
         "read formula", "see the formula", "formula of cell",
         "formula behind", "what is the formula", "formula at",
         "check formula", "formula text", "formulas used",
         "list formulas", "formulas in"),
        (),
    ),
    "convert_excel_formulas_to_values": _t(
        "write", {"range", "column"},
        ("formulas to values", "convert formulas", "paste as values",
         "paste values", "remove formulas", "hardcode", "hard code",
         "convert formula", "convert to values", "formula to value",
         "replace formulas with values", "keep values only",
         "values only", "static values", "make formulas static",
         "paste special values", "remove formula keep value",
         "convert to static", "turn formulas into values",
         "strip formulas"),
        ("convert",),
    ),

    # ================= CLEANING =================
    "remove_excel_duplicates": _t(
        "write", {"row", "data"},
        ("remove duplicate", "delete duplicate", "drop duplicate",
         "dedupe", "deduplicate", "remove repeated", "delete repeated",
         "get rid of duplicate", "eliminate duplicate",
         "remove same entries", "keep unique rows", "keep only unique",
         "unique rows only", "duplicate neekku", "duplicate remove",
         "duplicates neekku", "clean duplicates", "remove dupes",
         "delete dupes", "remove double entries"),
        ("unique", "dedupe"),
        negative=("highlight duplicate", "show duplicate", "find duplicate",
                  "count duplicate", "mark duplicate", "how many duplicate",
                  "list duplicate", "duplicate the first",
                  "duplicate the row", "duplicate first", "duplicate last"),
    ),
    "find_replace_excel_data": _t(
        "write", {"column", "range"},
        ("find and replace", "find replace", "replace all",
         "replace text", "replace value", "replace word", "replace with",
         "change all", "substitute", "swap text", "replace every",
         "replace occurrences", "replace in column", "replace string",
         "replace names", "change text from", "correct spelling of"),
        ("replace",),
        negative=("replace blank", "replace empty", "replace missing",
                  "replace null"),
    ),
    "fill_missing_excel_values": _t(
        "write", {"column", "range"},
        ("fill missing", "fill blank", "fill empty", "fill blanks",
         "fill forward", "fill down", "replace blank", "replace empty",
         "replace missing", "replace null", "fill null", "fill nan",
         "fill gaps", "fill empty cells", "fill blank cells", "impute",
         "missing values with", "blank cells with", "empty cells with",
         "forward fill", "backfill"),
        ("fill",),
    ),
    "clean_excel_text": _t(
        "write", {"column", "range"},
        ("trim", "extra spaces", "clean text", "proper case", "title case",
         "uppercase", "lowercase", "upper case", "lower case",
         "capitalize", "capitalise", "non-printing", "remove spaces",
         "remove leading spaces", "remove trailing spaces",
         "strip spaces", "convert to uppercase", "convert to lowercase",
         "make uppercase", "make lowercase"),
        ("trim", "clean", "capitalize"),
    ),
    "remove_excel_blank_rows": _t(
        "write", {"row", "data"},
        ("remove blank row", "delete blank row", "remove empty row",
         "delete empty row", "drop blank row", "drop empty row",
         "blank rows from", "empty rows from", "get rid of blank",
         "get rid of empty rows", "clean up blank rows",
         "eliminate blank rows", "blank lines", "remove blank lines",
         "remove empty lines", "delete empty lines", "remove gaps",
         "remove blank"),
        ("blank", "empty"),
        negative=("insert blank", "add blank", "insert empty", "add empty"),
    ),

    # ================= FILTER / SORT =================
    "filter_excel_data": _t(
        "write", {"data", "column"},
        ("filter rows", "filter data", "filter by", "filter where",
         "filter for", "filter the data", "filter the sheet",
         "filter column", "filter records", "show only rows",
         "show only where", "only show rows", "only rows where",
         "keep rows where", "show rows where", "show only the rows",
         "get rows where", "extract rows where", "filter out",
         "rows greater than"),
        ("show only", "containing", "greater than", "less than",
         "equal to"),
        negative=("clear filter", "remove filter", "reset filter",
                  "autofilter", "multiple conditions", "both conditions",
                  "two conditions", "delete row", "remove row"),
    ),
    "filter_excel_multi": _t(
        "write", {"column", "range"},
        ("multiple conditions", "multiple filters", "both conditions",
         "filter with and", "filter with or", "two conditions",
         "several filters", "filter by multiple", "and condition",
         "or condition", "filter on two columns",
         "filter by multiple columns", "filter using and",
         "filter using or", "combine filters", "more than one condition",
         "multi filter", "multi column filter"),
        (),
    ),
    "apply_excel_autofilter": _t(
        "write", {"column", "range"},
        ("autofilter", "auto filter", "filter arrows", "filter dropdown",
         "filter buttons", "filter in place", "add filter",
         "turn on filter", "enable filter", "filter icons",
         "filter headers", "enable autofilter", "apply autofilter",
         "add autofilter", "turn filter on", "filter toggle",
         "add filters to headers", "dropdown arrows on headers",
         "filter dropdowns"),
        (),
    ),
    "clear_excel_filter": _t(
        "write", {"data"},
        ("clear filter", "remove filter", "reset filter", "turn off filter",
         "disable filter", "show all rows", "unfilter", "remove autofilter",
         "clear autofilter", "clear all filters", "remove all filters",
         "reset the filter", "filter off", "undo filter", "cancel filter",
         "remove the filter", "clear the filter", "restore all rows",
         "display all rows"),
        (),
    ),
    "sort_excel_data": _t(
        "write", {"data", "column"},
        ("sort", "sort by", "sort data", "sort column", "sort ascending",
         "sort descending", "sort a to z", "sort z to a",
         "sort alphabetically", "sort highest first", "sort lowest first",
         "sort largest first", "sort smallest first",
         "sort from high to low", "sort from low to high", "order by",
         "arrange by", "arrange in order", "sort newest first",
         "sort oldest first", "sort pannu"),
        ("alphabetical", "ascending", "descending", "arrange"),
    ),

    # ================= ANALYSIS =================
    "analyze_excel_column": _t(
        "read", {"column", "data"},
        ("analyze", "analyse", "analysis of", "statistics", "stats",
         "distribution", "summary statistics", "average of", "mean of",
         "median of", "minimum of", "maximum of", "standard deviation",
         "distribution of", "what is the average", "min and max",
         "column summary", "average salary", "describe the column"),
        ("average", "mean", "median", "min", "max"),
    ),
    "calculate_excel_total": _t(
        "read", {"column", "data"},
        ("total of", "total for", "sum of", "sum the", "sum up", "add up",
         "grand total", "overall total", "calculate the total",
         "calculate total", "calculate the sum", "what is the total",
         "what is the sum", "find the total", "find the sum",
         "get the total", "total enna", "total evlo", "sum enna",
         "column total"),
        ("total", "sum", "calculate"),
    ),
    "compare_excel_datasets": _t(
        "read", {"sheet", "data"},
        ("compare", "comparison", "diff", "compare sheets",
         "compare datasets", "compare two sheets", "compare columns",
         "difference between sheets", "differences between",
         "what changed between", "match between sheets", "reconcile",
         "mismatch between", "compare the two", "compare data",
         "compare with", "compare against", "which rows differ"),
        ("versus", "vs"),
    ),
    "identify_excel_outliers": _t(
        "read", {"column", "data"},
        ("outlier", "anomaly", "anomalies", "unusual values",
         "abnormal values", "extreme values", "detect outliers",
         "find outliers", "values far from", "spikes", "irregular values",
         "odd values", "outlying", "iqr", "z score", "z-score",
         "unexpected values", "identify outliers"),
        (),
    ),
    "value_counts_excel_column": _t(
        "read", {"column", "data"},
        ("value counts", "value count", "frequency", "how many times",
         "count of each", "count by", "unique values", "distinct values",
         "count each", "occurrences", "how many of each", "breakdown by",
         "number of times", "top values", "most common", "most frequent",
         "count per", "count duplicates", "how many duplicates",
         "repeated values", "duplicate count"),
        ("count",),
    ),
    "validate_excel_data": _t(
        "read", {"data"},
        ("validate", "validate data", "validate the data",
         "data quality", "check errors", "find errors", "errors in",
         "check for errors", "error check", "invalid values",
         "invalid entries", "bad data", "data issues", "inconsistencies",
         "verify data", "data errors", "problems in data",
         "scan for errors", "check data quality"),
        ("check", "verify"),
    ),

    # ================= CHARTS =================
    "create_excel_chart": _t(
        "write", {"chart"},
        ("create chart", "add chart", "make chart", "insert chart",
         "draw chart", "build chart", "generate chart", "bar chart",
         "pie chart", "line chart", "column chart", "scatter chart",
         "scatter plot", "area chart", "doughnut chart", "histogram",
         "plot", "visualize", "visualise", "graph of", "chart of",
         "chart for"),
        ("chart", "graph", "pie", "bar", "line"),
        negative=("delete chart", "remove chart", "modify chart",
                  "change chart", "edit chart", "list chart",
                  "chart title", "chart data", "update chart",
                  "resize chart", "move chart"),
    ),
    "modify_excel_chart": _t(
        "write", {"chart"},
        ("modify chart", "change chart", "update chart", "edit chart",
         "chart title", "chart color", "chart colour", "chart type",
         "resize chart", "move chart", "axis title", "legend",
         "data labels", "axis min", "axis max", "change chart type",
         "rename chart", "change chart title", "chart style",
         "show legend", "hide legend", "add data labels",
         "chart position", "chart size"),
        ("modify", "edit", "update", "change", "move", "resize"),
        negative=("chart data", "chart range", "chart source"),
    ),
    "list_excel_charts": _t(
        "read", {"chart"},
        ("list charts", "show charts", "all charts", "chart names",
         "how many charts", "which charts", "what charts",
         "charts in sheet", "existing charts", "charts available",
         "charts in workbook", "number of charts", "chart list",
         "show all charts", "charts present", "charts on sheet"),
        ("list", "show"),
    ),
    "delete_excel_chart": _t(
        "write", {"chart"},
        ("delete chart", "remove chart", "drop chart", "delete the chart",
         "remove the chart", "erase chart", "get rid of chart",
         "delete all charts", "remove all charts", "delete this chart",
         "delete graph", "remove graph", "delete the graph",
         "remove that chart", "discard chart", "clear chart",
         "delete chart named", "remove chart from sheet"),
        ("delete", "remove"),
    ),
    "change_excel_chart_data": _t(
        "write", {"chart"},
        ("chart data", "chart source", "chart range", "change chart data",
         "update chart data", "change chart range", "extend chart",
         "chart source data", "change the data of chart",
         "point chart to", "chart should use", "update chart range",
         "extend chart range", "expand chart", "chart with new data",
         "include new rows in chart", "add series to chart",
         "add data to chart", "chart data range"),
        ("change", "update", "extend"),
    ),

    # ================= PIVOT =================
    "create_excel_native_pivot": _t(
        "write", {"pivot", "column"},
        ("native pivot", "real pivot", "pivottable", "pivot table",
         "pivot table in excel", "create pivot table",
         "insert pivot table", "excel pivot table", "actual pivot table",
         "interactive pivot", "pivot table with slicer",
         "pivot table from data", "create a pivottable",
         "new pivot table", "build pivot table", "make pivot table",
         "pivot table on new sheet", "add pivot table",
         "pivot table fields", "drag fields"),
        (),
    ),
    "create_excel_pivot_summary": _t(
        "write", {"pivot"},
        ("pivot summary", "summary pivot", "simple pivot", "static pivot",
         "pivot style", "summary table", "group by", "grouped summary",
         "crosstab", "cross tab", "summarize by", "summarise by",
         "sum by category", "totals by", "summary by", "aggregate by",
         "aggregate data", "group data by", "grouped totals",
         "breakdown table"),
        ("grouped",),
    ),
    "list_excel_pivots": _t(
        "read", {"pivot"},
        ("list pivots", "show pivots", "all pivots", "pivot names",
         "how many pivots", "which pivots", "what pivots",
         "existing pivots", "pivot tables in sheet", "list pivot tables",
         "show pivot tables", "pivots in workbook", "pivot list",
         "number of pivots", "available pivots", "list summaries"),
        ("list",),
    ),
    "delete_excel_pivot": _t(
        "write", {"pivot"},
        ("delete pivot", "remove pivot", "drop pivot", "delete summary",
         "delete the pivot", "remove the pivot", "erase pivot",
         "delete pivot table", "remove pivot table", "get rid of pivot",
         "delete all pivots", "delete pivot summary",
         "remove summary table", "delete the summary",
         "remove the summary", "discard pivot", "delete that pivot",
         "clear pivot"),
        ("delete", "remove"),
    ),
    "modify_excel_pivot": _t(
        "write", {"pivot"},
        ("modify pivot", "change pivot", "update pivot", "edit pivot",
         "refresh pivot", "pivot filter", "change pivot fields",
         "add field to pivot", "remove field from pivot",
         "change pivot values", "change aggregation", "pivot to average",
         "pivot sum to count", "add column field", "add row field",
         "pivot layout", "rearrange pivot", "pivot grand total",
         "show subtotals", "change pivot data source"),
        (),
    ),

    # ================= AUTOMATION =================
    "save_workflow": _t(
        "write", {"workflow"},
        ("save workflow", "save steps", "record workflow",
         "save as workflow", "save this as workflow",
         "remember these steps", "store workflow", "save these steps",
         "create workflow", "new workflow", "save routine", "save macro",
         "record these steps", "save automation", "save this process",
         "save sequence", "make this a workflow", "save the steps as"),
        ("workflow", "automate"),
    ),
    "run_workflow": _t(
        "write", {"workflow"},
        ("run workflow", "run the workflow", "execute workflow",
         "start workflow", "play workflow", "apply workflow",
         "run saved workflow", "run my workflow", "replay workflow",
         "trigger workflow", "run macro", "run the saved steps",
         "launch workflow", "run automation", "run routine", "run steps",
         "use workflow", "run workflow on"),
        ("run",),
    ),
    "list_workflows": _t(
        "read", {"workflow"},
        ("list workflows", "show workflows", "saved workflows",
         "my workflows", "all workflows", "which workflows",
         "what workflows", "existing workflows", "workflows available",
         "workflow names", "how many workflows", "available workflows",
         "workflow list", "show saved steps", "list saved workflows",
         "list macros"),
        ("list",),
    ),
    "delete_workflow": _t(
        "write", {"workflow"},
        ("delete workflow", "remove workflow", "drop workflow",
         "delete the workflow", "remove the workflow", "erase workflow",
         "delete saved workflow", "get rid of workflow",
         "delete all workflows", "delete this workflow",
         "remove saved workflow", "discard workflow", "delete macro",
         "remove macro", "delete automation", "workflow delete",
         "forget workflow"),
        ("delete", "remove"),
    ),
}


# ---------------- helpers ----------------

def _has(text: str, phrase: str) -> bool:
    """Whole-word / whole-phrase match (allows s, es, ed, d suffix)."""
    pattern = r"(?<!\w)" + re.escape(phrase) + r"(?:s|es|ed|d)?(?!\w)"
    return re.search(pattern, text) is not None


def _mentions_sheet(text: str, sheets: list[str]) -> bool:
    for sheet in sheets:
        name = sheet.strip().lower()
        if len(name) < 2:
            continue
        if re.search(r"(?<!\w)" + re.escape(name) + r"(?!\w)", text):
            return True
    return False


def detect_objects(text: str, sheets: list[str]) -> set[str]:
    """Objects the request talks about (sheet, row, column, range ...)."""
    objects: set[str] = set()
    if (any(_has(text, w) for w in ("sheet", "worksheet", "tab"))
            or _mentions_sheet(text, sheets)):
        objects.add("sheet")
    if _has(text, "row"):
        objects.add("row")
    if _has(text, "column"):
        objects.add("column")
    if (RANGE_RE.search(text) or _has(text, "range") or _has(text, "cell")):
        objects.add("range")
    if any(_has(text, w) for w in ("chart", "graph", "plot")):
        objects.add("chart")
    if _has(text, "pivot"):
        objects.add("pivot")
    if _has(text, "formula"):
        objects.add("formula")
    if any(_has(text, w) for w in ("workflow", "macro")):
        objects.add("workflow")
    return objects


def _prune(cands: list[str], drop: set[str], stage: int,
           removed: dict[str, int], hard: bool = False) -> list[str]:
    """Remove tools. Heuristic stages never go below MIN_KEEP tools."""
    drop = {n for n in drop if n in cands}
    if not drop:
        return cands
    keep = [n for n in cands if n not in drop]
    if not hard and len(keep) < min(MIN_KEEP, len(cands)):
        log.debug("router stage %d skipped (would keep %d)", stage, len(keep))
        return cands
    for name in drop:
        removed[name] = stage
    return keep


def sheet_only(names: list[str]) -> bool:
    """True when every known tool in names works only on worksheets."""
    infos = [TOOL_INFO[n] for n in names if n in TOOL_INFO]
    return bool(infos) and all(i["objects"] == SHEET_ONLY for i in infos)


# ---------------- main entry ----------------

def route_tools(user_text: str,
                available: set[str],
                sheets: list[str] | None = None,
                mode: str = "edit",
                readonly: set[str] | None = None) -> dict[str, Any]:
    """Return {"tools": [...], "removed": {tool: stage}, "scores": {...}}.

    "tools" is empty when the request is unclear (use the fallback shortlist).
    """
    text = f" {(user_text or '').lower().strip()} "
    sheets = [str(s) for s in (sheets or [])]
    removed: dict[str, int] = {}

    # ---- Stage 0: keyword shortlist ----
    hits: dict[str, tuple[int, int]] = {}
    for name, info in TOOL_INFO.items():
        if name not in available:
            continue
        strong = sum(1 for p in info["strong"] if _has(text, p))
        verb = sum(1 for p in info["verbs"] if _has(text, p))
        if strong or verb:
            hits[name] = (strong, verb)
    if not hits:
        return {"tools": [], "removed": {}, "scores": {}}

    base = {
        n: STRONG_WEIGHT * min(s, MAX_HITS) + VERB_WEIGHT * min(v, MAX_HITS)
        for n, (s, v) in hits.items()
    }
    cands = sorted(hits, key=lambda n: (-base[n], n))[:INITIAL_SHORTLIST]
    for name in set(hits) - set(cands):
        removed[name] = 0

    # ---- Stage 1: mode (hard) ----
    if mode == "chat" and readonly is not None:
        cands = _prune(cands, {n for n in cands if n not in readonly},
                       1, removed, hard=True)

    # ---- Stage 2: intent ----
    read_intent = any(_has(text, w) for w in READ_WORDS)
    write_intent = any(_has(text, w) for w in WRITE_WORDS)
    if write_intent and not read_intent:
        cands = _prune(cands,
                       {n for n in cands if TOOL_INFO[n]["kind"] == "read"},
                       2, removed)
    elif read_intent and not write_intent:
        cands = _prune(cands,
                       {n for n in cands if TOOL_INFO[n]["kind"] == "write"},
                       2, removed)

    # ---- Stage 3: object type ----
    objects = detect_objects(text, sheets)
    if objects:
        drop = set()
        for n in cands:
            info = TOOL_INFO[n]
            if ("any" in info["objects"] or info["objects"] & objects
                    or hits[n][0] > 0):
                continue
            drop.add(n)
        cands = _prune(cands, drop, 3, removed)

    # ---- Stage 4: sheet clue ----
    if "sheet" not in objects:
        cands = _prune(
            cands,
            {n for n in cands
             if TOOL_INFO[n]["objects"] == SHEET_ONLY and hits[n][0] == 0},
            4, removed)

    # ---- Stage 5: range clue ----
    if "range" not in objects:
        cands = _prune(
            cands,
            {n for n in cands
             if TOOL_INFO[n]["objects"] == RANGE_ONLY and hits[n][0] == 0},
            5, removed)

    # ---- Stage 6: required params (hard) ----
    cands = _prune(
        cands,
        {n for n in cands
         if TOOL_INFO[n]["needs"] is not None
         and not TOOL_INFO[n]["needs"].search(text)},
        6, removed, hard=True)

    # ---- Stage 7: workbook state (hard) ----
    if len(sheets) == 1:
        cands = _prune(cands, {"delete_worksheet"}, 7, removed, hard=True)

    # ---- Stage 8: negative words (hard) ----
    cands = _prune(
        cands,
        {n for n in cands
         if any(_has(text, p) for p in TOOL_INFO[n]["negative"])},
        8, removed, hard=True)

    # ---- Stage 9: score gap ----
    if cands:
        top = max(base[n] for n in cands)
        if top >= STRONG_WEIGHT:
            cands = _prune(
                cands,
                {n for n in cands if hits[n][0] == 0 and base[n] * 3 <= top},
                9, removed)

    # ---- Stage 10: final rank ----
    final_score = {
        n: base[n] + (OBJECT_BONUS if TOOL_INFO[n]["objects"] & objects else 0)
        for n in cands
    }
    ranked = sorted(cands, key=lambda n: (-final_score[n], n))
    final = ranked[:MAX_FINAL]
    if (len(ranked) > MAX_FINAL
            and final_score[ranked[MAX_FINAL]] == final_score[ranked[MAX_FINAL - 1]]):
        final = ranked[:MAX_FINAL_TIE]
    for name in ranked:
        if name not in final:
            removed[name] = 10

    log.info("router: tools=%s removed=%s", final, removed)
    return {"tools": final, "removed": removed,
            "scores": {n: final_score[n] for n in final}}