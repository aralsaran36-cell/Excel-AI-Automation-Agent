"""Offline router test: no LLM, no Excel, no MCP. Run: python test_router.py"""
from backend.router import TOOL_INFO, route_tools

SHEETS = ["Titanic-Dataset", "Sales", "Sales_Copy"]
AVAILABLE = set(TOOL_INFO)

CASES = [
    ("Create sheet Test1", "create_worksheet"),
    ("Rename Test1 to Sales", "rename_worksheet"),
    ("Copy Titanic-Dataset as Sales_Copy", "copy_worksheet"),
    ("List all sheets", "get_all_worksheets"),
    ("Move Sales_Copy to first position", "move_worksheet"),
    ("Hide Sales_Copy", "hide_unhide_worksheet"),
    ("Show Sales_Copy", "hide_unhide_worksheet"),
    ("Make Sales tab red", "set_sheet_tab_color"),
    ("Protect Sales with password 123", "protect_unprotect_worksheet"),
    ("Unprotect Sales with password 123", "protect_unprotect_worksheet"),
    ("Freeze top row in Sales", "freeze_panes"),
    ("Set Sales zoom to 80 and hide gridlines", "set_worksheet_zoom_view"),
    ("Set Sales to landscape, fit one page wide", "set_print_setup"),
    ("Show details of Sales", "get_worksheet_info"),
    ("Delete Sales_Copy", "delete_worksheet"),
]


def main() -> None:
    passed = 0
    for prompt, expected in CASES:
        res = route_tools(prompt, AVAILABLE, SHEETS, "edit", None)
        ok = expected in res["tools"]
        passed += ok
        mark = "PASS" if ok else "FAIL"
        print(f"{mark} | {prompt!r} -> {res['tools']} (expected {expected})")
    print(f"\n{passed}/{len(CASES)} passed")


if __name__ == "__main__":
    main()