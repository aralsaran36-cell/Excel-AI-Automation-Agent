import importlib
import json
import re
from pathlib import Path
from typing import Any

from excel_tools.workbook_manager import ExcelToolError, safe_tool

WORKFLOW_DIR = Path(__file__).resolve().parent.parent / "data" / "workflows"
NAME_RE = re.compile(r"^[A-Za-z0-9_-]{1,50}$")
MAX_STEPS = 20
BLOCKED_ARGS = {"confirm", "overwrite"}

# tool name -> module (workflow tools and ask_knowledge_base are excluded: no recursion)
TOOL_MODULES = {
    # reader
    "read_excel_data": "reader", "get_excel_sheet_names": "reader",
    "read_excel_range": "reader", "get_excel_column_info": "reader",
    "search_excel_value": "reader",
    # writer
    "write_excel_data": "writer", "append_excel_data": "writer",
    "update_excel_cell": "writer", "clear_excel_range": "writer",
    "fill_excel_series": "writer",
    # worksheet
    "create_worksheet": "worksheet", "rename_worksheet": "worksheet",
    "delete_worksheet": "worksheet", "copy_worksheet": "worksheet",
    "get_all_worksheets": "worksheet", "move_worksheet": "worksheet",
    "hide_unhide_worksheet": "worksheet", "set_sheet_tab_color": "worksheet",
    "protect_unprotect_worksheet": "worksheet", "freeze_panes": "worksheet",
    "set_worksheet_zoom_view": "worksheet", "set_print_setup": "worksheet",
    "get_worksheet_info": "worksheet",
    # cleaning
    "remove_excel_duplicates": "cleaning", "remove_excel_blank_rows": "cleaning",
    "find_replace_excel_data": "cleaning", "fill_missing_excel_values": "cleaning",
    "clean_excel_text": "cleaning",
    # formulas
    "apply_excel_formula": "formulas", "apply_excel_formula_range": "formulas",
    "add_excel_formula_column": "formulas", "get_excel_formula": "formulas",
    "convert_excel_formulas_to_values": "formulas",
    # formatting
    "format_excel_cells": "formatting", "set_excel_alignment": "formatting",
    "apply_excel_borders": "formatting", "autofit_excel_columns": "formatting",
    "set_excel_size": "formatting",
    # conditional formatting
    "highlight_excel_values": "conditional_formatting",
    "apply_conditional_formatting": "conditional_formatting",
    "apply_data_validation": "conditional_formatting",
    "clear_conditional_formatting": "conditional_formatting",
    "list_conditional_formatting": "conditional_formatting",
    # filter / sort
    "filter_excel_data": "filter_sort", "filter_excel_multi": "filter_sort",
    "apply_excel_autofilter": "filter_sort", "clear_excel_filter": "filter_sort",
    "sort_excel_data": "filter_sort",
    # charts
    "create_excel_chart": "charts", "modify_excel_chart": "charts",
    "list_excel_charts": "charts", "delete_excel_chart": "charts",
    "change_excel_chart_data": "charts",
    # analysis
    "analyze_excel_column": "analysis", "calculate_excel_total": "analysis",
    "identify_excel_outliers": "analysis", "compare_excel_datasets": "analysis",
    "value_counts_excel_column": "analysis",
    # validator
    "validate_excel_data": "validator",
    # pivots
    "create_excel_pivot_summary": "pivot_tables", "create_excel_native_pivot": "pivot_tables",
    "list_excel_pivots": "pivot_tables", "modify_excel_pivot": "pivot_tables",
    "delete_excel_pivot": "pivot_tables",
    # rows / columns
    "insert_excel_rows": "row_column_manager", "insert_excel_columns": "row_column_manager",
    "delete_excel_rows": "row_column_manager", "delete_excel_columns": "row_column_manager",
    "hide_unhide_excel_rows_columns": "row_column_manager",
    # ranges
    "copy_excel_range": "range_manager", "merge_excel_cells": "range_manager",
    "unmerge_excel_cells": "range_manager", "copy_excel_format": "range_manager",
    "clear_excel_formatting": "range_manager",
}


# ---------- helpers ----------

def _norm(name) -> str:
    """Workflow name: trim, spaces -> underscore. Lookup is not case-sensitive."""
    text = re.sub(r"\s+", "_", str(name).strip())
    if not NAME_RE.match(text):
        raise ExcelToolError("Workflow name: letters, numbers, _ and - only (max 50).")
    return text


def _find(name) -> Path | None:
    key = _norm(name).lower()
    if WORKFLOW_DIR.exists():
        for p in WORKFLOW_DIR.glob("*.json"):
            if p.stem.lower() == key:
                return p
    return None


def _resolve(tool: str):
    module = TOOL_MODULES.get(tool)
    if module is None:
        raise ExcelToolError(f"Tool '{tool}' cannot be used in a workflow.")
    try:
        return getattr(importlib.import_module(f"excel_tools.{module}"), tool)
    except (ImportError, AttributeError):
        raise ExcelToolError(f"Tool '{tool}' is not available yet.")


def _validate_steps(steps: Any) -> list[dict]:
    if not isinstance(steps, list) or not 1 <= len(steps) <= MAX_STEPS:
        raise ExcelToolError(f"steps must be a list of 1 to {MAX_STEPS} items.")
    clean = []
    for i, step in enumerate(steps, 1):
        if not isinstance(step, dict) or step.get("tool") not in TOOL_MODULES:
            raise ExcelToolError(f"Step {i}: a valid 'tool' name is required.")
        args = step.get("arguments", {})
        if not isinstance(args, dict):
            raise ExcelToolError(f"Step {i}: 'arguments' must be an object.")
        clean.append({"tool": step["tool"], "arguments": args})
    return clean


def _is_preview(result: dict) -> bool:
    """True when a step only previewed (confirm was stripped), so nothing was changed."""
    if not isinstance(result, dict):
        return False
    if result.get("needs_confirm") is True:
        return True
    if result.get("deleted") is False and result.get("found", 0) > 0:
        return True
    if result.get("updated") is False and result.get("filled", 0) > 0:
        return True
    if result.get("found", 0) > 0 and result.get("replaced") == 0:
        return True
    return False


def _short(result: dict) -> str:
    if not isinstance(result, dict):
        return ""
    text = result.get("error") or result.get("message") or ""
    return str(text)[:200]


# ---------- tools ----------

@safe_tool
def save_workflow(name: str, steps: list[dict], description: str = "",
                  overwrite: bool = False) -> dict:
    """
    PURPOSE:
        Save a workflow (ordered tool steps). Not registered for the LLM: used from code or a future record button.

    PARAMETERS:
        name: Workflow name (letters, numbers, _ and -).
        steps: List like [{"tool": "...", "arguments": {...}}], max 20.
        description: Short description (optional).
        overwrite: true to replace an existing workflow.

    RETURNS:
        success, saved, steps.
    """
    key = _norm(name)
    clean = _validate_steps(steps)
    existing = _find(key)
    if existing is not None and not overwrite:
        raise ExcelToolError(f"Workflow '{key}' already exists. Use overwrite=true to replace it.")
    WORKFLOW_DIR.mkdir(parents=True, exist_ok=True)
    path = existing or (WORKFLOW_DIR / f"{key}.json")
    path.write_text(json.dumps({"name": key, "description": str(description)[:300],
                                "steps": clean}, indent=2, ensure_ascii=False), encoding="utf-8")
    return {"success": True, "saved": key, "steps": len(clean)}


@safe_tool
def list_workflows() -> dict:
    """
    PURPOSE:
        List saved workflows (read-only).

    PARAMETERS:
        none.

    RETURNS:
        success, count, workflows (name, description, steps).
    """
    items = []
    if WORKFLOW_DIR.exists():
        for p in sorted(WORKFLOW_DIR.glob("*.json")):
            try:
                data = json.loads(p.read_text(encoding="utf-8"))
                items.append({"name": p.stem, "description": data.get("description", ""),
                              "steps": len(data.get("steps", []))})
            except (OSError, json.JSONDecodeError):
                items.append({"name": p.stem, "description": "(read error)", "steps": 0})
    return {"success": True, "count": len(items), "workflows": items}


@safe_tool
def run_workflow(name: str, stop_on_error: bool = True) -> dict:
    """
    PURPOSE:
        Run a saved workflow step by step on the active sheet. confirm and overwrite are removed from every step, so delete/sort/replace steps only preview.

    PARAMETERS:
        name: Workflow name (see list_workflows, not case-sensitive).
        stop_on_error: true to stop at the first failed step.

    RETURNS:
        success, workflow, steps_total, steps_run, preview_only_steps, results (step, tool, status, message).
    """
    path = _find(name)
    if path is None:
        raise ExcelToolError(f"Workflow '{name}' not found. Use list_workflows.")
    try:
        steps = _validate_steps(json.loads(path.read_text(encoding="utf-8")).get("steps"))
    except (OSError, json.JSONDecodeError):
        raise ExcelToolError(f"Workflow '{name}' file could not be read.")

    results, ok, previews = [], True, 0
    for i, step in enumerate(steps, 1):
        args = {k: v for k, v in step["arguments"].items() if k not in BLOCKED_ARGS}
        try:
            result = _resolve(step["tool"])(**args)
        except TypeError as exc:
            result = {"success": False, "error": f"Invalid arguments: {exc}"}
        except ExcelToolError as exc:
            result = {"success": False, "error": str(exc)}

        success = bool(result.get("success")) if isinstance(result, dict) else False
        if success and _is_preview(result):
            status = "preview_only"
            previews += 1
        else:
            status = "done" if success else "failed"
        results.append({"step": i, "tool": step["tool"], "status": status,
                        "message": _short(result)})
        if not success:
            ok = False
            if stop_on_error:
                break

    out = {"success": ok, "workflow": path.stem, "steps_total": len(steps),
           "steps_run": len(results), "preview_only_steps": previews, "results": results}
    if previews:
        out["note"] = (f"{previews} step(s) only previewed because confirmation is not allowed inside "
                       "workflows. Run those actions separately to apply them.")
    return out


@safe_tool
def delete_workflow(name: str, confirm: bool = False) -> dict:
    """
    PURPOSE:
        Delete a saved workflow. First call shows a preview, second call with confirm=true deletes.

    PARAMETERS:
        name: Workflow name (see list_workflows).
        confirm: Must be true to actually delete.

    RETURNS:
        success, deleted, name, message.
    """
    path = _find(name)
    if path is None:
        raise ExcelToolError(f"Workflow '{name}' not found. Use list_workflows.")
    if not confirm:
        return {"success": True, "deleted": False, "needs_confirm": True, "name": path.stem,
                "message": f"Workflow '{path.stem}' will be deleted. Please confirm."}
    path.unlink()
    return {"success": True, "deleted": True, "name": path.stem,
            "message": f"Workflow '{path.stem}' deleted."}