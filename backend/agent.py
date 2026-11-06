"""Excel agent: LLM decision loop with native tool calling over MCP.

Connect once (connect), run many requests (run), then close.
Returns dicts only: no print()/input() here (logging is used instead).

Performance / safety:
- Router (backend/router.py) picks 1-6 tools with an elimination pipeline.
  If the router is unsure, the category shortlist is used as fallback.
- The LLM sees SHORT tool schemas (one-line purpose + parameter hints).
- Static system prompt first, workbook context last (better prefix caching).
- Safety net: if the LLM says it has no tool, retry ONCE with a wider list.
- RAG knowledge is retrieved only for categories that benefit from it.
- Pure knowledge questions are answered in ONE call.
- Strict fast path: single call + single step + single-intent request.
- Single tool-call gate (_call): chat-mode block, timeout, error handling.
  Write tools are never auto-retried (duplicate write risk).
- Sheet/column validation errors go back to the model, never silently fixed.
- Stop support through stop_event.
"""

import asyncio
import copy
import json
import logging
import re
import threading
from typing import Any

from backend.knowledge_answer import answer_question
from backend.ollama_engine import OllamaEngine
from backend.rag import retrieve_context
from backend.router import TOOL_INFO, route_tools, sheet_only
from mcp_layer.client import MCPClient, MCPConnectionError

log = logging.getLogger("excel_agent")

# ---------------- LIMITS ----------------

MAX_STEPS = 6
MAX_RESULT_CHARS = 3000
MAX_ROWS_FOR_LLM = 20
KNOWLEDGE_TOP_K = 2
KNOWLEDGE_CHUNK_CHARS = 500
MAX_TOOLS_TO_LLM = 12        # fallback / widened list cap (schemas are short)
TOOL_TIMEOUT = 120
HISTORY_LIMIT = 20           # stored messages
HISTORY_WINDOW = 4           # messages sent to the LLM
MAX_SHEETS_CONTEXT = 5
PURPOSE_CHARS = 160
PARAM_DESC_CHARS = 90

# Parameters hidden from the LLM and controlled by the agent.
HIDDEN_PARAMS = {"workbook_name", "confirm", "overwrite"}
CONFIRM_PARAMS = ("confirm", "overwrite")
AUTO_CONFIRM = True 
# Tool parameters containing column names.
COLUMN_PARAMS = ("column", "group_by", "value_column")

# Tool parameters that must point to an EXISTING sheet.
SHEET_REF_PARAMS = (
    "sheet_name", "old_sheet_name", "source_sheet_name",
    "before_sheet", "after_sheet",
)

# For these tools "sheet_name" is a NEW name, so it is not validated
# against existing sheets. (Other sheet params are still validated.)
SHEET_CREATING_TOOLS = {"create_worksheet", "copy_worksheet", "rename_worksheet"}

# Tools that create new columns: skip column-name validation.
COLUMN_CHECK_SKIP = {"insert_excel_columns"}

# RAG runs only for these categories.
RAG_CATEGORIES = {
    "formulas", "analysis", "charts", "pivot",
    "dashboard", "conditional", "kb",
}

# If the request contains these, it may be multi-step: no fast path.
SINGLE_STEP_BLOCKERS = (" and ", " & ", ",", " + ")

STOPPED_REPLY = "Stopped."

# Phrases that show the model gave up without calling a tool.
INABILITY_MARKERS = (
    "do not have", "don't have", "no tool", "not have a tool", "cannot",
    "can't", "unable", "not able", "not available", "no available tool",
)


# ---------------- READ-ONLY TOOLS ----------------
# Must match the real tool names registered in mcp_layer/server.py.

READONLY_TOOLS = {
    # Both (8)
    "get_excel_sheet_names", "get_all_worksheets", "get_worksheet_info",
    "read_excel_data", "read_excel_range", "get_excel_column_info",
    "search_excel_value", "get_excel_formula",
    # Chat-only analysis / list (11)
    "analyze_excel_column", "calculate_excel_total",
    "value_counts_excel_column", "identify_excel_outliers",
    "compare_excel_datasets", "validate_excel_data",
    "list_excel_charts", "list_excel_pivots", "list_dashboards",
    "list_conditional_formatting", "list_workflows",
}

# Always part of the FALLBACK shortlist (reserved before the cap).
ALWAYS_TOOLS = {
    "get_workbook_info",
    "get_excel_sheet_names",
    "read_excel_data",
    "write_excel_data",
    "ask_knowledge_base",
}

GENERAL_QA = False   # general Excel theory is out of scope

DECLINE_REPLY = "Ask anything about this dataset."

CHAT_MODE_NOTE = (
    "\nMODE: Chat (analyst, read-only). Answer from the tool result in one "
    "or two plain sentences. Never paste raw JSON. "
    "Give the number or list directly."
)

CHAT_BLOCKED_REPLY = ("I can't change the workbook in Chat mode. "
                      "Switch to Edit mode.")

GREETING_WORDS = {
    "hi", "hii", "hello", "hey", "vanakkam",
    "thanks", "thank", "ok", "okay", "bye",
}

FABRICATION_REPLY = (
    "I couldn't verify the data through the tool, so I won't provide a number "
    "(I should not give incorrect information). "
    "Please try again or rephrase your request."
)


# ---------------- FALLBACK SHORTLISTING (categories) ----------------

CATEGORY_USER_WORDS = {
    "read": (
        "read", "show", "display", "view", "open", "list", "fetch",
        "get data", "retrieve", "preview", "see data", "check data",
        "contents", "rows", "columns", "records", "values",
    ),
    "write": (
        "write", "put", "enter", "set cell", "type", "update cell",
        "edit cell", "change cell", "modify cell", "fill cell",
        "write value", "insert value", "replace value",
    ),
    "append": (
        "append", "add record", "add row", "new record", "new row",
        "insert record", "add data", "add entry", "add new data",
        "append data", "add at the end",
    ),
    "sheets": (
        "sheet", "worksheet", "tab", "tab color", "tab colour",
        "rename", "duplicate sheet", "duplicate worksheet",
        "copy", "move", "hide", "unhide", "show sheet",
        "freeze", "unfreeze", "freeze panes",
        "protect", "unprotect", "lock sheet",
        "zoom", "gridline", "gridlines",
        "print", "print area", "landscape", "portrait",
        "details of", "sheet info", "sheet details",
        "list sheets", "all sheets", "create sheet", "new sheet",
        "delete sheet", "remove sheet", "switch sheet", "sheet name",
    ),
    "rows_cols": (
        "insert row", "delete row", "remove row", "add row",
        "insert column", "delete column", "remove column",
        "add column", "new row", "new column",
        "row insert", "column insert",
    ),
    "range": (
        "copy", "paste", "move range", "copy range", "paste range",
        "merge cells", "unmerge cells", "merge range",
        "duplicate range", "copy data to", "move data to",
    ),
    "formulas": (
        "formula", "apply formula", "insert formula", "excel formula",
        "sum formula", "average formula", "count formula",
        "countif", "sumif", "averageif", "vlookup", "xlookup",
        "index match", "if formula", "logical formula",
        "date formula", "text formula", "calculate using formula",
    ),
    "format": (
        "format", "formatting", "font", "font size", "font style",
        "font color", "background color", "fill color",
        "bold", "italic", "underline", "alignment",
        "align", "number format", "date format", "percentage format",
        "width", "height", "row height", "column width",
        "border", "autofit", "auto fit", "cell style",
    ),
    "conditional": (
        "highlight", "conditional formatting", "conditional format",
        "data bar", "color scale", "colour scale",
        "icon set", "dropdown", "drop down",
        "data validation", "validation list",
        "highlight duplicates", "highlight blanks", "highlight values",
    ),
    "clean": (
        "remove duplicates", "delete duplicates", "duplicate rows",
        "remove blank rows", "delete blank rows", "empty rows",
        "find and replace", "replace text", "replace values",
        "missing values", "fill missing", "handle missing",
        "clean data", "data cleaning", "data quality",
        "remove empty", "trim spaces", "fix data",
    ),
    "filter_sort": (
        "filter", "apply filter", "clear filter",
        "sort", "sort data", "sort ascending", "sort descending",
        "ascending", "descending", "order by",
        "highest", "lowest", "largest", "smallest",
        "rank", "top records", "bottom records",
        "arrange data", "filter data",
    ),
    "analysis": (
        "total", "sum of", "calculate total", "find total",
        "average", "mean", "minimum", "maximum",
        "min value", "max value", "count records",
        "statistics", "statistical analysis", "stats",
        "outlier", "compare", "comparison", "insight",
        "trend", "growth", "analyze", "analyse", "validate",
        "kpi", "distribution", "year wise", "year-wise",
        "monthly analysis", "summary", "performance analysis",
    ),
    "charts": (
        "chart", "charts", "graph", "graphs", "plot",
        "visualize", "visualise", "visualization", "visualisation",
        "pie chart", "bar chart", "column chart",
        "line chart", "scatter chart", "area chart",
        "doughnut chart", "combo chart", "create graph",
        "draw chart", "plot data",
    ),
    "pivot": (
        "pivot", "pivot table", "pivot report",
        "summary table", "group by", "grouped data",
        "grouping", "cross tabulation",
        "summarize by category", "summarise by category",
    ),
    "dashboard": (
        "dashboard", "create dashboard", "build dashboard",
        "auto dashboard", "guided dashboard",
        "kpi card", "summary page", "visual report",
        "management report", "executive dashboard",
    ),
    "import": (
        "import", "import data", "load file",
        "another file", "other workbook", "external workbook",
        "csv", "merge workbook", "append workbook",
        "compare workbook", "combine files",
        "consolidate reports", "copy from another file",
    ),
    "automation": (
        "workflow", "automate", "automation",
        "save steps", "repeat task", "recurring task",
        "macro", "run workflow", "execute workflow",
        "automate report", "multi step task",
    ),
    "web": (
        "web", "internet", "search online", "google",
        "latest information", "latest data", "news",
        "external data", "search web",
    ),
    "kb": (
        "what is", "what are", "how to", "how do",
        "explain", "meaning", "difference",
        "why", "definition", "shortcut",
        "formula syntax", "teach me", "help me understand",
    ),
}

# Real tool names (plus held-back tools, used automatically once registered).
CATEGORY_FRAGMENTS = {
    "read": (
        "read_excel_data", "get_excel_sheet_names", "get_excel_range_data",
        "get_column_information", "get_used_range", "get_workbook_info",
    ),
    "write": ("write_excel_data",),
    "append": ("append_excel_data",),
    "sheets": (
        "create_worksheet", "rename_worksheet", "delete_worksheet",
        "copy_worksheet", "move_worksheet", "hide_unhide_worksheet",
        "get_all_worksheets", "get_worksheet_info", "freeze_panes",
        "protect_unprotect_worksheet", "set_sheet_tab_color",
        "set_worksheet_zoom_view", "set_print_setup",
    ),
    "rows_cols": (
        "insert_excel_rows", "insert_excel_columns", "delete_excel_rows",
        "delete_excel_columns", "delete_excel_sheet",
    ),
    "range": (
        "merge_excel_cells", "unmerge_excel_cells", "copy_excel_range",
        "move_excel_range",
    ),
    "formulas": (
        "apply_excel_formula", "generate_excel_formula",
        "explain_excel_formula",
    ),
    "format": (
        "format_excel_cells", "apply_excel_borders", "autofit_excel_columns",
        "set_column_width", "set_row_height",
    ),
    "conditional": (
        "apply_conditional_formatting", "highlight_excel_values",
        "apply_data_validation",
    ),
    "clean": (
        "remove_excel_duplicates", "remove_excel_blank_rows",
        "find_replace_excel_data", "fill_missing_excel_values",
        "validate_excel_data",
    ),
    "filter_sort": ("sort_excel_data", "filter_excel_data", "clear_excel_filter"),
    "analysis": (
        "analyze_excel_column", "calculate_excel_total",
        "compare_excel_datasets", "generate_excel_insights",
        "identify_excel_outliers", "validate_excel_data",
    ),
    "charts": ("create_excel_chart", "modify_excel_chart"),
    "pivot": ("create_excel_pivot_summary", "modify_excel_pivot"),
    "dashboard": ("auto_dashboard", "guided_dashboard"),
    "import": (
        "import_excel_workbook", "merge_excel_workbooks",
        "append_excel_workbooks", "compare_excel_workbooks",
    ),
    "automation": (
        "list_workflows", "save_workflow", "run_workflow", "delete_workflow",
    ),
    "web": ("web_search",),
    "kb": ("ask_knowledge_base",),
}


# ---------------- AGENT BEHAVIOR ----------------

MULTI_STEP_HINTS = (
    " then ", " after that ", " after ", " next ", " followed by ",
    " once done ", " and then ", " before that ", " first ", " finally ",
)

EXPLICIT_ACTION_HINTS = (
    "apply ", "insert ", "write ", "put ", "enter ", "fill ",
    "update ", "change ", "modify ", "replace ", "delete ",
    "remove ", "create ", "add ", "calculate ", "compute ",
    "sort ", "filter ", "highlight ", "format ", "move ",
    "copy ", "paste ", "generate in the sheet", "in the workbook",
    "to this sheet", "to the sheet", "in excel cell", "cell ",
)

KNOWLEDGE_HINTS = (
    "what is ", "what are ", "how to ", "how do ", "how can ", "how does ",
    "explain ", "meaning of ", "difference between ",
    "what is the difference ", "why ", "why does ", "why is ",
    "formula for ", "excel formula ", "syntax of ", "example of ",
    "examples of ", "shortcut ", "excel shortcut ", "function of ",
    "how it works ", "how does it work ", "teach me ",
    "help me understand ",
)


# ---------------- SYSTEM PROMPT (short and static) ----------------

SYSTEM_PROMPT = (
    "You are an Excel agent controlling the user's open workbook through "
    "tools.\n"
    "Rules:\n"
    "- To read or change the workbook, call the matching tool. Do not just "
    "explain how.\n"
    "- Use only sheet names, column names and values from the user's message "
    "or the workbook context below. Never invent them.\n"
    "- Call one tool at a time and check its result. Never claim success "
    "unless the tool result confirms it.\n"
    "- If a tool returns an error, report that actual error. Say a tool is "
    "missing only if it is truly not in your tool list.\n"
    "- Destructive actions are confirmed by the system. Do not ask yourself.\n"
    "- Saving is not available: tell the user to press Cmd+S.\n"
    "- Keep replies short. Always reply in English. If the user writes "
    "(Tanglish -> simple Tanglish).\n"
    "- Greetings and general questions: answer directly without tools.\n"
)


# ---------------- KNOWLEDGE TOOL (short) ----------------

KNOWLEDGE_TOOL = {
    "type": "function",
    "function": {
        "name": "ask_knowledge_base",
        "description": (
            "Answer Excel concept or how-to questions from the knowledge "
            "base. Not for reading or changing workbook data."
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "question": {
                    "type": "string",
                    "description": "The user's Excel question.",
                }
            },
            "required": ["question"],
        },
    },
}


# ---------------- MODULE HELPERS ----------------

_HEADING_RE = re.compile(r"^[A-Z][A-Z ]+:$")
_PARAM_RE = re.compile(r"^([A-Za-z_][A-Za-z0-9_]*):\s*(.*)$")


def _parse_doc(doc: str | None) -> tuple[str, dict[str, str]]:
    """Return (one-line purpose, {param: short description}) from a docstring.

    Understands the PURPOSE / WHEN TO USE / PARAMETERS / RETURNS / NOTES
    layout. Falls back to the first lines when there are no headings.
    """
    text = (doc or "").strip()
    if not text:
        return "No description available.", {}

    section: str | None = None
    purpose: list[str] = []
    preface: list[str] = []
    params: dict[str, str] = {}
    last: str | None = None

    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if _HEADING_RE.match(line):
            section = line[:-1]
            last = None
            continue
        if section == "PURPOSE":
            purpose.append(line)
        elif section == "PARAMETERS":
            match = _PARAM_RE.match(line)
            if match and not line.startswith("-"):
                last = match.group(1)
                params[last] = match.group(2).strip()
            elif last:
                params[last] = (params[last] + " " + line).strip()
        elif section is None:
            preface.append(line)

    first = " ".join(purpose) or " ".join(preface[:2])
    return (first[:PURPOSE_CHARS] or "No description available."), params


def _to_ollama_tool(tool: Any) -> dict[str, Any]:
    """Convert an MCP tool into a SHORT Ollama function-calling schema."""
    schema = copy.deepcopy(tool.inputSchema or {})
    properties = schema.get("properties", {})

    for param in HIDDEN_PARAMS:
        properties.pop(param, None)

    schema["required"] = [
        p for p in schema.get("required", []) if p not in HIDDEN_PARAMS
    ]
    schema.pop("title", None)

    purpose, param_docs = _parse_doc(tool.description)
    for pname, prop in properties.items():
        if isinstance(prop, dict):
            prop.pop("title", None)
            doc = param_docs.get(pname)
            if doc:
                prop["description"] = doc[:PARAM_DESC_CHARS]

    return {
        "type": "function",
        "function": {
            "name": tool.name,
            "description": purpose,
            "parameters": schema,
        },
    }


def _is_connection_error(exc: BaseException) -> bool:
    if isinstance(exc, (MCPConnectionError, ConnectionError)):
        return True
    name = type(exc).__name__
    return any(w in name for w in ("Closed", "BrokenResource", "EndOfStream"))


# ============================================================
# EXCEL AGENT
# ============================================================

class ExcelAgent:

    def __init__(self, llm: OllamaEngine | None = None,
                 mcp: MCPClient | None = None) -> None:
        self.llm = llm or OllamaEngine()
        self.mcp = mcp or MCPClient()

        self.tools: list[dict[str, Any]] = []
        self.tool_names: set[str] = set()
        self.tool_params: dict[str, set[str]] = {}   # tool -> schema param names

        self.history: list[dict[str, str]] = []
        self.pending: dict[str, Any] | None = None
        self.stop_event = threading.Event()

        # Workbook context
        self.workbook = ""
        self.sheets: list[str] = []
        self.active_sheet = ""
        self.columns: list[str] = []
        self.total_rows = 0
        self.sheet_columns: dict[str, list[str]] = {}
        self.mcp_available = False
        self.mcp_error = ""

    # ---------------- CONNECTION ----------------

    @staticmethod
    def _exception_details(exc: BaseException) -> str:
        """Expand ExceptionGroup so the real nested MCP failure is visible."""
        lines: list[str] = []

        def walk(error: BaseException, depth: int = 0) -> None:
            prefix = "  " * depth
            lines.append(f"{prefix}{type(error).__name__}: {error}")
            nested = getattr(error, "exceptions", None)
            if nested:
                for child in nested:
                    walk(child, depth + 1)

        walk(exc)
        return "\n".join(lines)

    async def connect(self) -> None:
        """Connect MCP tools; keep knowledge-only Chat usable if MCP is offline."""
        try:
            await self.mcp.connect()
            mcp_tools = await self.mcp.get_tools()

            self.tools = [_to_ollama_tool(t) for t in mcp_tools] + [KNOWLEDGE_TOOL]
            self.tool_names = {t.name for t in mcp_tools} | {"ask_knowledge_base"}
            self.tool_params = {
                t.name: set((t.inputSchema or {}).get("properties", {}))
                for t in mcp_tools
            }
            self.mcp_available = True
            self.mcp_error = ""
            log.info("%d MCP tools registered, %d total LLM tools, "
                     "short schemas ~%d chars",
                     len(mcp_tools), len(self.tools),
                     len(json.dumps(self.tools)))

            # Spot tools the router and the fallback can never reach.
            known = {n for frags in CATEGORY_FRAGMENTS.values() for n in frags}
            unreachable = sorted(
                n for n in self.tool_names
                if n not in TOOL_INFO and not any(f in n for f in known)
            )
            if unreachable:
                log.warning("Tools never shortlisted (add to router/fragments): %s",
                            unreachable)

        except Exception as exc:
            self.mcp_available = False
            self.mcp_error = self._exception_details(exc)
            log.exception("MCP initialization failed; starting knowledge-only mode")

            try:
                await self.mcp.disconnect()
            except Exception:
                log.exception("MCP cleanup after failed initialization also failed")

            self.tools = [KNOWLEDGE_TOOL]
            self.tool_names = {"ask_knowledge_base"}
            self.tool_params = {}

    async def close(self) -> None:
        if self.mcp_available:
            try:
                await self.mcp.disconnect()
            finally:
                self.mcp_available = False

    async def _reconnect(self) -> None:
        try:
            await self.mcp.disconnect()
        except Exception:
            log.exception("Error while closing MCP before reconnect")
        try:
            await self.mcp.connect()
            self.mcp_available = True
            self.mcp_error = ""
        except Exception as exc:
            self.mcp_available = False
            self.mcp_error = self._exception_details(exc)
            raise

    def request_stop(self) -> None:
        """Called by the UI Stop button (thread-safe)."""
        self.stop_event.set()

    # ---------------- SAFE TOOL CALL (single gate) ----------------

    async def _call(self, name: str, args: dict[str, Any],
                    mode: str) -> dict[str, Any]:
        """All MCP calls go through here: mode gate + timeout + error handling."""
        if mode == "chat" and name not in READONLY_TOOLS:
            return {"success": False, "blocked": True,
                    "error": "Chat mode is read-only."}
        try:
            return await self._call_once(name, args)
        except asyncio.TimeoutError:
            msg = f"Tool '{name}' timed out."
            if name not in READONLY_TOOLS:
                msg += " It may or may not have been applied. Check the workbook."
            return {"success": False, "error": msg}
        except Exception as exc:
            if _is_connection_error(exc):
                if name in READONLY_TOOLS:       # safe to retry
                    try:
                        await self._reconnect()
                        return await self._call_once(name, args)
                    except Exception as exc2:
                        return {"success": False,
                                "error": f"Reconnect failed: {exc2}"}
                return {"success": False,
                        "error": "Connection lost during a write. It may or "
                                 "may not have been applied. Check the "
                                 "workbook before retrying."}
            return {"success": False,
                    "error": f"{type(exc).__name__}: {exc}"}

    async def _call_once(self, name: str, args: dict[str, Any]) -> dict[str, Any]:
        result = await asyncio.wait_for(
            self.mcp.call_tool(name, args), TOOL_TIMEOUT
        )
        if not isinstance(result, dict):
            return {"success": False, "error": "Tool returned invalid result."}
        return result

    # ---------------- MAIN REQUEST ----------------

    async def run(self, user_text: str, mode: str = "edit") -> dict[str, Any]:
        """Run one user request. Returns reply, steps, needs_confirmation."""
        self.pending = None
        self.stop_event.clear()

        categories = self._categories(user_text)
        action_cats = categories - {"kb"}

        # Explain-only requests (e.g. "What is VLOOKUP?") need no workbook.
        if GENERAL_QA and self._is_explanation_only(user_text):
            try:
                kb = await asyncio.to_thread(answer_question, user_text)
            except Exception as exc:
                log.exception("Knowledge answer failed")
                kb = {"success": False, "error": self._exception_details(exc)}

            if isinstance(kb, dict) and kb.get("success") and kb.get("answer"):
                return self._finish(user_text, str(kb["answer"]), [])

            if not self.mcp_available:
                detail = (kb.get("error") if isinstance(kb, dict) else None) or \
                         "The Excel knowledge base did not return an answer."
                return {
                    "success": False,
                    "error": f"Knowledge answer unavailable: {detail}\n"
                             f"MCP is also unavailable: {self.mcp_error}",
                    "steps": [],
                }

        if not self.mcp_available:
            return {
                "success": False,
                "error": "Excel workbook tools are unavailable because the MCP "
                         "connection failed.\n"
                         f"Underlying error: {self.mcp_error}",
                "steps": [],
            }

        if not self.tool_names:
            return {"success": False, "error": "Agent connect() was not called."}

        await self._load_context()

        # No category matched but the request names a real sheet.
        if not action_cats and self._mentions_sheet(user_text):
            categories = categories | {"sheets"}
            action_cats = categories - {"kb"}

        knowledge = ""
        if (self._should_retrieve_knowledge(user_text)
                and (categories & RAG_CATEGORIES)):
            knowledge = await asyncio.to_thread(self._knowledge_for, user_text)

        # ---- Tool selection: router first, category shortlist as fallback ----
        route = route_tools(user_text, self.tool_names, self.sheets, mode,
                            READONLY_TOOLS)
        routed: list[str] = route["tools"]
        if mode == "chat":
            probe = route_tools(user_text, self.tool_names, self.sheets,
                                "edit", READONLY_TOOLS)["tools"]
            if probe and probe[0] not in READONLY_TOOLS:
                return self._finish(user_text, CHAT_BLOCKED_REPLY, [])
            # Scope check: out-of-dataset question -> fixed reply from code.
            cols = self.sheet_columns.get(self.active_sheet) or self.columns
            judge = [
                {"role": "system", "content": (
                    "You decide if a question is about the data in an Excel "
                    "workbook.\n"
                    f"Sheets: {self.sheets}\nColumns: {cols}\n"
                    "Reply YES if the question is about this workbook's data, "
                    "sheets or columns (counts, sums, averages, names, values). "
                    "Reply NO if it is general knowledge, Excel theory, "
                    "chit-chat, weather or any other topic. "
                    "Answer with one word: YES or NO.")},
                {"role": "user", "content": user_text},
            ]
            try:
                verdict = await asyncio.to_thread(
                    self.llm.chat, judge, None, 3, False, None)
                if (not any(t in READONLY_TOOLS for t in routed)
                        and verdict.get("success")
                        and str(verdict.get("content", "")).strip()
                        .upper().startswith("NO")):
                    return self._finish(user_text, DECLINE_REPLY, [])
            except Exception:
                log.exception("scope check failed")
        if routed:
            names = list(routed)
            if (mode != "chat" and "kb" in categories
                    and "ask_knowledge_base" not in names):
                names.append("ask_knowledge_base")
            active_tools = self._tools_by_names(names)
        else:
            active_tools = self._tools_by_names(self._category_names(categories))
        if mode == "chat":
            active_tools = self._chat_filter(active_tools)
            if not active_tools and action_cats - {"read", "analysis", "kb"}:
                return self._finish(user_text, CHAT_BLOCKED_REPLY, [])
            
        slim = bool(routed) and sheet_only(routed)
        log.info("route=%s fallback=%s active=%s",
                 routed, not routed,
                 [t["function"]["name"] for t in active_tools])

        mode_note = CHAT_MODE_NOTE if mode == "chat" else ""
        messages: list[dict[str, Any]] = (
            [{"role": "system",
              "content": self._build_system_prompt(knowledge, slim) + mode_note}]
            + self.history[-HISTORY_WINDOW:]
            + [{"role": "user", "content": user_text}]
        )

        steps: list[dict[str, Any]] = []
        nudged = False
        widened = False

        for _ in range(MAX_STEPS):
            if self.stop_event.is_set():
                return self._finish(user_text, STOPPED_REPLY, steps)

            res = await asyncio.to_thread(
                self.llm.chat, messages, active_tools, None, False,
                self.stop_event,
            )
            if res.get("stopped"):
                return self._finish(user_text, STOPPED_REPLY, steps)
            if not res["success"]:
                return {"success": False, "error": res["error"], "steps": steps}

            calls = res["tool_calls"]

            if not calls:
                # Safety net: the router may have removed the right tool.
                if (not steps and routed and not widened and action_cats
                        and self._claims_inability(res["content"])):
                    widened = True
                    active_tools = self._widen(
                        routed, route["removed"], categories, mode)
                    log.info("widened tools: %s",
                             [t["function"]["name"] for t in active_tools])
                    continue

                if (not steps and not nudged
                        and self._should_nudge(res["content"], bool(action_cats))):
                    nudged = True
                    messages.append({
                        "role": "user",
                        "content": (
                            "Do not answer from memory or guess. Use the "
                            f'matching tool now for: "{user_text}". If it '
                            "truly cannot be done with the tools, say so briefly."
                        ),
                    })
                    continue
                if (not steps and nudged
                        and self._looks_fabricated(res["content"])):
                    return self._finish(user_text, FABRICATION_REPLY, steps)
                return self._finish(user_text, res["content"], steps)

            messages.append({
                "role": "assistant",
                "content": res["content"],
                "tool_calls": [
                    {"function": {"name": c["name"], "arguments": c["arguments"]}}
                    for c in calls
                ],
            })

            for call in calls:
                if self.stop_event.is_set():
                    return self._finish(user_text, STOPPED_REPLY, steps)

                name = call["name"]
                raw = dict(call["arguments"])

                if mode == "chat" and name not in READONLY_TOOLS:
                    return self._finish(user_text, CHAT_BLOCKED_REPLY, steps)

                if name not in self.tool_names:
                    args = raw
                    result = {"success": False, "error": f"Unknown tool: {name}"}

                elif name == "ask_knowledge_base":
                    args = raw
                    kb = await asyncio.to_thread(
                        answer_question, args.get("question") or user_text
                    )
                    result = {
                        "success": kb["success"],
                        "answer": kb.get("answer", kb.get("error")),
                        "sources": kb.get("sources", []),
                    }

                else:
                    args, err = self._sanitize_args(name, raw)
                    if err:
                        result = {"success": False, "error": err}
                    else:
                        result = await self._call(name, args, mode)
                        if result.get("blocked"):
                            return self._finish(user_text, CHAT_BLOCKED_REPLY, steps)
                        self._annotate_read(name, args, result)

                if (mode == "edit" and name == "calculate_excel_total"
                        and result.get("success")):
                    note = await self._write_total_to_cell(args, result, user_text)
                    if note:
                        result["written"] = note
                        result["message"] = (
                            f"{str(args.get('operation', 'sum')).title()} of "
                            f"{args.get('column')} = {result.get('result')}. {note}"
                        )
                steps.append({"tool": name, "arguments": args, "result": result})
                messages.append({
                    "role": "tool",
                    "tool_name": name,
                    "content": self._compact(result),
                })

                param = self._needs_confirmation(name, result)
                if param:
                    if AUTO_CONFIRM:
                        self.pending = {"tool": name, "arguments": args, "param": param}
                        return await self.confirm_pending()
                    self.pending = {"tool": name, "arguments": args, "param": param}
                    note = (result.get("message") or result.get("error")
                            or "Confirmation needed.")
                    reply = f"{note} Do you want to proceed?"
                    self._remember(user_text, reply, steps)
                    return {"success": True, "reply": reply, "steps": steps,
                            "needs_confirmation": True}

                if name not in READONLY_TOOLS and result.get("success"):
                    await self._load_context()
                    messages[0]["content"] = (
                        self._build_system_prompt(knowledge, slim) + mode_note
                    )

                # Fast path: 1 call, 1 step, single-intent, ready-made reply.
                if (len(calls) == 1 and len(steps) == 1
                        and result.get("success")
                        and self._is_single_step(user_text)):
                    fast = (result.get("message") or result.get("answer")
                            or self._quick_reply(name, result))
                    if fast:
                        return self._finish(user_text, str(fast), steps)

        return self._finish(
            user_text,
            "Reached the maximum number of steps. Please simplify the request.",
            steps,
        )

    # ---------------- CONFIRMATION ----------------

    async def confirm_pending(self) -> dict[str, Any]:
        """Execute the pending action after user confirmation."""
        if not self.pending:
            return {"success": False, "error": "Nothing is pending to confirm."}

        name = self.pending["tool"]
        args = dict(self.pending["arguments"])
        param = self.pending.get("param", "overwrite")
        self.pending = None

        args[param] = True
        result = await self._call(name, args, "edit")

        reply = (result.get("message") or self._quick_reply(name, result)
                 or result.get("error") or "Done.")
        if result.get("deleted") and result.get("removed") is not None:
            reply = (f"Deleted {result['removed']} rows. "
                     f"Backup: {result.get('backup')}")

        steps = [{"tool": name, "arguments": args, "result": result}]
        self._remember("yes, confirmed", reply, steps)
        if result.get("success"):
            await self._load_context()
        return {"success": bool(result.get("success")), "reply": reply,
                "steps": steps, "needs_confirmation": False}

    def cancel_pending(self) -> None:
        if self.pending:
            self._remember("no, cancel", "Cancelled. Nothing was changed.")
        self.pending = None

    def _needs_confirmation(self, name: str, result: dict[str, Any]) -> str | None:
        """Return the confirm param name (from the schema), else None."""
        params = self.tool_params.get(name, set())
        error = str(result.get("error", "")).lower()

        if not result.get("success"):
            for p in CONFIRM_PARAMS:
                if p in params and p in error:
                    return p
            return None

        # Dry-run / preview: nothing deleted yet, confirmation needed.
        if ("confirm" in params
                and name.startswith(("remove_", "delete_"))
                and not result.get("deleted")):
            found = result.get("found")
            if found is None:
                if result.get("deleted") is False:
                    return "confirm"
            elif (found or 0) > 0:
                return "confirm"

        # Any tool whose preview message asks for confirmation.
        msg = str(result.get("message", "")).lower()
        if result.get("success") and ("confirm to apply" in msg
                                      or "confirm=true" in msg):
            for p in CONFIRM_PARAMS:
                if p in params:
                    return p

        return None

    # ---------------- REQUEST CLASSIFIERS ----------------

    @staticmethod
    def _words(user_text: str) -> list[str]:
        return (user_text.strip().lower().replace(",", " ")
                .replace("!", " ").replace("?", " ").split())

    @classmethod
    def _should_retrieve_knowledge(cls, user_text: str) -> bool:
        words = cls._words(user_text)
        if not words:
            return False
        return not (words[0] in GREETING_WORDS and len(words) <= 5)

    @staticmethod
    def _is_multi_step(user_text: str) -> bool:
        text = f" {user_text.lower()} "
        return any(h in text for h in MULTI_STEP_HINTS)

    @classmethod
    def _is_single_step(cls, user_text: str) -> bool:
        text = f" {user_text.lower()} "
        return (not cls._is_multi_step(user_text)
                and not any(b in text for b in SINGLE_STEP_BLOCKERS))

    @staticmethod
    def _is_knowledge_question(user_text: str) -> bool:
        text = f" {user_text.lower()} "
        return any(h in text for h in KNOWLEDGE_HINTS)

    @classmethod
    def _is_explanation_only(cls, user_text: str) -> bool:
        """True for concept/how-to questions, false for explicit workbook edits."""
        text = f" {user_text.lower().strip()} "
        if not cls._is_knowledge_question(user_text):
            return False
        return not any(h in text for h in EXPLICIT_ACTION_HINTS)

    @staticmethod
    def _categories(user_text: str) -> set[str]:
        """Whole-word match (so 'tab' does not match 'table')."""
        text = f" {user_text.lower()} "
        found = set()
        for category, words in CATEGORY_USER_WORDS.items():
            for w in words:
                pattern = r"(?<!\w)" + re.escape(w.strip()) + r"(?:s|es)?(?!\w)"
                if re.search(pattern, text):
                    found.add(category)
                    break
        return found

    def _mentions_sheet(self, user_text: str) -> bool:
        """True if the request names a real sheet of the open workbook."""
        text = user_text.lower()
        for sheet in self.sheets:
            name = str(sheet).strip().lower()
            if len(name) < 2:
                continue
            if re.search(r"(?<!\w)" + re.escape(name) + r"(?!\w)", text):
                return True
        return False

    @staticmethod
    def _claims_inability(content: str) -> bool:
        reply = (content or "").lower()
        return any(m in reply for m in INABILITY_MARKERS)

    # ---------------- TOOL LISTS ----------------

    def _category_names(self, categories: set[str]) -> list[str]:
        """Fallback shortlist: category tools, ALWAYS tools reserved first."""
        category_names: list[str] = []
        for category, fragments in CATEGORY_FRAGMENTS.items():   # stable order
            if category not in categories:
                continue
            for fragment in fragments:
                for name in sorted(self.tool_names):
                    if fragment in name and name not in category_names:
                        category_names.append(name)

        always = [n for n in sorted(ALWAYS_TOOLS) if n in self.tool_names]
        budget = max(0, MAX_TOOLS_TO_LLM - len(always))
        chosen = [n for n in category_names if n not in always][:budget]
        return chosen + always

    def _tools_by_names(self, names: list[str]) -> list[dict[str, Any]]:
        by_name = {t["function"]["name"]: t for t in self.tools}
        return [by_name[n] for n in names if n in by_name]

    @staticmethod
    def _chat_filter(tools: list[dict[str, Any]]) -> list[dict[str, Any]]:
        return [t for t in tools if t["function"]["name"] in READONLY_TOOLS]

    def _widen(self, routed: list[str], removed: dict[str, int],
               categories: set[str], mode: str) -> list[dict[str, Any]]:
        """Wider list for the one-time retry: router picks, removed tools,
        then the category shortlist."""
        names: list[str] = []
        for n in list(routed) + list(removed) + self._category_names(categories):
            if n in self.tool_names and n not in names:
                names.append(n)
        tools = self._tools_by_names(names[:MAX_TOOLS_TO_LLM])
        return self._chat_filter(tools) if mode == "chat" else tools

    # ---------------- WORKBOOK CONTEXT ----------------

    async def _load_context(self) -> None:
        """Real sheets + headers of every sheet (max 5)."""
        self.workbook = ""
        self.sheets = []
        self.active_sheet = ""
        self.columns = []
        self.total_rows = 0
        self.sheet_columns = {}

        names = await self._call("get_excel_sheet_names", {}, "chat")
        if not names.get("success"):
            return

        self.workbook = names.get("workbook", "")
        self.sheets = list(names.get("sheets", []))
        self.active_sheet = names.get("active_sheet", "")

        for sheet in self.sheets[:MAX_SHEETS_CONTEXT]:
            head = await self._call(
                "read_excel_data", {"max_rows": 1, "sheet_name": sheet}, "chat"
            )
            if head.get("success"):
                cols = list(head.get("columns", []))
                self.sheet_columns[sheet] = cols
                if sheet == self.active_sheet:
                    self.columns = cols
                    self.total_rows = int(head.get("total_rows", 0))

    # ---------------- SYSTEM CONTEXT ----------------

    def _build_system_prompt(self, knowledge: str = "", slim: bool = False) -> str:
        """Static prompt first, workbook context last.

        slim=True (sheet-only requests) leaves out columns and headers.
        """
        if not self.sheets:
            context = "\nNo workbook data could be read right now."
        elif slim:
            return (SYSTEM_PROMPT
                    + f"\nOpen workbook: {self.workbook}. Sheets: {self.sheets}. "
                      f"Active sheet: {self.active_sheet}.")
        else:
            context = (f"\nOpen workbook: {self.workbook}. Sheets: {self.sheets}. "
                       f"Active sheet: {self.active_sheet} "
                       f"({self.total_rows} data rows).")

        if self.columns:
            letters = ", ".join(
                f"{self._col_letter(i)}={n}" for i, n in enumerate(self.columns)
            )
            last_row = self.total_rows + 1
            free_col = self._col_letter(len(self.columns))
            context += (
                f" Columns (use these exact names): {self.columns}. "
                f"Column letters: {letters}. "
                f"Data is in rows 2 to {last_row}. "
                f"First empty column for results: {free_col}."
            )

        others = {s: c for s, c in self.sheet_columns.items()
                  if s != self.active_sheet and c}
        if others:
            context += f" Headers of other sheets: {others}."

        if knowledge:
            context += (
                "\n\nRELEVANT EXCEL KNOWLEDGE (retrieved):\n" + knowledge +
                "\n\nUse it as domain guidance, not a fixed template. Tool "
                "schemas are the authority for execution. If it is irrelevant, "
                "rely on verified workbook context. Never invent details."
            )
        return SYSTEM_PROMPT + context

    @staticmethod
    def _col_letter(index: int) -> str:
        """0 -> A, 1 -> B, 25 -> Z, 26 -> AA."""
        letters = ""
        index += 1
        while index:
            index, rem = divmod(index - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    @staticmethod
    def _knowledge_for(user_text: str) -> str:
        """Retrieve top knowledge-base chunks."""
        try:
            chunks = retrieve_context(user_text, top_k=KNOWLEDGE_TOP_K)
        except Exception:
            return ""
        return "\n---\n".join(c["content"][:KNOWLEDGE_CHUNK_CHARS] for c in chunks)

    # ---------------- ARGUMENT SANITIZING ----------------

    @staticmethod
    def _match(value: str, options: list[str]) -> str | None:
        if value in options:
            return value
        lowered = value.strip().lower()
        for option in options:
            if str(option).strip().lower() == lowered:
                return option
        return None

    def _sanitize_args(self, tool: str,
                       raw: dict[str, Any]) -> tuple[dict[str, Any], str | None]:
        """Returns (args, error). If error is set the tool is NOT run;
        the error goes back to the model so it can correct itself."""
        args = {k: v for k, v in raw.items() if k not in HIDDEN_PARAMS}

        # 1. Sheet references must point to an existing sheet.
        if self.sheets:
            for key in SHEET_REF_PARAMS:
                value = args.get(key)
                if not isinstance(value, str):
                    continue
                if key == "sheet_name" and tool in SHEET_CREATING_TOOLS:
                    continue            # new name, not an existing sheet
                if not value.strip():
                    args.pop(key)       # empty optional value: drop it
                    continue
                real = self._match(value, self.sheets)
                if not real:
                    return args, (f"Sheet '{value}' does not exist. "
                                  f"Available sheets: {self.sheets}. "
                                  "Use an exact sheet name.")
                args[key] = real

        # 2. Columns: check against the headers of the target sheet.
        target = args.get("sheet_name", self.active_sheet)
        cols = self.sheet_columns.get(target) or (
            self.columns if target == self.active_sheet else []
        )
        if cols and tool not in COLUMN_CHECK_SKIP:
            for key in COLUMN_PARAMS:
                value = args.get(key)
                if isinstance(value, str):
                    fixed = self._match(value, cols)
                    if fixed:
                        args[key] = fixed
                    elif not re.fullmatch(r"[A-Z]{1,3}", value):  # letter is OK
                        return args, (f"Column '{value}' does not exist. "
                                      f"Available columns: {cols}.")
            if isinstance(args.get("columns"), list):
                fixed_list = []
                for c in args["columns"]:
                    if isinstance(c, str):
                        fc = self._match(c, cols)
                        if not fc and not re.fullmatch(r"[A-Z]{1,3}", c):
                            return args, (f"Column '{c}' does not exist. "
                                          f"Available columns: {cols}.")
                        fixed_list.append(fc or c)
                    else:
                        fixed_list.append(c)
                args["columns"] = fixed_list

        # 3. Read row limit.
        if tool == "read_excel_data":
            try:
                rows = int(args.get("max_rows", MAX_ROWS_FOR_LLM))
            except (TypeError, ValueError):
                rows = MAX_ROWS_FOR_LLM
            args["max_rows"] = max(1, min(rows, MAX_ROWS_FOR_LLM))

        return args, None

    @staticmethod
    def _annotate_read(name: str, args: dict[str, Any],
                       result: dict[str, Any]) -> None:
        """Tell the model about partial reads (prevents wrong totals)."""
        if name != "read_excel_data" or not result.get("success"):
            return
        try:
            total = int(result.get("total_rows", 0))
        except (TypeError, ValueError):
            return
        shown = args.get("max_rows", MAX_ROWS_FOR_LLM)
        if total > shown:
            result["note"] = (
                f"Only first {shown} of {total} rows shown. Do NOT compute "
                "totals/averages/counts from this sample. Use an analysis "
                "or formula tool."
            )


    async def _write_total_to_cell(self, args: dict[str, Any],
                                   result: dict[str, Any],
                                   user_text: str) -> str | None:
        """Edit mode: write a calculated total into a cell. Returns note or None."""
        value = result.get("result")
        if not isinstance(value, (int, float)):
            return None

        sheet = args.get("sheet_name") or self.active_sheet
        cols = self.sheet_columns.get(sheet) or self.columns
        column = args.get("column")
        if column not in cols:
            return None

        m = re.search(r"\b(?:in|at|to)\s+([A-Za-z]{1,3}\d{1,7})\b", user_text)
        if m:
            cell = m.group(1).upper()
        else:
            letter = self._col_letter(cols.index(column))
            rows = self.total_rows if sheet == self.active_sheet else 0
            if not rows:
                return None
            cell = f"{letter}{rows + 2}"

        if not m:
            # Auto-picked cell: never overwrite existing data.
            probe_cell = await self._call(
                "read_excel_range",
                {"cell_range": cell, "sheet_name": sheet}, "chat")
            vals = probe_cell.get("values") or [[None]]
            if probe_cell.get("success") and any(
                    v not in (None, "") for row in vals for v in row):
                return None 
                # cell not empty: never overwrite data
        write = await self._call(
            "update_excel_cell",
            {"cell": cell, "value": value, "sheet_name": sheet},
            "edit",
        )
        if write.get("success"):
            return f"Written to {sheet}!{cell}."
        return None

    # ---------------- HELPERS ----------------

    @staticmethod
    def _quick_reply(name: str, result: dict[str, Any]) -> str | None:
        """Ready-made reply for worksheet tools (saves one LLM call)."""
        if not result.get("success"):
            return None
        g = result.get
        try:
            if name == "create_worksheet":
                return f"Created sheet '{g('created')}'."
            if name == "rename_worksheet":
                return f"Renamed '{g('old_name')}' to '{g('new_name')}'."
            if name == "copy_worksheet":
                return f"Copied '{g('source_sheet')}' as '{g('copied_sheet')}'."
            if name == "delete_worksheet" and g("deleted"):
                return f"Deleted sheet '{g('sheet')}'."
            if name == "move_worksheet":
                return f"Moved '{g('moved')}' to position {g('new_position')}."
            if name == "hide_unhide_worksheet":
                return f"Sheet '{g('sheet')}' is now {g('state')}."
            if name == "set_sheet_tab_color":
                if g("tab_color") is None:
                    return f"Cleared the tab color of '{g('sheet')}'."
                return f"Set the tab color of '{g('sheet')}' to {g('tab_color')}."
            if name == "protect_unprotect_worksheet":
                state = "protected" if g("protected") else "unprotected"
                return f"Sheet '{g('sheet')}' is now {state}."
            if name == "freeze_panes":
                return (f"'{g('sheet')}': frozen rows {g('frozen_rows')}, "
                        f"frozen columns {g('frozen_columns')}.")
            if name == "set_worksheet_zoom_view":
                grid = "on" if g("show_gridlines") else "off"
                return f"'{g('sheet')}': zoom {g('zoom')}%, gridlines {grid}."
            if name == "set_print_setup":
                return (f"'{g('sheet')}': {g('orientation')}, print area "
                        f"{g('print_area') or 'none'}, fit one page wide: "
                        f"{g('fit_to_one_page_wide')}.")
            if name == "get_all_worksheets":
                items = g("sheets") or []
                parts = [
                    s["name"] if s.get("state", "visible") == "visible"
                    else f"{s['name']} ({s['state']})"
                    for s in items
                ]
                return f"{len(parts)} sheets: " + ", ".join(parts) + "."
            if name == "get_worksheet_info":
                return (f"'{g('sheet')}': {g('rows')} rows x {g('columns')} "
                        f"columns, used range {g('used_range')}, {g('state')}, "
                        f"protected: {g('protected')}, frozen rows "
                        f"{g('frozen_rows')}, zoom {g('zoom')}.")
            if name == "fill_excel_series":
                return f"Filled {g('count')} cells in {g('written_range')}."
        except Exception:
            return None
        return None

    @staticmethod
    def _should_nudge(content: str, has_action: bool) -> bool:
        """Nudge only for action-looking requests (saves an extra CPU call)."""
        if not has_action:
            return False
        reply = (content or "").lower()
        return not any(m in reply for m in
                       ("cmd+s", "cannot", "can't", "unable to"))

    @staticmethod
    def _looks_fabricated(content: str) -> bool:
        return any(ch.isdigit() for ch in content or "")

    @staticmethod
    def _compact(result: dict[str, Any]) -> str:
        text = json.dumps(result, default=str, ensure_ascii=False)
        if len(text) <= MAX_RESULT_CHARS:
            return text
        slim = dict(result)
        for key, value in result.items():
            if isinstance(value, list) and len(value) > 5:
                slim[key] = value[:5]
                slim["truncated"] = True
        text = json.dumps(slim, default=str, ensure_ascii=False)
        return (text if len(text) <= MAX_RESULT_CHARS
                else text[:MAX_RESULT_CHARS] + "...(truncated)")

    @staticmethod
    def _short_args(args: dict[str, Any]) -> str:
        text = json.dumps(args, default=str, ensure_ascii=False)
        return text if len(text) <= 80 else text[:80] + "..."

    def _remember(self, user_text: str, reply: str,
                  steps: list[dict[str, Any]] | None = None) -> None:
        """Save tool actions in history too (context for 'undo that')."""
        stored = reply
        if steps:
            acts = "; ".join(
                f"{s['tool']}({self._short_args(s['arguments'])})"
                f"->{'ok' if s['result'].get('success') else 'fail'}"
                for s in steps[-5:]
            )
            stored = f"{reply}\n[actions: {acts}]"
        self.history += [
            {"role": "user", "content": user_text},
            {"role": "assistant", "content": stored},
        ]
        self.history = self.history[-HISTORY_LIMIT:]

    def _finish(self, user_text: str, reply: str,
                steps: list[dict[str, Any]]) -> dict[str, Any]:
        reply = (reply or "").split("[actions:")[0].rstrip()
        if reply.lstrip().startswith("{") and steps:
            last = steps[-1]["result"]
            vals = last.get("values") or last.get("data")
            if isinstance(vals, list):
                if vals and isinstance(vals[0], list):
                    reply = "\n".join(
                        " | ".join("" if v is None else str(v) for v in row)
                        for row in vals)
                else:
                    shown = ", ".join(str(v) for v in vals[:20])
                    reply = f"{last.get('column', 'Values')}: {shown}"
        if not reply.strip() and steps:
            last = steps[-1]["result"]
            reply = str(last.get("message") or last.get("error")
                        or self._compact(last))
        self._remember(user_text, reply, steps)
        return {"success": True, "reply": reply, "steps": steps,
                "needs_confirmation": False}


# ============================================================
# SMOKE TEST
# ============================================================

async def _smoke_test() -> None:
    agent = ExcelAgent()
    await agent.connect()

    try:
        for prompt in [
            "Hello",
            "which sheets are in this workbook?",
        ]:
            out = await agent.run(prompt)
            print(f"\nYOU: {prompt}")
            print("AGENT:", out.get("reply") or out.get("error"))
    finally:
        await agent.close()


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    asyncio.run(_smoke_test())