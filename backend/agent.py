"""Excel agent: LLM decision loop with native tool calling over MCP.

Connect once (connect), run many requests (run), then close.
Returns dicts only: no print()/input() here.
Every request auto-retrieves knowledge-base chunks (RAG) so the LLM can pick formulas itself.
Smoke test (Ollama + Excel open, read-only prompts): python -m backend.agent
"""
import asyncio
import copy
import json
from typing import Any

from backend.knowledge_answer import answer_question
from backend.ollama_engine import OllamaEngine
from backend.rag import retrieve_context
from mcp_layer.client import MCPClient

MAX_STEPS = 5
MAX_RESULT_CHARS = 3000
MAX_ROWS_FOR_LLM = 20
KNOWLEDGE_TOP_K = 3
KNOWLEDGE_CHUNK_CHARS = 700
# Hidden from the model; the agent controls these, never the LLM.
HIDDEN_PARAMS = {"workbook_name", "confirm", "overwrite"}
# Tool parameters that hold a column name (matched case-insensitively against real columns).
COLUMN_PARAMS = ("column", "group_by", "value_column")
GREETING_WORDS = {"hi", "hii", "hello", "hey", "vanakkam", "thanks", "thank", "ok", "okay", "bye"}
FABRICATION_REPLY = (
    "Data-va tool vazhiya check panna mudiyala, adhunaala number solla maaten (thappaana number kudukka koodaadhu). "
    "Thirumba try pannunga, illa request-a konjam vera maadhiri sollunga."
)

SYSTEM_PROMPT = (
    "You are an Excel assistant working on the user's currently open Excel workbook.\n"
    "Rules:\n"
    "- When the user asks to read, analyze, check or change data, call the right tool. "
    "Never ask permission to run a read-only tool: just run it. Call tools one at a time.\n"
    "- Use ONLY the sheet and column names listed below. Do not pass sheet_name unless the user names a sheet.\n"
    "- Tool guide: duplicate rows -> remove_excel_duplicates; blank rows -> remove_excel_blank_rows; "
    "overall data quality -> validate_excel_data; "
    "total of a column for a year or per group (e.g. 2024 revenue, revenue by region) -> calculate_excel_total; "
    "overall stats of a whole column (sum/average/min/max) -> analyze_excel_column; "
    "general Excel how-to or concept questions (what is VLOOKUP, how to remove duplicates) -> ask_knowledge_base; "
    "show rows -> read_excel_data; grouped summary -> create_excel_pivot_summary; filtering -> filter_excel_data; "
    "write a plain value into a cell -> write_excel_data (data=[[value]], start_cell=cell); "
    "apply_excel_formula only for formulas that start with '='.\n"
    "- Delete and overwrite tools only preview first. The system asks the user to confirm, so never ask yourself.\n"
    "- You cannot save or open workbooks. If asked to save, tell the user to press Cmd+S in Excel.\n"
    "- If a requested column does not exist, say so and list the real columns. Never invent numbers.\n"
    "- Keep answers short (max 4 sentences). No big tables. "
    "Reply in the user's language; if the user writes Tanglish (Tamil in English letters), reply in simple Tanglish.\n"
    "- For greetings or general questions, answer directly without tools."
)

KNOWLEDGE_TOOL = {
    "type": "function",
    "function": {
        "name": "ask_knowledge_base",
        "description": (
            "Answer general Excel knowledge questions (what is VLOOKUP, how to remove duplicates, "
            "shortcuts, pivot tables, errors). Do NOT use for questions about the user's own data."
        ),
        "parameters": {
            "type": "object",
            "properties": {"question": {"type": "string", "description": "The user's Excel question"}},
            "required": ["question"],
        },
    },
}


def _to_ollama_tool(tool: Any) -> dict[str, Any]:
    schema = copy.deepcopy(tool.inputSchema)
    props = schema.get("properties", {})
    for name in HIDDEN_PARAMS:
        props.pop(name, None)
    schema["required"] = [r for r in schema.get("required", []) if r not in HIDDEN_PARAMS]
    return {
        "type": "function",
        "function": {
            "name": tool.name,
            "description": tool.description or "",
            "parameters": schema,
        },
    }


class ExcelAgent:
    def __init__(self, llm: OllamaEngine | None = None, mcp: MCPClient | None = None) -> None:
        self.llm = llm or OllamaEngine()
        self.mcp = mcp or MCPClient()
        self.tools: list[dict[str, Any]] = []
        self.tool_names: set[str] = set()
        self.history: list[dict[str, str]] = []
        self.pending: dict[str, Any] | None = None
        # Workbook context, refreshed at the start of every request.
        self.workbook: str = ""
        self.sheets: list[str] = []
        self.active_sheet: str = ""
        self.columns: list[str] = []
        self.total_rows: int = 0

    async def connect(self) -> None:
        await self.mcp.connect()
        mcp_tools = await self.mcp.get_tools()
        self.tools = [_to_ollama_tool(t) for t in mcp_tools] + [KNOWLEDGE_TOOL]
        self.tool_names = {t.name for t in mcp_tools} | {"ask_knowledge_base"}

    async def close(self) -> None:
        await self.mcp.disconnect()

    async def run(self, user_text: str) -> dict[str, Any]:
        """Run one user request. Returns reply, steps, needs_confirmation."""
        if not self.tool_names:
            return {"success": False, "error": "Agent connect() call pannala."}
        self.pending = None
        await self._load_context()
        knowledge = await asyncio.to_thread(self._knowledge_for, user_text)
        messages: list[dict[str, Any]] = (
            [{"role": "system", "content": self._build_system_prompt(knowledge)}]
            + self.history[-6:]
            + [{"role": "user", "content": user_text}]
        )
        steps: list[dict[str, Any]] = []
        nudged = False

        for _ in range(MAX_STEPS):
            res = await asyncio.to_thread(self.llm.chat, messages, self.tools)
            if not res["success"]:
                return {"success": False, "error": res["error"], "steps": steps}

            calls = res["tool_calls"]
            if not calls:
                if not steps and not nudged and self._should_nudge(user_text, res["content"]):
                    nudged = True
                    messages.append({"role": "user", "content": (
                        f'Do not answer from memory or guess. Use the matching tool now for this request: "{user_text}". '
                        "If it truly cannot be done with the tools, say so briefly."
                    )})
                    continue
                if not steps and nudged and self._looks_fabricated(res["content"]):
                    return self._finish(user_text, FABRICATION_REPLY, steps)
                return self._finish(user_text, res["content"], steps)

            messages.append({
                "role": "assistant",
                "content": res["content"],
                "tool_calls": [
                    {"function": {"name": c["name"], "arguments": c["arguments"]}} for c in calls
                ],
            })
            for call in calls:  # always sequential: never parallel Excel calls
                name = call["name"]
                if name not in self.tool_names:
                    args = dict(call["arguments"])
                    result: dict[str, Any] = {"success": False, "error": f"Unknown tool: {name}"}
                elif name == "ask_knowledge_base":
                    args = dict(call["arguments"])
                    kb = await asyncio.to_thread(answer_question, args.get("question") or user_text)
                    result = {
                        "success": kb["success"],
                        "answer": kb.get("answer", kb.get("error")),
                        "sources": kb.get("sources", []),
                    }
                else:
                    args = self._sanitize_args(name, call["arguments"], user_text)
                    missing = self._missing_column(args)
                    if missing:
                        reply = f"'{missing}' column illa. Available columns: {', '.join(self.columns)}."
                        return self._finish(user_text, reply, steps)
                    result = await self.mcp.call_tool(name, args)
                steps.append({"tool": name, "arguments": args, "result": result})
                messages.append({
                    "role": "tool",
                    "tool_name": name,
                    "content": self._compact(result),
                })
                if self._needs_confirmation(name, result):
                    self.pending = {"tool": name, "arguments": args}
                    note = result.get("message") or result.get("error") or "Confirmation venum."
                    reply = f"{note} Confirm pannanuma?"
                    self._remember(user_text, reply)
                    return {
                        "success": True,
                        "reply": reply,
                        "steps": steps,
                        "needs_confirmation": True,
                    }

        return self._finish(user_text, "Max steps mudinjiduchu. Request-ah simple-ah sollunga.", steps)

    async def confirm_pending(self) -> dict[str, Any]:
        """Execute the pending destructive action (after user said yes)."""
        if not self.pending:
            return {"success": False, "error": "Confirm panna edhuvum pending illa."}
        name, args = self.pending["tool"], dict(self.pending["arguments"])
        self.pending = None
        args["confirm" if name.startswith("remove_") else "overwrite"] = True
        result = await self.mcp.call_tool(name, args)
        reply = result.get("message") or result.get("error") or "Mudinjidichu."
        if result.get("deleted") and result.get("removed") is not None:
            reply = f"{result['removed']} rows delete pannitten. Backup: {result.get('backup')}"
        self._remember("yes, confirmed", reply)
        return {
            "success": bool(result.get("success")),
            "reply": reply,
            "steps": [{"tool": name, "arguments": args, "result": result}],
            "needs_confirmation": False,
        }

    def cancel_pending(self) -> None:
        if self.pending:
            self._remember("no, cancel", "Cancelled. Nothing was changed.")
        self.pending = None

    # ---------------- context ----------------
    async def _load_context(self) -> None:
        """Fetch real sheet and column names so the LLM never has to guess them.

        MCP connection errors are NOT caught here: the Orchestrator reconnects on them.
        """
        self.workbook, self.sheets, self.active_sheet, self.columns, self.total_rows = "", [], "", [], 0
        names = await self.mcp.call_tool("get_excel_sheet_names", {})
        if not names.get("success"):
            return
        self.workbook = names.get("workbook", "")
        self.sheets = list(names.get("sheets", []))
        self.active_sheet = names.get("active_sheet", "")
        head = await self.mcp.call_tool("read_excel_data", {"max_rows": 1})
        if head.get("success"):
            self.columns = list(head.get("columns", []))
            self.total_rows = int(head.get("total_rows", 0))

    def _build_system_prompt(self, knowledge: str = "") -> str:
        if not self.sheets:
            return SYSTEM_PROMPT + "\nNo workbook data could be read right now."
        context = (
            f"\nOpen workbook: {self.workbook}. Sheets: {self.sheets}. Active sheet: {self.active_sheet} "
            f"({self.total_rows} data rows)."
        )
        if self.columns:
            letters = ", ".join(
                f"{self._col_letter(i)}={name}" for i, name in enumerate(self.columns)
            )
            last_row = self.total_rows + 1
            free_col = self._col_letter(len(self.columns) + 1)
            context += (
                f" Columns (use these exact names): {self.columns}. Column letters: {letters}. "
                f"Data is in rows 2 to {last_row}. Empty cells for results start at {free_col}1."
            )
        if knowledge:
            context += (
                "\n\nKNOWLEDGE (from the Excel knowledge base):\n" + knowledge +
                "\n\nIf the user wants a calculation or lookup done IN the sheet, pick the right formula "
                "from the knowledge above, replace the example ranges with the real column letters and "
                "rows, and call apply_excel_formula to put it in an empty cell."
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
        """Top knowledge-base chunks for this request (empty string if retrieval fails)."""
        try:
            chunks = retrieve_context(user_text, top_k=KNOWLEDGE_TOP_K)
        except Exception:
            return ""
        return "\n---\n".join(c["content"][:KNOWLEDGE_CHUNK_CHARS] for c in chunks)

    # ---------------- argument sanitizing ----------------
    @staticmethod
    def _match(value: str, options: list[str]) -> str | None:
        if value in options:
            return value
        lowered = value.strip().lower()
        for option in options:
            if str(option).lower() == lowered:
                return option
        return None

    def _sanitize_args(self, tool: str, raw: dict[str, Any], user_text: str) -> dict[str, Any]:
        """Remove hidden params, drop invented sheet names, fix column-name casing, cap rows."""
        args = {k: v for k, v in raw.items() if k not in HIDDEN_PARAMS}

        sheet = args.get("sheet_name")
        if isinstance(sheet, str) and self.sheets:
            real = self._match(sheet, self.sheets)
            if real:
                args["sheet_name"] = real
            elif sheet.strip().lower() not in user_text.lower():
                args.pop("sheet_name")  # the LLM invented it, the user never said it

        if self.columns:
            for key in COLUMN_PARAMS:
                value = args.get(key)
                if isinstance(value, str):
                    args[key] = self._match(value, self.columns) or value
            if isinstance(args.get("columns"), list):
                args["columns"] = [
                    (self._match(c, self.columns) or c) if isinstance(c, str) else c
                    for c in args["columns"]
                ]

        if tool == "read_excel_data":
            try:
                rows = int(args.get("max_rows", MAX_ROWS_FOR_LLM))
            except (TypeError, ValueError):
                rows = MAX_ROWS_FOR_LLM
            args["max_rows"] = max(1, min(rows, MAX_ROWS_FOR_LLM))
        return args

    def _missing_column(self, args: dict[str, Any]) -> str | None:
        """Return a requested column name that does not exist on the active sheet, else None."""
        if not self.columns or args.get("sheet_name", self.active_sheet) != self.active_sheet:
            return None
        for key in COLUMN_PARAMS:
            value = args.get(key)
            if isinstance(value, str) and value not in self.columns:
                return value
        extra = args.get("columns")
        if isinstance(extra, list):
            for value in extra:
                if isinstance(value, str) and value not in self.columns:
                    return value
        return None

    # ---------------- helpers ----------------
    @staticmethod
    def _should_nudge(user_text: str, content: str) -> bool:
        """True when the LLM answered an Excel request without calling any tool."""
        words = user_text.strip().lower().replace(",", " ").replace("!", " ").replace("?", " ").split()
        if words and words[0] in GREETING_WORDS and len(words) <= 5:
            return False  # small talk
        reply = (content or "").lower()
        return not any(marker in reply for marker in ("cmd+s", "cannot", "can't", "unable to"))

    @staticmethod
    def _looks_fabricated(content: str) -> bool:
        return any(ch.isdigit() for ch in content or "")

    @staticmethod
    def _needs_confirmation(name: str, result: dict[str, Any]) -> bool:
        if name.startswith("remove_"):
            return bool(result.get("success")) and not result.get("deleted") and (result.get("found") or 0) > 0
        return not result.get("success") and "overwrite=True" in str(result.get("error", ""))

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
        return text if len(text) <= MAX_RESULT_CHARS else text[:MAX_RESULT_CHARS] + "...(truncated)"

    def _remember(self, user_text: str, reply: str) -> None:
        self.history += [{"role": "user", "content": user_text}, {"role": "assistant", "content": reply}]

    def _finish(self, user_text: str, reply: str, steps: list[dict[str, Any]]) -> dict[str, Any]:
        if not (reply or "").strip() and steps:
            last = steps[-1]["result"]
            reply = str(last.get("message") or last.get("error") or self._compact(last))
        self._remember(user_text, reply)
        return {"success": True, "reply": reply, "steps": steps, "needs_confirmation": False}


async def _smoke_test() -> None:
    agent = ExcelAgent()
    await agent.connect()
    try:
        for prompt in ["Hello", "intha workbook-la enna sheets irukku?"]:
            out = await agent.run(prompt)
            print(f"\nYOU: {prompt}")
            print("AGENT:", out.get("reply") or out.get("error"))
    finally:
        await agent.close()


if __name__ == "__main__":
    asyncio.run(_smoke_test())