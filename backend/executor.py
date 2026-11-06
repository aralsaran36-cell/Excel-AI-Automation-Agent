"""Executor for the Excel AI Agent.

Executes structured plans sequentially through the existing MCP client.
Does not make planning decisions or perform LLM reasoning itself.
If a step has no arguments, an injected arg_builder (LLM) fills them,
using guidance for the selected tool only.
"""

import asyncio
from typing import Any, Awaitable, Callable

from backend.rag import get_guidance


# Tools that need "how to use" knowledge from the KB.
NEEDS_GUIDANCE = {
    "apply_excel_formula",
    "apply_excel_formula_range",
    "add_excel_formula_column",
    "remove_excel_duplicates",
    "fill_missing_excel_values",
    "clean_excel_text",
    "apply_conditional_formatting",
    "apply_data_validation",
    "create_excel_chart",
    "create_excel_pivot_summary",
    "create_excel_native_pivot",
    "delete_excel_rows",
    "delete_excel_columns",
    "delete_worksheet",
    "clear_excel_range",
    "filter_excel_multi",
    "find_replace_excel_data",
}

RETRY_DELAYS = (1, 3)
CONNECTION_ERRORS = (ConnectionError, TimeoutError, OSError)

# arg_builder(task, tool_name, guidance, context) -> arguments dict
ArgBuilder = Callable[
    [str, str, str, dict[str, Any]],
    Awaitable[dict[str, Any]],
]


class TaskExecutor:
    def __init__(
        self,
        mcp: Any,
        tool_names: set[str],
        mode: str = "edit",
        arg_builder: ArgBuilder | None = None,
    ) -> None:
        self.mcp = mcp
        self.tool_names = set(tool_names)
        self.mode = mode
        self.arg_builder = arg_builder

    async def execute_plan(
        self,
        plan: dict[str, Any],
        context: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        """Execute a validated plan sequentially."""

        context = context or {}

        if not isinstance(plan, dict):
            return self._failure("Invalid execution plan.")

        if plan.get("needs_clarification"):
            return {
                "success": True,
                "status": "needs_clarification",
                "reply": plan.get(
                    "clarification_question",
                    "Please clarify your request.",
                ),
                "steps": [],
            }

        if not plan.get("requires_action"):
            return {
                "success": True,
                "status": "no_action_required",
                "steps": [],
            }

        planned_steps = plan.get("steps")

        if not isinstance(planned_steps, list) or not planned_steps:
            return self._failure("Plan contains no executable steps.")

        results: list[dict[str, Any]] = []

        for position, step in enumerate(planned_steps, start=1):
            if not isinstance(step, dict):
                return self._failure(
                    f"Invalid step definition at position {position}.",
                    results,
                )

            tool_name = step.get("tool")
            task = step.get("task", "")

            if not isinstance(tool_name, str) or not tool_name:
                return self._failure(
                    f"Step {position} has no valid tool selected.",
                    results,
                )

            if tool_name not in self.tool_names:
                return self._failure(
                    f"Step {position}: unknown tool '{tool_name}'.",
                    results,
                )

            arguments = step.get("arguments") or {}

            if not isinstance(arguments, dict):
                return self._failure(
                    f"Step {position}: invalid tool arguments.",
                    results,
                )

            # ---- Fill arguments (selected tool's guidance only)
            guidance_used = False

            if not arguments and self.arg_builder:
                guidance = await self._get_guidance(task, tool_name)
                guidance_used = bool(guidance)

                try:
                    arguments = await self.arg_builder(
                        task,
                        tool_name,
                        guidance,
                        context,
                    )
                except Exception as exc:
                    results.append({
                        "step": position,
                        "task": task,
                        "tool": tool_name,
                        "status": "error",
                        "arguments": {},
                        "error": f"Argument building failed: {exc}",
                    })
                    return self._failure(
                        f"Execution stopped at step {position}.",
                        results,
                    )

                if not isinstance(arguments, dict):
                    return self._failure(
                        f"Step {position}: invalid tool arguments.",
                        results,
                    )

            # ---- Call the tool (retry on connection errors)
            try:
                result = await self._call_with_retry(tool_name, arguments)

            except Exception as exc:
                results.append({
                    "step": position,
                    "task": task,
                    "tool": tool_name,
                    "status": "error",
                    "arguments": arguments,
                    "error": f"{type(exc).__name__}: {exc}",
                })
                return self._failure(
                    f"Execution stopped at step {position}.",
                    results,
                )

            result = self._normalize_result(result)
            success = bool(result.get("success"))

            results.append({
                "step": position,
                "task": task,
                "tool": tool_name,
                "status": "completed" if success else "failed",
                "arguments": arguments,
                "guidance_used": guidance_used,
                "result": result,
            })

            if not success:
                return self._failure(
                    f"Execution failed at step {position}: "
                    f"{result.get('error', 'Unknown tool error')}",
                    results,
                )

            context[f"step_{position}_result"] = result

        return {
            "success": True,
            "status": "completed",
            "steps": results,
            "context": context,
        }

    # -------------------------------------------------
    # Helpers
    # -------------------------------------------------

    async def _get_guidance(self, task: str, tool_name: str) -> str:
        """Guidance for the selected tool only. Empty if none."""
        if tool_name not in NEEDS_GUIDANCE:
            return ""
        try:
            return await asyncio.to_thread(
                get_guidance, task, [tool_name]
            )
        except Exception:
            # RAG problem must not stop the edit.
            return ""

    async def _call_with_retry(
        self,
        tool_name: str,
        arguments: dict[str, Any],
    ) -> Any:
        last_error: Exception | None = None

        for delay in (0, *RETRY_DELAYS):
            if delay:
                await asyncio.sleep(delay)
            try:
                return await self.mcp.call_tool(
                    tool_name,
                    arguments,
                    mode=self.mode,
                )
            except CONNECTION_ERRORS as exc:
                last_error = exc

        raise ConnectionError(
            f"Connection failed after retries: {last_error}"
        )

    @staticmethod
    def _normalize_result(result: Any) -> dict[str, Any]:
        if isinstance(result, dict):
            return result

        # client.call_tool returns text for guard errors
        if isinstance(result, str):
            if result.startswith("Error"):
                return {"success": False, "error": result}
            return {"success": True, "output": result}

        return {
            "success": False,
            "error": "MCP tool returned an invalid result.",
        }

    @staticmethod
    def _failure(
        message: str,
        results: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        return {
            "success": False,
            "status": "failed",
            "error": message,
            "steps": results or [],
        }