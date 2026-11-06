"""Task planner for the Excel AI Agent.

Converts a natural-language request into a structured execution plan.
Does not execute Excel tools. Execution remains the Agent/Executor's job.
"""

import json
from typing import Any


MAX_PLAN_STEPS = 8
MAX_TOOL_DESCRIPTION_CHARS = 500


PLANNER_SYSTEM_PROMPT = """
You are the task planner for a local Excel AI Agent.

Your job is to convert the user's request into a clear, ordered execution plan.

RULES:
1. Understand the user's actual intent, including informal English and Tanglish.
2. Decide whether the request needs an Excel action or is only a question/conversation.
3. Break complex Excel requests into small, logical steps.
4. Respect dependencies: a later step may depend on an earlier step's result.
5. Use ONLY tools from the supplied available tool list.
6. Never invent tool names or capabilities.
7. Do not invent worksheet names, column names, cell references or values.
8. If essential information is missing, mark needs_clarification as true.
9. Do not create unnecessary steps for simple requests.
10. Do not execute anything. Only create a plan.
11. You only help with the user's workbook. Decide scope using WORKBOOK CONTEXT.
12. If the request is about the workbook (its sheets, columns, cells, values, or an edit on them), it is IN scope. If it mentions a sheet or column name from WORKBOOK CONTEXT, it is IN scope, even when phrased as a question ("what is sum of units").
13. If the request is general Excel theory, a definition, or any other topic, set out_of_scope to true. Do not plan steps.
14. In scope but the column or sheet is unclear: set needs_clarification to true and ask which one.

Scope examples:
- "what is excel" -> out_of_scope true
- "what is vlookup" -> out_of_scope true
- "weather enna" -> out_of_scope true
- "what is sum of Units" (Units is a column) -> out_of_scope false, requires_action true
- "total salary enna" (Salary is a column) -> out_of_scope false, requires_action true

Return ONLY valid JSON. Do not use Markdown or code fences.

Required JSON structure:
{
  "intent": "short description of the user's goal",
  "out_of_scope": false,
  "requires_action": true,
  "needs_clarification": false,
  "clarification_question": "",
  "steps": [
    {
      "step": 1,
      "task": "What this step must accomplish",
      "tool": "exact_available_tool_name_or_null",
      "success_criteria": "How completion should be checked"
    }
  ]
}

For a question about workbook data (totals, averages, counts, values):
- Set requires_action to true. A read or analysis tool must answer it.

For greetings or small talk about the agent itself:
- Set requires_action to false and use an empty steps list.

For a request requiring Excel operations:
- Set requires_action to true.
- Create ordered steps.
- Select a tool only if its purpose clearly matches the step.
- Use null when the correct tool cannot be identified confidently.

If clarification is required:
- Set needs_clarification to true.
- Ask one concise question.
- Do not guess missing essential information.
"""


class TaskPlanner:
    """Creates and validates structured plans using the existing Ollama engine."""

    def __init__(self, llm: Any) -> None:
        self.llm = llm

    def create_plan(
        self,
        user_text: str,
        tools: list[dict[str, Any]],
        workbook_context: str = "",
    ) -> dict[str, Any]:
        """Return a validated plan or a safe fallback plan."""

        if not isinstance(user_text, str) or not user_text.strip():
            return {
                "success": False,
                "error": "Empty user request.",
                "plan": self._fallback_plan(user_text or ""),
            }

        tool_catalog = self._build_tool_catalog(tools)

        messages = [
            {
                "role": "system",
                "content": PLANNER_SYSTEM_PROMPT,
            },
            {
                "role": "user",
                "content": (
                    f"USER REQUEST:\n{user_text.strip()}\n\n"
                    f"WORKBOOK CONTEXT:\n"
                    f"{workbook_context or 'Not available'}\n\n"
                    f"AVAILABLE TOOLS:\n{tool_catalog}\n\n"
                    "Create the execution plan using the required JSON structure."
                ),
            },
        ]

        try:
            response = self.llm.chat(messages)

            if not response or not response.get("success"):
                error = (response or {}).get(
                    "error", "Planner LLM request failed."
                )
                return {
                    "success": False,
                    "error": str(error),
                    "plan": self._fallback_plan(user_text),
                }

            content = response.get("content") or ""
            raw_plan = self._extract_json(content)
            plan = self._validate_plan(raw_plan, tools)

            return {
                "success": True,
                "plan": plan,
            }

        except (ValueError, TypeError, KeyError, json.JSONDecodeError) as exc:
            return {
                "success": False,
                "error": f"Invalid planner response: {exc}",
                "plan": self._fallback_plan(user_text),
            }
        except Exception as exc:
            return {
                "success": False,
                "error": f"Planner error: {type(exc).__name__}: {exc}",
                "plan": self._fallback_plan(user_text),
            }

    @staticmethod
    def _build_tool_catalog(tools: list[dict[str, Any]]) -> str:
        """Give the planner concise descriptions of the available tools."""

        entries = []

        for item in tools:
            function = item.get("function", {})
            name = function.get("name")

            if not name:
                continue

            description = function.get("description", "")
            description = str(description)[:MAX_TOOL_DESCRIPTION_CHARS]

            entries.append(f"- {name}: {description}")

        return "\n".join(entries) if entries else "No tools available."

    @staticmethod
    def _extract_json(content: str) -> dict[str, Any]:
        """Extract a JSON object, including when the model adds code fences."""

        if not isinstance(content, str) or not content.strip():
            raise ValueError("The model returned an empty planning response.")

        text = content.strip()

        if text.startswith("```"):
            lines = text.splitlines()
            if lines and lines[0].startswith("```"):
                lines = lines[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            text = "\n".join(lines).strip()

        decoder = json.JSONDecoder()
        start = text.find("{")

        if start == -1:
            raise ValueError("No JSON object found in planner response.")

        result, _ = decoder.raw_decode(text[start:])

        if not isinstance(result, dict):
            raise ValueError("Planner response must be a JSON object.")

        return result

    @staticmethod
    def _validate_plan(
        raw: dict[str, Any],
        tools: list[dict[str, Any]],
    ) -> dict[str, Any]:
        """Validate the plan and reject tool names not present in MCP registry."""

        available_names = {
            item.get("function", {}).get("name")
            for item in tools
            if item.get("function", {}).get("name")
        }

        intent = raw.get("intent", "")
        if not isinstance(intent, str):
            intent = ""

        out_of_scope = raw.get("out_of_scope", False)
        if not isinstance(out_of_scope, bool):
            out_of_scope = False

        requires_action = raw.get("requires_action", False)
        needs_clarification = raw.get("needs_clarification", False)

        if not isinstance(requires_action, bool):
            requires_action = False

        if not isinstance(needs_clarification, bool):
            needs_clarification = False

        question = raw.get("clarification_question", "")
        if not isinstance(question, str):
            question = ""

        raw_steps = raw.get("steps", [])
        if not isinstance(raw_steps, list):
            raw_steps = []

        steps = []

        for item in raw_steps[:MAX_PLAN_STEPS]:
            if not isinstance(item, dict):
                continue

            task = item.get("task", "")
            if not isinstance(task, str) or not task.strip():
                continue

            tool = item.get("tool")
            if not isinstance(tool, str) or tool not in available_names:
                tool = None

            criteria = item.get("success_criteria", "")
            if not isinstance(criteria, str):
                criteria = ""

            steps.append({
                "step": len(steps) + 1,
                "task": task.strip(),
                "tool": tool,
                "success_criteria": criteria.strip(),
            })

        if out_of_scope:
            requires_action = False
            needs_clarification = False
            steps = []

        if needs_clarification:
            requires_action = False
            steps = []

        if not requires_action:
            steps = []

        return {
            "intent": intent.strip(),
            "out_of_scope": out_of_scope,
            "requires_action": requires_action,
            "needs_clarification": needs_clarification,
            "clarification_question": question.strip(),
            "steps": steps,
        }

    @staticmethod
    def _fallback_plan(user_text: str) -> dict[str, Any]:
        """Preserve the request for the existing Agent flow if planning fails."""

        return {
            "intent": user_text.strip(),
            "out_of_scope": False,
            "requires_action": True,
            "needs_clarification": False,
            "clarification_question": "",
            "steps": [
                {
                    "step": 1,
                    "task": user_text.strip(),
                    "tool": None,
                    "success_criteria": (
                        "Use the existing Agent tool-calling flow to determine "
                        "whether the request can be completed."
                    ),
                }
            ],
        }