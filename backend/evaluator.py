
"""Evaluator for the Excel AI Agent.

Checks tool execution results against each plan step's success criteria.
This module does not execute tools or modify Excel.
"""

from typing import Any


class TaskEvaluator:
    def evaluate_step(
        self,
        step: dict[str, Any],
        result: dict[str, Any],
    ) -> dict[str, Any]:
        """Evaluate whether a tool call completed successfully."""

        if not isinstance(step, dict):
            return self._result(False, "Invalid plan step.")

        if not isinstance(result, dict):
            return self._result(False, "Invalid tool result.")

        if not result.get("success", False):
            return self._result(
                False,
                str(result.get("error", "Tool execution failed.")),
            )

        criteria = step.get("success_criteria", "")

        return self._result(
            True,
            "Tool reported successful execution. "
            "The success criteria have not yet been independently verified.",
            criteria=criteria,
            verified=False,
        )

    def evaluate_plan(
        self,
        plan: dict[str, Any],
        execution: dict[str, Any],
    ) -> dict[str, Any]:
        """Evaluate the overall execution outcome."""

        if not isinstance(execution, dict):
            return self._result(False, "Invalid execution report.")

        if execution.get("status") != "completed":
            return self._result(
                False,
                str(execution.get("error", "Plan did not complete.")),
                status=execution.get("status", "failed"),
            )

        steps = execution.get("steps", [])

        if not isinstance(steps, list) or not steps:
            return self._result(
                False,
                "No completed steps were available for evaluation.",
            )

        failed_steps = [
            item for item in steps
            if not isinstance(item, dict)
            or item.get("status") != "completed"
        ]

        if failed_steps:
            return self._result(
                False,
                "One or more execution steps failed.",
                failed_steps=failed_steps,
            )

        return self._result(
            True,
            "All planned tool calls reported successful execution. "
            "Independent workbook verification is still required.",
            completed_steps=len(steps),
            verified=False,
        )

    @staticmethod
    def _result(
        success: bool,
        message: str,
        **extra: Any,
    ) -> dict[str, Any]:
        return {
            "success": success,
            "message": message,
            **extra,
        }
