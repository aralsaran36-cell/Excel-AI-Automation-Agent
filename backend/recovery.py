
"""Error recovery decisions for the Excel AI Agent."""

from typing import Any


class RecoveryManager:
    def __init__(self, max_retries: int = 1):
        self.max_retries = max(0, max_retries)

    def analyze_error(
        self,
        error: Any,
        retry_count: int = 0,
    ) -> dict[str, Any]:
        """Decide whether a failed operation can be retried."""

        message = str(error).strip()
        normalized = message.lower()

        if not message:
            return self._decision(
                "stop",
                "Unknown error. Execution stopped.",
            )

        if any(word in normalized for word in (
            "permission denied",
            "access denied",
            "protected sheet",
            "read-only",
        )):
            return self._decision(
                "ask_user",
                "The workbook may not allow this operation. "
                "Check permissions or sheet protection.",
            )

        if any(word in normalized for word in (
            "column not found",
            "worksheet not found",
            "sheet not found",
            "missing required",
            "invalid argument",
        )):
            return self._decision(
                "ask_user",
                "Required workbook information or tool arguments "
                "may be incorrect. Clarification is needed.",
            )

        if any(word in normalized for word in (
            "timeout",
            "temporarily unavailable",
            "connection error",
            "connection refused",
        )):
            if retry_count < self.max_retries:
                return self._decision(
                    "retry",
                    "Temporary connection or timeout error.",
                )

        return self._decision(
            "stop",
            "The error cannot safely be recovered automatically.",
        )

    @staticmethod
    def _decision(action: str, reason: str) -> dict[str, Any]:
        return {
            "action": action,
            "reason": reason,
        }
