"""Orchestrator: routing, workflow coordination, error management.

Sits between the Streamlit frontend and the Agent (and Dashboard workflows).
One background thread owns ONE event loop; the MCP connection lives on that
loop for the whole process, so Streamlit reruns never create or close it.

Changes:
- No full-request re-run on connection errors (agent handles it per tool call)
- Dashboard keyword hijack fixed (questions go to the agent / knowledge)
- Stop support (request_stop)
- Ollama health check before first request
- Structured dict return (process_user_request_full) + old str API kept
- Out-of-scope requests get one fixed reply (OUT_OF_SCOPE_REPLY)
"""
import asyncio
import concurrent.futures
import threading
from typing import Any

from backend.agent import ExcelAgent
from dashboard.auto_dashboard import auto_dashboard

REQUEST_TIMEOUT = 300

# User-typed confirmation words (input matching, so Tamil words stay).
YES_WORDS = {
    "aam", "ஆம்", "yes", "y", "ok", "okay", "confirm", "sari", "seri",
    "haan", "pannu", "delete pannu", "go ahead",
}
NO_WORDS = {
    "illa", "illai", "இல்லை", "vendaam", "venam", "வேண்டாம்",
    "no", "n", "cancel", "stop",
}

DASHBOARD_NEEDS_EDIT = (
    "Creating a dashboard adds a new sheet to the workbook. "
    "Please switch to Edit mode."
)
CHAT_PENDING_CANCELLED = (
    "Chat mode is read-only, so I cancelled the pending change. "
    "Switch to Edit mode if you want to apply it."
)
OUT_OF_SCOPE_REPLY = "Please ask me anything about your dataset."

DASHBOARD_ACTION_WORDS = (
    "create", "build", "make", "generate", "auto",
    "pannu", "panni", "podu", "seyyu", "venum",
)

# Questions like "How to make a dashboard?" must not create a dashboard.
QUESTION_STARTS = (
    "what", "how", "why", "explain", "can you explain", "tell me",
    "enna", "eppadi", "epdi", "yen", "ethukku", "edhuku",
)
QUESTION_PHRASES = (
    "how to", "how do", "how can", "what is", "what are",
    "meaning of", "explain", "eppadi", "epdi", "enna meaning",
)


def _normalize(text: str) -> str:
    return text.strip().lower().strip(".!?, ")


def _is_question(text: str) -> bool:
    lowered = _normalize(text)
    if lowered.startswith(QUESTION_STARTS):
        return True
    return any(phrase in f" {lowered} " for phrase in QUESTION_PHRASES)


def _dashboard_intent(text: str) -> str | None:
    """'auto' / 'guided' when the user really asks to build a dashboard."""
    lowered = text.lower()
    if "dashboard" not in lowered:
        return None
    if _is_question(text):
        return None
    if not any(word in lowered for word in DASHBOARD_ACTION_WORDS):
        return None
    return "guided" if "guided" in lowered else "auto"


def _failure(message: str, **extra: Any) -> dict[str, Any]:
    return {
        "success": False,
        "error": message,
        "reply": message,
        "steps": [],
        "needs_confirmation": False,
        **extra,
    }


class Orchestrator:

    def __init__(self) -> None:
        self.agent: ExcelAgent | None = None
        self._llm_checked = False

    # ---------------- LIFECYCLE ----------------

    async def initialize(self) -> None:
        """Create the agent and connect the MCP client once."""
        if self.agent is not None:
            return
        agent = ExcelAgent()
        await agent.connect()
        self.agent = agent

    async def close(self) -> None:
        agent, self.agent = self.agent, None
        if agent is not None:
            await agent.close()

    def request_stop(self) -> None:
        """Stop button. Does not take the lock; thread-safe."""
        agent = self.agent
        if agent is not None:
            agent.request_stop()

    async def _ensure_llm(self) -> str | None:
        """Ollama health check on the first request. Returns an error message or None."""
        if self._llm_checked or self.agent is None:
            return None
        check = await asyncio.to_thread(self.agent.llm.health_check)
        if not check.get("success"):
            return check.get("error", "Ollama check failed.")
        self._llm_checked = True
        return None

    # ---------------- ROUTING ----------------

    async def route_request(
        self,
        user_text: str,
        mode: str = "edit",
    ) -> dict[str, Any]:
        """Route one user message. Always returns a dict."""
        try:
            await self.initialize()
            assert self.agent is not None

            llm_error = await self._ensure_llm()
            if llm_error:
                return _failure(llm_error)

            word = _normalize(user_text)

            # In chat mode, cancel any pending change.
            if mode == "chat" and self.agent.pending:
                self.agent.cancel_pending()
                if word in YES_WORDS or word in NO_WORDS:
                    return {
                        "success": True,
                        "reply": CHAT_PENDING_CANCELLED,
                        "steps": [],
                        "needs_confirmation": False,
                    }

            # Pending confirmation answer
            if self.agent.pending:
                if word in YES_WORDS:
                    return await self.agent.confirm_pending()
                if word in NO_WORDS:
                    self.agent.cancel_pending()
                    return {
                        "success": True,
                        "reply": "Okay, cancelled. Nothing was changed.",
                        "steps": [],
                        "needs_confirmation": False,
                    }

            # Dashboard (only real create requests, never questions)
            intent = _dashboard_intent(user_text)
            if intent and mode == "chat":
                return {
                    "success": True,
                    "reply": DASHBOARD_NEEDS_EDIT,
                    "steps": [],
                    "needs_confirmation": False,
                }
            if intent == "auto":
                return await asyncio.to_thread(self._run_auto_dashboard)
            if intent == "guided":
                return {
                    "success": True,
                    "reply": (
                        "Guided Dashboard is not ready yet. "
                        "Say 'create auto dashboard' instead."
                    ),
                    "steps": [],
                    "needs_confirmation": False,
                }

            # Normal agent flow. Reconnects are handled per tool call inside
            # the agent, so the full request is never re-run (prevents
            # duplicate writes).
            result = await self.agent.run(user_text, mode)

            # The planner decided the request is not about the workbook.
            if isinstance(result, dict) and result.get("out_of_scope"):
                return {
                    "success": True,
                    "reply": OUT_OF_SCOPE_REPLY,
                    "steps": [],
                    "needs_confirmation": False,
                }

            return result

        except Exception as exc:
            return _failure(
                f"Orchestrator Error: {type(exc).__name__}: {exc}",
                reply="Sorry, an unexpected error occurred while processing your request.",
            )

    @staticmethod
    def _run_auto_dashboard() -> dict[str, Any]:
        res = auto_dashboard()
        if not res.get("success"):
            return {
                "success": True,
                "reply": f"Dashboard could not be created: {res.get('error')}",
                "steps": [],
                "needs_confirmation": False,
            }
        notes = f" Notes: {'; '.join(res['notes'])}." if res.get("notes") else ""
        reply = (
            f"Created the '{res['sheet']}' sheet ({res['title']}). "
            f"KPIs: {', '.join(res['kpis'])}. "
            f"Charts: {len(res['charts'])}.{notes} "
            "Open the dashboard sheet in Excel to view it."
        )
        return {
            "success": True,
            "reply": reply,
            "steps": [],
            "needs_confirmation": False,
        }


# ============================================================
# EVENT LOOP THREAD
# ============================================================

class _LoopThread:
    """A daemon thread running one event loop forever."""

    def __init__(self) -> None:
        self.loop = asyncio.new_event_loop()
        self.thread = threading.Thread(
            target=self._run, name="mcp-loop", daemon=True
        )
        self.thread.start()

    def _run(self) -> None:
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def submit(self, coro, timeout: float):
        future = asyncio.run_coroutine_threadsafe(coro, self.loop)
        try:
            return future.result(timeout)
        except concurrent.futures.TimeoutError:
            future.cancel()
            raise


_global_orchestrator = Orchestrator()
_loop_thread: _LoopThread | None = None
_state_lock = threading.Lock()      # protects loop creation
_request_lock = threading.Lock()    # one request at a time (Excel handles one command at a time)


def _get_loop_thread() -> _LoopThread:
    global _loop_thread
    with _state_lock:
        if _loop_thread is None:
            _loop_thread = _LoopThread()
        return _loop_thread


# ============================================================
# PUBLIC API (Streamlit / desktop frontend)
# ============================================================

def process_user_request_full(user_text: str, mode: str = "edit") -> dict[str, Any]:
    """Sync entry point. Returns the full dict:
    {success, reply, steps, needs_confirmation, error?}
    """
    text = (user_text or "").strip()
    mode = (mode or "edit").lower()
    if mode not in ("chat", "edit"):
        mode = "chat"   # unknown mode -> safest (read-only)

    if not text:
        return {
            "success": True,
            "reply": "Please type a request.",
            "steps": [],
            "needs_confirmation": False,
        }

    with _request_lock:
        try:
            return _get_loop_thread().submit(
                _global_orchestrator.route_request(text, mode),
                REQUEST_TIMEOUT,
            )
        except concurrent.futures.TimeoutError:
            # Stop the LLM/tool still running in the background.
            _global_orchestrator.request_stop()
            return _failure("The request took too long (timeout).")
        except Exception as exc:
            return _failure(f"{type(exc).__name__}: {exc}")


def process_user_request(user_text: str, mode: str = "edit") -> str:
    """Old sync API (returns reply text only). Keeps the existing frontend working."""
    res = process_user_request_full(user_text, mode)
    if not res.get("success"):
        return f"Error: {res.get('error', 'Unknown error')}"
    return res.get("reply", "Done.")


def stop_current_request() -> None:
    """The UI Stop button should call this."""
    _global_orchestrator.request_stop()