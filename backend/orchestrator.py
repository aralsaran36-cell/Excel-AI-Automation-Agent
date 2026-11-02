"""Orchestrator: routing, workflow coordination, error management.

Sits between the Streamlit frontend and the Agent (and later the Dashboard workflows).
One background thread owns ONE event loop; the MCP connection lives on that loop for
the whole process, so Streamlit reruns never create or close it.
"""
import asyncio
import concurrent.futures
import threading
from typing import Any

from backend.agent import ExcelAgent
from dashboard.auto_dashboard import auto_dashboard
from mcp_layer.client import MCPConnectionError

REQUEST_TIMEOUT = 300
YES_WORDS = {"aam", "ஆம்", "yes", "y", "ok", "okay", "confirm", "sari", "seri", "haan",
             "pannu", "delete pannu", "go ahead"}
NO_WORDS = {"illa", "illai", "இல்லை", "vendaam", "venam", "வேண்டாம்", "no", "n", "cancel", "stop"}


DASHBOARD_ACTION_WORDS = ("create", "build", "make", "generate", "auto", "pannu", "panni", "podu", "seyyu", "venum")


def _dashboard_intent(text: str) -> str | None:
    """'auto' / 'guided' when the user asks to build a dashboard, else None."""
    lowered = text.lower()
    if "dashboard" not in lowered or not any(w in lowered for w in DASHBOARD_ACTION_WORDS):
        return None
    return "guided" if "guided" in lowered else "auto"


def _normalize(text: str) -> str:
    return text.strip().lower().strip(".!?, ")


def _is_connection_error(exc: BaseException) -> bool:
    if isinstance(exc, (MCPConnectionError, ConnectionError)):
        return True
    name = type(exc).__name__
    return any(word in name for word in ("Closed", "BrokenResource", "EndOfStream"))


class Orchestrator:
    def __init__(self) -> None:
        self.agent: ExcelAgent | None = None

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

    async def route_request(self, user_text: str) -> dict[str, Any]:
        """Route one user message. Pending delete/overwrite confirmations are answered here."""
        try:
            await self.initialize()
            word = _normalize(user_text)
            if self.agent.pending:
                if word in YES_WORDS:
                    return await self._with_reconnect(lambda agent: agent.confirm_pending())
                if word in NO_WORDS:
                    self.agent.cancel_pending()
                    return {"success": True, "reply": "Seri, cancel pannitten. Edhuvum maaralai."}
            mode = _dashboard_intent(user_text)
            if mode == "auto":
                return await asyncio.to_thread(self._run_auto_dashboard)
            if mode == "guided":
                return {"success": True, "reply": "Guided Dashboard innum ready illa. 'create auto dashboard' nu sollunga."}
            return await self._with_reconnect(lambda agent: agent.run(user_text))
        except Exception as exc:
            return {
                "success": False,
                "error": f"Orchestrator Error: {type(exc).__name__}: {exc}",
                "reply": "Sorry, an unexpected error occurred while processing your request.",
            }

    @staticmethod
    def _run_auto_dashboard() -> dict[str, Any]:
        res = auto_dashboard()
        if not res.get("success"):
            return {"success": True, "reply": f"Dashboard create aagala: {res.get('error')}"}
        notes = f" Notes: {'; '.join(res['notes'])}." if res.get("notes") else ""
        reply = (f"'{res['sheet']}' sheet create pannitten ({res['title']}). "
                 f"KPIs: {', '.join(res['kpis'])}. Charts: {len(res['charts'])}.{notes} "
                 "Excel-la Dashboard sheet-a paarunga.")
        return {"success": True, "reply": reply}

    async def _with_reconnect(self, action) -> dict[str, Any]:
        """Run action(agent); if the MCP connection died, reconnect once and retry."""
        try:
            return await action(self.agent)
        except Exception as exc:
            if not _is_connection_error(exc):
                raise
            await self.close()
            await self.initialize()
            return await action(self.agent)


class _LoopThread:
    """A daemon thread running one event loop forever."""

    def __init__(self) -> None:
        self.loop = asyncio.new_event_loop()
        self.thread = threading.Thread(target=self._run, name="mcp-loop", daemon=True)
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
_state_lock = threading.Lock()    # protects loop creation
_request_lock = threading.Lock()  # one request at a time: Excel handles one command at a time


def _get_loop_thread() -> _LoopThread:
    global _loop_thread
    with _state_lock:
        if _loop_thread is None:
            _loop_thread = _LoopThread()
        return _loop_thread


def process_user_request(user_text: str) -> str:
    """Sync entry point for Streamlit. Returns the reply text."""
    text = (user_text or "").strip()
    if not text:
        return "Edhavadhu type pannunga."
    with _request_lock:
        try:
            res = _get_loop_thread().submit(_global_orchestrator.route_request(text), REQUEST_TIMEOUT)
        except concurrent.futures.TimeoutError:
            return "Error: Request romba neram aachu (timeout)."
        except Exception as exc:
            return f"Error: {type(exc).__name__}: {exc}"
    if not res.get("success"):
        return f"Error: {res.get('error', 'Unknown error')}"
    return res.get("reply", "Done.")