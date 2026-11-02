"""Step 1-2 integration test: ALL 17 tools called through the MCP server.

Reuses the checks in test_all_tools.py, but every tool call goes
test -> MCP client -> mcp_layer.server (separate process) -> excel_tools -> Excel.

Run (venv active, dataset workbook open in Excel, Excel-la click pannaadhinga):
    python test_mcp_all_tools.py                 # all 17
    python test_mcp_all_tools.py chart pivot     # only matching tools
"""
import asyncio
import json
import os
import sys
import threading
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

import inspect

import test_all_tools as T

SERVER_MODULE = "mcp_layer.server"
CALL_TIMEOUT = 120


class McpBridge:
    """Keeps ONE MCP session alive on a background event loop; sync code can call tools."""

    def __init__(self) -> None:
        self.loop = asyncio.new_event_loop()
        self.thread = threading.Thread(target=self.loop.run_forever, daemon=True)
        self.session: ClientSession | None = None
        self.error: str | None = None
        self._ready = threading.Event()
        self._stop: asyncio.Event | None = None

    async def _serve(self) -> None:
        self._stop = asyncio.Event()
        params = StdioServerParameters(
            command=sys.executable, args=["-m", SERVER_MODULE],
            env=dict(os.environ), cwd=str(ROOT),
        )
        try:
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    self.session = session
                    self._ready.set()
                    await self._stop.wait()          # connection opened and closed in the SAME task
        except BaseException as exc:
            self.error = f"{type(exc).__name__}: {exc}"
            self._ready.set()

    def start(self, timeout: int = 60) -> None:
        self.thread.start()
        asyncio.run_coroutine_threadsafe(self._serve(), self.loop)
        if not self._ready.wait(timeout):
            raise RuntimeError("MCP server start aagala (timeout)")
        if self.error:
            raise RuntimeError(self.error)

    def run(self, coro):
        return asyncio.run_coroutine_threadsafe(coro, self.loop).result(CALL_TIMEOUT)

    def stop(self) -> None:
        if self._stop is not None:
            self.loop.call_soon_threadsafe(self._stop.set)
        self.thread.join(10)


def to_dict(result) -> dict:
    text = "".join(getattr(block, "text", "") for block in result.content)
    if result.isError:
        return {"success": False, "error": text or "MCP isError"}
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return {"success": False, "error": f"JSON illa: {text[:120]}"}
    return data if isinstance(data, dict) else {"success": False, "error": f"dict illa: {text[:120]}"}


def main() -> int:
    words = [a.lower() for a in sys.argv[1:]]
    names = {n for n, _ in T.TESTS
             if n != "save_excel_workbook" and (not words or any(w in n for w in words))}
    if not names:
        print("Match aana tool illa.")
        return 1
    if any(n.startswith("remove_") for n in names):
        names.add("append_excel_data")           # remove_* tests need the rows append adds
    selected = [(n, f) for n, f in T.TESTS if n in names]

    problem = T.preflight()
    if problem:
        print(f"PREFLIGHT FAILED (dataset-ai thodala): {problem}")
        return 1

    bridge = McpBridge()
    try:
        bridge.start()
    except Exception as exc:
        print(f"MCP SERVER START FAILED: {exc}")
        return 1

    listed = {t.name for t in bridge.run(bridge.session.list_tools()).tools}
    print(f"MCP server connected. Tools registered: {len(listed)}")
    print(f"Dataset: {T.WORKBOOK} / {T.SHEET}. Tests thodangudhu, Excel-la click pannaadhinga...\n")

    def mcp_call(tool, **kwargs):
        """Drop-in replacement for test_all_tools.call(): same arguments, but through MCP."""
        kwargs.setdefault("workbook_name", T.S["book"].name)
        if "sheet_name" in inspect.signature(tool).parameters:
            kwargs.setdefault("sheet_name", T.SHEET)
        if tool.__name__ not in listed:
            return {"success": False, "error": f"'{tool.__name__}' MCP server-la register aagala"}
        return to_dict(bridge.run(bridge.session.call_tool(tool.__name__, kwargs)))

    T.call = mcp_call                              # tests look up `call` in T's globals

    results = []
    try:
        for number, (name, func) in enumerate(selected, 1):
            status, detail = T.run_one(func)
            if name not in listed and status == "PASS":
                status, detail = "FAIL", f"{name} MCP-la illa"
            results.append((name, status, detail))
            print(f"[{number:>2}/{len(selected)}] {status:<5} {name:<28} {detail}")
    finally:
        try:
            T.cleanup()
        except Exception as exc:
            print(f"CLEANUP ERROR: {type(exc).__name__}: {exc}")
        bridge.stop()

    intact, problems = T.dataset_intact()
    passed = sum(1 for r in results if r[1] == "PASS")
    print("\n" + "=" * 64)
    print(f"RESULT (via MCP): {passed}/{len(results)} passed")
    for name, status, detail in results:
        if status != "PASS":
            print(f"  - {name} [{status}]: {detail}")
    if intact:
        print("DATASET CHECK: PASS (values, sheets, charts ellaam test-ku munnadi maadhiri irukku)")
    else:
        print("DATASET CHANGED:", "; ".join(problems))
    print("=" * 64)
    return 0 if passed == len(results) and intact else 1


if __name__ == "__main__":
    sys.exit(main())