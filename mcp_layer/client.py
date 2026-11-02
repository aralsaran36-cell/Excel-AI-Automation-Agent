"""MCP client: connects ONCE to the Excel MCP server (stdio), discovers and calls tools.

The connection is owned by ONE long-lived task. connect() starts it, disconnect()
stops it, and the task itself closes the connection (anyio requires open and close
to happen in the same task).

Run smoke test (Excel open): python -m mcp_layer.client
"""
import asyncio
import json
import sys
from pathlib import Path
from typing import Any

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

ROOT = Path(__file__).resolve().parent.parent
CONNECT_TIMEOUT = 60
CLOSE_TIMEOUT = 10


class MCPConnectionError(RuntimeError):
    """MCP connection illa illa udainjiduchu. Reconnect pannanum."""


class MCPClient:
    def __init__(self) -> None:
        self.session: ClientSession | None = None
        self._owner: asyncio.Task | None = None
        self._ready: asyncio.Event | None = None
        self._stop: asyncio.Event | None = None
        self._error: BaseException | None = None

    @property
    def is_connected(self) -> bool:
        return self.session is not None and self._owner is not None and not self._owner.done()

    async def connect(self) -> None:
        if self.is_connected:
            return
        await self.disconnect()  # clear any half-dead state first
        self._ready, self._stop, self._error = asyncio.Event(), asyncio.Event(), None
        self._owner = asyncio.create_task(
            self._own_connection(self._ready, self._stop), name="mcp-connection"
        )
        try:
            await asyncio.wait_for(self._ready.wait(), CONNECT_TIMEOUT)
        except asyncio.TimeoutError:
            await self.disconnect()
            raise MCPConnectionError("MCP server start aagala (timeout).")
        if self._error is not None:
            error = self._error
            await self.disconnect()
            raise error

    async def _own_connection(self, ready: asyncio.Event, stop: asyncio.Event) -> None:
        params = StdioServerParameters(
            command=sys.executable,
            args=["-m", "mcp_layer.server"],
            cwd=str(ROOT),
        )
        try:
            async with stdio_client(params) as (read, write):
                async with ClientSession(read, write) as session:
                    await session.initialize()
                    self.session = session
                    ready.set()
                    await stop.wait()  # stay alive until disconnect() asks us to stop
        except Exception as exc:
            if not ready.is_set():
                self._error = exc
        finally:
            self.session = None
            ready.set()

    async def get_tools(self) -> list[Any]:
        session = self._require_session()
        result = await session.list_tools()
        return list(result.tools)

    async def call_tool(self, name: str, arguments: dict[str, Any] | None = None) -> dict[str, Any]:
        session = self._require_session()
        result = await session.call_tool(name, arguments or {})
        text = "".join(c.text for c in result.content if getattr(c, "type", "") == "text")
        if result.isError:
            return {"success": False, "error": text or "MCP tool error"}
        try:
            data = json.loads(text)
        except json.JSONDecodeError:
            return {"success": True, "result": text}
        return data if isinstance(data, dict) else {"success": True, "result": data}

    async def disconnect(self) -> None:
        owner, stop = self._owner, self._stop
        self._owner = self._stop = self._ready = None
        self.session = None
        if owner is None:
            return
        if stop is not None:
            stop.set()
        if not owner.done():
            await asyncio.wait({owner}, timeout=CLOSE_TIMEOUT)
            if not owner.done():
                owner.cancel()
                await asyncio.wait({owner}, timeout=CLOSE_TIMEOUT)
        if owner.done() and not owner.cancelled():
            owner.exception()  # mark as retrieved, avoids "never retrieved" warnings

    def _require_session(self) -> ClientSession:
        if self.session is None or self._owner is None or self._owner.done():
            raise MCPConnectionError("MCP connection illa. connect() call pannala illa connection udainjiduchu.")
        return self.session


async def _smoke_test() -> None:
    client = MCPClient()
    try:
        await client.connect()
        tools = await client.get_tools()
        print("Connected. Total tools:", len(tools))
        print("get_excel_sheet_names ->", await client.call_tool("get_excel_sheet_names", {}))
    finally:
        await client.disconnect()


if __name__ == "__main__":
    asyncio.run(_smoke_test())