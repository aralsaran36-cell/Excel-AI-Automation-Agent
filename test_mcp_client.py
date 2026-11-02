"""Level 3b: tests tools through MCP client -> server (stdio) with Excel open.
One connection, sequential calls. Safe: read tools, remove_* preview, and
write/formula WITHOUT overwrite (must be refused). Never sends confirm=True.
Run: python test_mcp_client.py
"""
import asyncio
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from mcp_layer.client import MCPClient  # noqa: E402

results: list[tuple[str, bool]] = []


def check(label: str, ok: bool, detail: str = "") -> None:
    results.append((label, ok))
    print(f"[{'PASS' if ok else 'FAIL'}] {label}" + (f" | {detail}" if detail else ""))


def snapshot(data: dict) -> str:
    return json.dumps(data, sort_keys=True, default=str)


async def main() -> None:
    client = MCPClient()
    try:
        await client.connect()
        tools = await client.get_tools()
        check("connect once + 17 tools", len(tools) == 17, f"tools={len(tools)}")

        before = await client.call_tool("read_excel_data", {"max_rows": 500})
        check("read_excel_data (snapshot before)", before.get("success") is True)

        r = await client.call_tool("get_excel_sheet_names", {})
        check("get_excel_sheet_names", r.get("success") is True, str(r.get("sheets")))

        r = await client.call_tool("read_excel_data", {"max_rows": 5})
        check("read_excel_data max_rows=5", r.get("success") is True)

        r = await client.call_tool("analyze_excel_column", {"column": "revenue"})
        check("analyze_excel_column revenue", r.get("success") is True, str(r)[:150])

        r = await client.call_tool("validate_excel_data", {})
        check("validate_excel_data", r.get("success") is True, str(r)[:150])

        r = await client.call_tool("remove_excel_duplicates", {})
        check("remove_excel_duplicates preview (no confirm)",
              r.get("success") is True and r.get("deleted") is not True, str(r)[:150])

        r = await client.call_tool("remove_excel_blank_rows", {})
        check("remove_excel_blank_rows preview (no confirm)",
              r.get("success") is True and r.get("deleted") is not True, str(r)[:150])

        r = await client.call_tool("write_excel_data", {"data": [["TEST"]], "start_cell": "A1"})
        check("write_excel_data without overwrite is refused",
              r.get("success") is False, f"error={str(r.get('error'))[:150]}")

        r = await client.call_tool("apply_excel_formula", {"cell": "A1", "formula": "=1+1"})
        check("apply_excel_formula without overwrite is refused",
              r.get("success") is False, f"error={str(r.get('error'))[:150]}")

        r = await client.call_tool("read_excel_data", {"workbook_name": "nope.xlsx"})
        check("wrong workbook_name gives success False", r.get("success") is False)

        r = await client.call_tool("analyze_excel_column", {})
        check("missing required arg gives success False", r.get("success") is False,
              str(r.get("error"))[:100])

        try:
            r = await client.call_tool("not_a_real_tool", {})
            check("unknown tool gives success False", r.get("success") is False,
                  str(r.get("error"))[:100])
        except Exception as exc:  # also acceptable: client raises
            check("unknown tool raises error", True, f"{type(exc).__name__}")

        after = await client.call_tool("read_excel_data", {"max_rows": 500})
        check("DATASET unchanged (before == after)", snapshot(before) == snapshot(after))
    finally:
        await client.disconnect()

    passed = sum(1 for _, ok in results if ok)
    print("=" * 60)
    print(f"RESULT: {passed}/{len(results)} passed")


asyncio.run(main())