"""Level 1 test for backend/ollama_engine.py (no Excel, no MCP).
Run: python test_engine.py   (ollama serve must be running)
"""
from backend.ollama_engine import OllamaEngine

TOOLS = [{
    "type": "function",
    "function": {
        "name": "get_total_sales",
        "description": "Calculate the total of a numeric column in the Excel sheet.",
        "parameters": {
            "type": "object",
            "properties": {"column": {"type": "string", "description": "Column name"}},
            "required": ["column"],
        },
    },
}]

results = []


def check(name: str, condition, detail: str = "") -> None:
    results.append(bool(condition))
    print(f"{'PASS' if condition else 'FAIL'}  {name}  {detail}")


engine = OllamaEngine()

r = engine.chat([{"role": "user", "content": "Hello, how are you?"}], tools=TOOLS)
check("plain chat, no tool call", r["success"] and r["content"] and not r["tool_calls"], str(r)[:90])

r = engine.chat([{"role": "user", "content": "What is the total of the revenue column?"}], tools=TOOLS)
good = r["success"] and r["tool_calls"] and r["tool_calls"][0]["name"] == "get_total_sales" \
    and r["tool_calls"][0]["arguments"].get("column") == "revenue"
check("tool call returned", good, str(r.get("tool_calls", r)))

history = [
    {"role": "user", "content": "What is the total of the revenue column?"},
    {"role": "assistant", "content": "", "tool_calls": [
        {"function": {"name": "get_total_sales", "arguments": {"column": "revenue"}}}]},
    {"role": "tool", "tool_name": "get_total_sales", "content": "119231.42"},
]
r = engine.chat(history, tools=TOOLS)
check("history + tool result -> final answer", r["success"] and "119" in r["content"], str(r)[:90])

check("empty messages -> clean error", engine.chat([])["success"] is False)

bad = OllamaEngine(model="no-such-model:1b").chat([{"role": "user", "content": "hi"}])
check("wrong model -> clean error, no crash", bad["success"] is False, bad.get("error", "")[:70])

check("old generate_response still works", OllamaEngine().generate_response("Say hi in 3 words"))

print(f"\nRESULT: {sum(results)}/{len(results)} passed")