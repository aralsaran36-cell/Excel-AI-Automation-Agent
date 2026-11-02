"""Standalone test: does Gemma 4 E2B call a dummy tool? (No Excel, no MCP)
Run: python test_ollama_tools.py
"""
import time

import ollama

MODEL = "gemma4:e2b"

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "get_total_sales",
            "description": "Calculate the total of a numeric column in the Excel sheet.",
            "parameters": {
                "type": "object",
                "properties": {
                    "column": {
                        "type": "string",
                        "description": "Column name, e.g. revenue",
                    }
                },
                "required": ["column"],
            },
        },
    }
]

PROMPTS = [
    "What is the total of the revenue column?",
    "Hello, how are you?",  # should NOT call a tool
]


def run(prompt: str, think: bool) -> None:
    start = time.time()
    try:
        response = ollama.chat(
            model=MODEL,
            messages=[{"role": "user", "content": prompt}],
            tools=TOOLS,
            think=think,
        )
    except Exception as exc:
        print(f"  ERROR: {type(exc).__name__}: {exc}")
        return

    elapsed = time.time() - start
    message = response.message
    print(f"  time     : {elapsed:.1f}s")
    print(f"  content  : {(message.content or '').strip()[:150]!r}")
    if message.tool_calls:
        for call in message.tool_calls:
            print(f"  TOOL CALL: {call.function.name}({dict(call.function.arguments)})")
    else:
        print("  TOOL CALL: none")


for think in (False, True):
    print("=" * 60)
    print(f"think={think}")
    for prompt in PROMPTS:
        print(f"\nPrompt: {prompt}")
        run(prompt, think)
