"""Additional test: only the failing/review cases. Read-only.
Run: python test_additional.py   (ollama serve + Excel open)
"""
import asyncio

from backend.agent import ExcelAgent

CASES = [
    ("profit column total sollu", "profit"),
]


async def main() -> None:
    agent = ExcelAgent()
    await agent.connect()
    try:
        for question, bad_column in CASES:
            out = await agent.run(question)
            reply = out.get("reply") or out.get("error") or ""
            used = [(s["tool"], s["arguments"]) for s in out.get("steps", [])]
            wrong = any(
                bad_column not in str(args.get("column", "")).lower() and s_ok
                for (_, args), s_ok in [(u, True) for u in used]
            )
            says_missing = "illa" in reply.lower() or "not" in reply.lower() or "doesn" in reply.lower()
            status = "PASS" if (says_missing and not wrong) else "FAIL"
            print(f"{status}  Q: {question}")
            print("      reply:", reply[:150])
            print("      tools:", used)
    finally:
        await agent.close()


asyncio.run(main())