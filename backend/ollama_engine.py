
"""Ollama engine: the only file that talks to Ollama/Gemma.
Knows nothing about Excel or MCP.
"""
import ollama

DEFAULT_MODEL = "gemma4:e2b"


class OllamaEngine:

    def __init__(self, model: str = DEFAULT_MODEL, think: bool = False):
        self.model = model
        self.think = think  # False = fast (4s). True = slow (27s on simple chat)

    def chat(self, messages: list[dict], tools: list[dict] | None = None) -> dict:
        """Send a message list (+ optional tools) to Gemma.

        Returns:
            {"success": True, "content": str, "tool_calls": [{"name": str, "arguments": dict}]}
            {"success": False, "error": str}
        """
        if not messages:
            return {"success": False, "error": "messages list empty."}
        try:
            kwargs = {"model": self.model, "messages": messages, "think": self.think}
            if tools:
                kwargs["tools"] = tools
            response = ollama.chat(**kwargs)
        except Exception as exc:
            return {"success": False, "error": f"{type(exc).__name__}: {exc}"}

        message = response.message
        tool_calls = [
            {"name": call.function.name, "arguments": dict(call.function.arguments)}
            for call in (message.tool_calls or [])
        ]
        return {
            "success": True,
            "content": (message.content or "").strip(),
            "tool_calls": tool_calls,
        }

    def generate_response(self, prompt: str) -> str:
        """Old simple method, kept so existing code does not break."""
        result = self.chat([{"role": "user", "content": prompt}])
        if not result["success"]:
            return f"Ollama Error: {result['error']}"
        return result["content"]


if __name__ == "__main__":
    engine = OllamaEngine()
    print(engine.generate_response("Hello! Introduce yourself briefly."))