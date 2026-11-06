"""Ollama engine: the only file that talks to Ollama/Gemma.
Knows nothing about Excel or MCP.

Features:
- Tool-call turns get a bigger output budget (JSON is not cut off)
- temperature 0.1 (deterministic tool calling)
- Context overflow warning (detects silent truncation)
- Optional streaming + stop_event (the Stop button really works)
- Optional JSON mode (for the planner)
- health_check() (clear message when Ollama is down or the model is missing)
- [TIMING] log per call
"""
import threading
import time
from typing import Any

import ollama

import json
import re


def _parse_raw_call(text, valid_names):
    """Gemma raw text name{a:1,b:<|"|>x<|"|>} -> {"name","arguments"} or None."""
    m = re.search(r"(\w+)\s*(\{.*\})", text or "", re.S)
    if not m or m.group(1) not in valid_names:
        return None
    strings = []

    def stash(mm):
        strings.append(mm.group(1))
        return f'"__S{len(strings) - 1}__"'

    body = re.sub(r'<\|"\|>(.*?)<\|"\|>', stash, m.group(2), flags=re.S)
    body = re.sub(r'([{,]\s*)(\w+)\s*:', r'\1"\2":', body)
    try:
        args = json.loads(body)
    except json.JSONDecodeError:
        return None

    def restore(v):
        if isinstance(v, str):
            s = re.fullmatch(r"__S(\d+)__", v)
            return strings[int(s.group(1))] if s else v
        return v

    return {"name": m.group(1), "arguments": {k: restore(v) for k, v in args.items()}}

DEFAULT_MODEL = "gemma4:e2b"
DEFAULT_NUM_CTX = 4096
NUM_PREDICT_CHAT = 400       # plain chat
NUM_PREDICT_TOOLS = 512      # tool-call turn (arguments JSON is small)
DEFAULT_TEMPERATURE = 0.1
DEFAULT_KEEP_ALIVE = "30m"
CTX_WARN_MARGIN = 600        # warn when prompt + output is close to num_ctx


def _metric(response: object, key: str) -> int:
    """Read an Ollama metric from object-style or dict-style responses."""
    value = getattr(response, key, None)
    if value is None and isinstance(response, dict):
        value = response.get(key)
    return value or 0


class OllamaEngine:

    def __init__(
        self,
        model: str = DEFAULT_MODEL,
        think: bool = False,
        timing: bool = True,
    ) -> None:
        self.model = model
        self.think = think      # False = fast. True = slow on CPU.
        self.timing = timing    # True = print a [TIMING] line per LLM call

    # ---------------- HEALTH ----------------

    def health_check(self) -> dict[str, Any]:
        """Check that Ollama is running and the model is installed."""
        try:
            listing = ollama.list()
        except Exception as exc:
            return {
                "success": False,
                "error": (
                    "Ollama is not running. Start it with 'ollama serve' "
                    f"in a terminal. ({type(exc).__name__})"
                ),
            }

        models = getattr(listing, "models", None)
        if models is None and isinstance(listing, dict):
            models = listing.get("models", [])

        names: set[str] = set()
        for item in models or []:
            name = getattr(item, "model", None)
            if name is None and isinstance(item, dict):
                name = item.get("model") or item.get("name")
            if name:
                names.add(str(name))

        wanted = self.model
        found = any(
            n == wanted or n.split(":")[0] == wanted.split(":")[0] and wanted in n
            or n == f"{wanted}:latest"
            for n in names
        )
        if not found:
            return {
                "success": False,
                "error": (
                    f"Model '{wanted}' is not installed. "
                    f"Run 'ollama pull {wanted}'. "
                    f"Available: {sorted(names)}"
                ),
            }
        return {"success": True}

    # ---------------- CHAT ----------------

    def chat(
        self,
        messages: list[dict],
        tools: list[dict] | None = None,
        num_predict: int | None = None,
        json_mode: bool = False,
        stop_event: threading.Event | None = None,
    ) -> dict:
        """Send a message list (+ optional tools) to the model.

        Returns:
            {"success": True, "content": str,
             "tool_calls": [{"name": str, "arguments": dict}]}
            {"success": False, "error": str}
            {"success": False, "stopped": True, "error": "Stopped."}
        """
        if not messages:
            return {"success": False, "error": "The messages list is empty."}

        options = {
            "num_ctx": DEFAULT_NUM_CTX,
            "num_predict": num_predict
            or (NUM_PREDICT_TOOLS if tools else NUM_PREDICT_CHAT),
            "temperature": DEFAULT_TEMPERATURE,
        }

        kwargs: dict[str, Any] = {
            "model": self.model,
            "messages": messages,
            "think": self.think,
            "options": options,
            "keep_alive": DEFAULT_KEEP_ALIVE,
        }
        if tools:
            kwargs["tools"] = tools
        elif json_mode:
            kwargs["format"] = "json"   # do not combine tools and format

        try:
            t0 = time.perf_counter()

            if stop_event is not None:
                streamed = self._chat_stream(kwargs, stop_event)
                if streamed is None:
                    return {"success": False, "stopped": True, "error": "Stopped."}
                content, tool_calls, final = streamed
            else:
                final = ollama.chat(**kwargs)
                message = final.message
                content = (message.content or "").strip()
                tool_calls = self._extract_calls(message)

            elapsed = time.perf_counter() - t0

        except Exception as exc:
            return {"success": False, "error": f"{type(exc).__name__}: {exc}"}

        prompt_tokens = _metric(final, "prompt_eval_count")
        gen_tokens = _metric(final, "eval_count")

        if self.timing:
            print(
                f"[TIMING] {elapsed:.1f}s | "
                f"prompt={prompt_tokens} tok | "
                f"gen={gen_tokens} tok | "
                f"prefill={_metric(final, 'prompt_eval_duration') / 1e9:.1f}s | "
                f"gen_time={_metric(final, 'eval_duration') / 1e9:.1f}s"
            )

        if prompt_tokens + options["num_predict"] >= DEFAULT_NUM_CTX - CTX_WARN_MARGIN:
            print(
                f"[WARN] prompt={prompt_tokens} tok, num_ctx={DEFAULT_NUM_CTX}. "
                "Close to the context limit, the prompt may be truncated. "
                "Reduce tools / history / knowledge."
            )
        if not tool_calls and content:
            names = {t["function"]["name"] for t in (tools or [])}
            parsed = _parse_raw_call(content, names)
            if parsed:
                tool_calls = [parsed]
                content = ""
        return {
            "success": True,
            "content": content,
            "tool_calls": tool_calls,
        }

    # ---------------- INTERNAL ----------------

    @staticmethod
    def _extract_calls(message: Any) -> list[dict[str, Any]]:
        return [
            {
                "name": call.function.name,
                "arguments": dict(call.function.arguments or {}),
            }
            for call in (getattr(message, "tool_calls", None) or [])
        ]

    def _chat_stream(
        self,
        kwargs: dict[str, Any],
        stop_event: threading.Event,
    ) -> tuple[str, list[dict[str, Any]], Any] | None:
        """Stream the response. Returns None if stopped midway."""
        parts: list[str] = []
        calls: list[dict[str, Any]] = []
        last = None

        stream = ollama.chat(**kwargs, stream=True)
        try:
            for chunk in stream:
                if stop_event.is_set():
                    return None
                last = chunk
                message = chunk.message
                if message.content:
                    parts.append(message.content)
                calls.extend(self._extract_calls(message))
        finally:
            close = getattr(stream, "close", None)
            if callable(close):
                try:
                    close()
                except Exception:
                    pass

        return "".join(parts).strip(), calls, last

    # ---------------- LEGACY ----------------

    def generate_response(self, prompt: str) -> str:
        """Old simple method, kept so existing code does not break."""
        result = self.chat([{"role": "user", "content": prompt}])
        if not result["success"]:
            return f"Ollama Error: {result['error']}"
        return result["content"]


if __name__ == "__main__":
    engine = OllamaEngine()
    print(engine.health_check())
    print(engine.generate_response("Hello! Introduce yourself briefly."))