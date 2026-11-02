"""Knowledge answer: RAG retrieval + Gemma generation.
Returns a short generated answer, not the raw markdown chunks.
"""
from backend.ollama_engine import OllamaEngine
from backend.rag import retrieve_context

TOP_K = 3

SYSTEM_PROMPT = (
    "You are an Excel expert assistant. Answer the user's question using ONLY the "
    "context below. Be short and clear (max 6 lines). Use a small example if the "
    "context has one. If the context does not contain the answer, say you don't know."
)


def answer_question(question: str, engine: OllamaEngine | None = None) -> dict:
    """Retrieve relevant chunks and let Gemma write the final answer."""
    if not question or not question.strip():
        return {"success": False, "error": "question empty."}

    engine = engine or OllamaEngine()
    try:
        chunks = retrieve_context(question, top_k=TOP_K)
    except Exception as exc:
        return {"success": False, "error": f"Retrieval error: {type(exc).__name__}: {exc}"}
    if not chunks:
        return {"success": True, "answer": "Knowledge base la relevant info kidaikkala.", "sources": []}

    context = "\n\n---\n\n".join(c["content"] for c in chunks)
    messages = [
        {"role": "system", "content": f"{SYSTEM_PROMPT}\n\nContext:\n{context}"},
        {"role": "user", "content": question},
    ]
    result = engine.chat(messages)
    if not result["success"]:
        return {"success": False, "error": result["error"]}

    sources = sorted({c["source"] for c in chunks})
    return {"success": True, "answer": result["content"], "sources": sources}


if __name__ == "__main__":
    while True:
        q = input("\nQuestion (exit to stop): ").strip()
        if q.lower() == "exit":
            break
        r = answer_question(q)
        print(r["answer"] if r["success"] else f"ERROR: {r['error']}")
        if r.get("sources"):
            print("Sources:", ", ".join(r["sources"]))