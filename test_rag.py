from backend.rag import retrieve_context

while True:

    question = input("\nEnter your question (exit to stop): ")

    if question.lower() == "exit":
        break

    results = retrieve_context(question)

    print("\n========== HYBRID SEARCH RESULTS ==========\n")

    if not results:
        print("No relevant results found.")
        continue

    for i, result in enumerate(results, start=1):

        print(f"\nResult {i}")
        print("Source:", result["source"])
        print("Combined Score:", result["score"])
        print("Vector Score:", result["vector_score"])
        print("BM25 Score:", result["bm25_score"])
        print("Content:")
        print(result["content"])
        print("-" * 60)