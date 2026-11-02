from backend.rag import rebuild_knowledge_base

if __name__ == "__main__":

    print("Rebuilding Knowledge Base...")

    index = rebuild_knowledge_base()

    print("\nKnowledge Base rebuilt successfully!")