import re
from pathlib import Path

import chromadb

from llama_index.core import (
    SimpleDirectoryReader,
    StorageContext,
    VectorStoreIndex,
    Settings,
)
from llama_index.core.node_parser import MarkdownNodeParser
from llama_index.embeddings.huggingface import HuggingFaceEmbedding
from llama_index.vector_stores.chroma import ChromaVectorStore
from llama_index.retrievers.bm25 import BM25Retriever


# =====================================================
# PATH CONFIGURATION
# =====================================================

BASE_DIR = Path(__file__).resolve().parent.parent
EXECUTION_KB_DIR = BASE_DIR / "excution_kb"
CHROMA_DB_DIR = BASE_DIR / "data" / "chroma_db"
COLLECTION_NAME = "excel_knowledge_base"


# =====================================================
# RAG CONFIGURATION
# =====================================================

EMBEDDING_MODEL = "BAAI/bge-small-en-v1.5"

DEFAULT_TOP_K = 2          # max chunks per step
MIN_VECTOR_SCORE = 0.45    # weak hits skipped (tune with test_rag.py)
RRF_K = 60

TOOLS_LINE = re.compile(r"^\s*Tools:\s*(.+)$", re.MULTILINE)


# =====================================================
# GLOBAL VARIABLES
# =====================================================

_index = None
_chroma_client = None
_chroma_collection = None
_nodes = None


# =====================================================
# EMBEDDING MODEL + CHUNKER
# =====================================================

def initialize_embedding_model():
    Settings.embed_model = HuggingFaceEmbedding(
        model_name=EMBEDDING_MODEL
    )
    # One "##" section = one chunk
    Settings.node_parser = MarkdownNodeParser()


# =====================================================
# CHROMADB
# =====================================================

def initialize_chromadb():
    global _chroma_client
    global _chroma_collection

    CHROMA_DB_DIR.mkdir(parents=True, exist_ok=True)

    _chroma_client = chromadb.PersistentClient(
        path=str(CHROMA_DB_DIR)
    )
    _chroma_collection = _chroma_client.get_or_create_collection(
        name=COLLECTION_NAME
    )
    return _chroma_collection


# =====================================================
# LOAD DOCUMENTS
# =====================================================

def load_knowledge_documents():
    if not EXECUTION_KB_DIR.exists():
        raise FileNotFoundError(
            f"Execution KB not found: {EXECUTION_KB_DIR}"
        )

    documents = SimpleDirectoryReader(
        input_dir=str(EXECUTION_KB_DIR),
        recursive=True,
        required_exts=[".md"],
    ).load_data()

    if not documents:
        raise ValueError(
            f"No Markdown documents found in {EXECUTION_KB_DIR}"
        )

    print(f"Loaded {len(documents)} documents from {EXECUTION_KB_DIR.name}.")
    return documents


# =====================================================
# BUILD NODES (chunks) WITH TOOL METADATA
# =====================================================

def extract_tools(text):
    """Read the 'Tools:' line of a chunk -> list of tool names."""
    match = TOOLS_LINE.search(text)
    if not match:
        return []
    parts = re.split(r"->|,|\|", match.group(1))
    return [p.strip() for p in parts if p.strip()]


def build_nodes():
    documents = load_knowledge_documents()
    nodes = Settings.node_parser.get_nodes_from_documents(documents)

    for node in nodes:
        tools = extract_tools(node.get_content())
        node.metadata["tools"] = ",".join(tools)

        # metadata must not go into embedding or into the LLM text
        node.excluded_embed_metadata_keys = list(node.metadata.keys())
        node.excluded_llm_metadata_keys = list(node.metadata.keys())

    tagged = sum(1 for n in nodes if n.metadata["tools"])
    print(f"{len(nodes)} chunks, {tagged} with Tools: tag.")
    return nodes


def get_nodes():
    global _nodes
    if _nodes is None:
        initialize_embedding_model()
        _nodes = build_nodes()
    return _nodes


# =====================================================
# BUILD INDEX
# =====================================================

def build_index(force_rebuild=False):
    global _index
    global _nodes

    initialize_embedding_model()
    collection = initialize_chromadb()

    if force_rebuild:
        _chroma_client.delete_collection(name=COLLECTION_NAME)
        collection = initialize_chromadb()
        _nodes = None

    vector_store = ChromaVectorStore(chroma_collection=collection)

    if collection.count() > 0 and not force_rebuild:
        print("Loading existing Knowledge Base index...")
        _index = VectorStoreIndex.from_vector_store(
            vector_store=vector_store
        )
        return _index

    storage_context = StorageContext.from_defaults(
        vector_store=vector_store
    )

    print("Creating Knowledge Base embeddings...")
    _index = VectorStoreIndex(
        nodes=get_nodes(),
        storage_context=storage_context,
        show_progress=True,
    )
    print("Knowledge Base indexing completed.")
    return _index


def get_index():
    global _index
    if _index is None:
        _index = build_index()
    return _index


# =====================================================
# TOOL FILTER
# =====================================================

def _matches_tools(tools_csv, wanted):
    if wanted is None:
        return True          # no filter: search everything
    have = {t for t in (tools_csv or "").split(",") if t}
    return bool(have & wanted)


# =====================================================
# HYBRID SEARCH (tool-filtered)
# =====================================================

def retrieve_context(question, tools=None, top_k=DEFAULT_TOP_K):
    """
    tools = list of selected tool names. Only chunks whose
    'Tools:' line contains one of them are searched.
    tools=None -> search all chunks (for testing only).
    """
    if not question or not question.strip():
        return []

    wanted = set(tools) if tools else None

    index = get_index()
    pool = max(_chroma_collection.count(), 1)

    # ---- Vector search, then keep only tool-matching chunks
    vector_results = index.as_retriever(
        similarity_top_k=pool
    ).retrieve(question)

    vector_results = [
        r for r in vector_results
        if _matches_tools(r.node.metadata.get("tools"), wanted)
    ]

    # ---- BM25 only on tool-matching chunks
    candidates = [
        n for n in get_nodes()
        if _matches_tools(n.metadata.get("tools"), wanted)
    ]

    bm25_results = []
    if candidates:
        try:
            bm25 = BM25Retriever.from_defaults(
                nodes=candidates,
                similarity_top_k=min(len(candidates), top_k * 2),
            )
            bm25_results = bm25.retrieve(question)
        except Exception:
            bm25_results = []

    # ---- Reciprocal Rank Fusion
    combined = {}

    for search_type, results in (
        ("vector", vector_results),
        ("bm25", bm25_results),
    ):
        for rank, item in enumerate(results, start=1):
            content = item.node.get_content()
            key = content.strip()

            if key not in combined:
                combined[key] = {
                    "content": content,
                    "source": item.node.metadata.get("file_name", "Unknown"),
                    "tools": item.node.metadata.get("tools", ""),
                    "score": 0.0,
                    "vector_score": None,
                    "bm25_score": None,
                }

            combined[key]["score"] += 1 / (RRF_K + rank)
            combined[key][f"{search_type}_score"] = item.score

    ranked = sorted(
        combined.values(),
        key=lambda x: x["score"],
        reverse=True,
    )

    # ---- Skip weak hits
    ranked = [
        r for r in ranked
        if (r["vector_score"] or 0) >= MIN_VECTOR_SCORE
    ]

    return ranked[:top_k]


# =====================================================
# GUIDANCE FOR EXECUTOR (selected tools only)
# =====================================================

def get_guidance(step_text, tool_names, max_chunks=DEFAULT_TOP_K):
    """Returns guidance text, or "" if nothing relevant."""
    results = retrieve_context(
        step_text,
        tools=tool_names,
        top_k=max_chunks,
    )
    return "\n\n".join(r["content"] for r in results)


# =====================================================
# INFO + REBUILD
# =====================================================

def get_knowledge_base_info():
    collection = initialize_chromadb()
    return {
        "knowledge_base_path": str(EXECUTION_KB_DIR),
        "chroma_db_path": str(CHROMA_DB_DIR),
        "total_vectors": collection.count(),
        "collection_name": COLLECTION_NAME,
    }


def rebuild_knowledge_base():
    global _index
    global _nodes

    _index = None
    _nodes = None
    return build_index(force_rebuild=True)