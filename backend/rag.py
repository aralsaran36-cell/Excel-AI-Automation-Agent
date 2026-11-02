
from pathlib import Path

import chromadb

from llama_index.core import (
    SimpleDirectoryReader,
    StorageContext,
    VectorStoreIndex,
    Settings,
)
from llama_index.core.node_parser import SentenceSplitter
from llama_index.embeddings.huggingface import HuggingFaceEmbedding
from llama_index.vector_stores.chroma import ChromaVectorStore

from llama_index.retrievers.bm25 import BM25Retriever


# =====================================================
# PATH CONFIGURATION
# =====================================================

BASE_DIR = Path(__file__).resolve().parent.parent

KNOWLEDGE_BASE_DIR = BASE_DIR / "knowledge_base"

CHROMA_DB_DIR = BASE_DIR / "data" / "chroma_db"

COLLECTION_NAME = "excel_knowledge_base"


# =====================================================
# RAG CONFIGURATION
# =====================================================

EMBEDDING_MODEL = "BAAI/bge-small-en-v1.5"

CHUNK_SIZE = 300
CHUNK_OVERLAP = 50

DEFAULT_TOP_K = 5

BM25_TOP_K = 5


# =====================================================
# GLOBAL VARIABLES
# =====================================================

_index = None
_chroma_client = None
_chroma_collection = None

_bm25_retriever = None


# =====================================================
# INITIALIZE EMBEDDING MODEL
# =====================================================

def initialize_embedding_model():

    Settings.embed_model = HuggingFaceEmbedding(
        model_name=EMBEDDING_MODEL
    )

    Settings.node_parser = SentenceSplitter(
        chunk_size=CHUNK_SIZE,
        chunk_overlap=CHUNK_OVERLAP
    )


# =====================================================
# INITIALIZE CHROMADB
# =====================================================

def initialize_chromadb():

    global _chroma_client
    global _chroma_collection

    CHROMA_DB_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    _chroma_client = chromadb.PersistentClient(
        path=str(CHROMA_DB_DIR)
    )

    _chroma_collection = (
        _chroma_client.get_or_create_collection(
            name=COLLECTION_NAME
        )
    )

    return _chroma_collection


# =====================================================
# LOAD KNOWLEDGE BASE DOCUMENTS
# =====================================================

def load_knowledge_documents():

    if not KNOWLEDGE_BASE_DIR.exists():

        raise FileNotFoundError(
            f"Knowledge Base not found: {KNOWLEDGE_BASE_DIR}"
        )

    documents = SimpleDirectoryReader(
        input_dir=str(KNOWLEDGE_BASE_DIR),
        recursive=True,
        required_exts=[".md"]
    ).load_data()

    if not documents:

        raise ValueError(
            "No Markdown documents found."
        )

    print(f"Loaded {len(documents)} documents.")

    return documents


# =====================================================
# BUILD BM25 RETRIEVER
# =====================================================

def build_bm25_retriever():

    global _bm25_retriever

    initialize_embedding_model()

    documents = load_knowledge_documents()

    nodes = Settings.node_parser.get_nodes_from_documents(
        documents
    )

    _bm25_retriever = BM25Retriever.from_defaults(
        nodes=nodes,
        similarity_top_k=BM25_TOP_K
    )

    print(
        f"BM25 Retriever initialized with {len(nodes)} chunks."
    )

    return _bm25_retriever


# =====================================================
# BUILD KNOWLEDGE BASE INDEX
# =====================================================

def build_index(force_rebuild=False):

    global _index
    global _chroma_collection
    global _bm25_retriever

    initialize_embedding_model()

    collection = initialize_chromadb()

    if force_rebuild:

        _chroma_client.delete_collection(
            name=COLLECTION_NAME
        )

        collection = initialize_chromadb()

        _bm25_retriever = None

    vector_store = ChromaVectorStore(
        chroma_collection=collection
    )

    if collection.count() > 0 and not force_rebuild:

        print("Loading existing Knowledge Base index...")

        _index = VectorStoreIndex.from_vector_store(
            vector_store=vector_store
        )

        return _index

    documents = load_knowledge_documents()

    storage_context = StorageContext.from_defaults(
        vector_store=vector_store
    )

    print("Creating Knowledge Base embeddings...")

    _index = VectorStoreIndex.from_documents(
        documents,
        storage_context=storage_context,
        show_progress=True
    )

    print("Knowledge Base indexing completed.")

    return _index


# =====================================================
# GET INDEX
# =====================================================

def get_index():

    global _index

    if _index is None:

        _index = build_index()

    return _index


# =====================================================
# GET BM25 RETRIEVER
# =====================================================

def get_bm25_retriever():

    global _bm25_retriever

    if _bm25_retriever is None:

        _bm25_retriever = build_bm25_retriever()

    return _bm25_retriever


# =====================================================
# HYBRID SEARCH
# =====================================================

def retrieve_context(
    question,
    top_k=DEFAULT_TOP_K
):

    if not question or not question.strip():

        return []

    index = get_index()

    # Vector Search
    vector_retriever = index.as_retriever(
        similarity_top_k=top_k
    )

    vector_results = vector_retriever.retrieve(
        question
    )

    # BM25 Search
    bm25_retriever = get_bm25_retriever()

    bm25_retriever.similarity_top_k = BM25_TOP_K

    bm25_results = bm25_retriever.retrieve(
        question
    )

    # Reciprocal Rank Fusion
    combined_scores = {}

    all_results = [
        ("vector", vector_results),
        ("bm25", bm25_results)
    ]

    RRF_K = 60

    for search_type, results in all_results:

        for rank, item in enumerate(results, start=1):

            node = item.node

            content = node.get_content()

            key = content.strip()

            if key not in combined_scores:

                combined_scores[key] = {
                    "content": content,
                    "source": node.metadata.get(
                        "file_name",
                        "Unknown"
                    ),
                    "score": 0.0,
                    "vector_score": None,
                    "bm25_score": None
                }

            combined_scores[key]["score"] += (
                1 / (RRF_K + rank)
            )

            if search_type == "vector":

                combined_scores[key]["vector_score"] = (
                    item.score
                )

            else:

                combined_scores[key]["bm25_score"] = (
                    item.score
                )

    # Sort by combined score
    ranked_results = sorted(
        combined_scores.values(),
        key=lambda x: x["score"],
        reverse=True
    )

    return ranked_results[:top_k]


# =====================================================
# FORMAT RETRIEVED CONTEXT
# =====================================================

def get_context_text(
    question,
    top_k=DEFAULT_TOP_K
):

    results = retrieve_context(
        question,
        top_k
    )

    if not results:

        return "No relevant Knowledge Base content found."

    context_parts = []

    for index, result in enumerate(results, start=1):

        context_parts.append(
            f"""
Source {index}: {result['source']}

{result['content']}
"""
        )

    return "\n\n".join(context_parts)


# =====================================================
# KNOWLEDGE BASE STATISTICS
# =====================================================

def get_knowledge_base_info():

    collection = initialize_chromadb()

    return {
        "knowledge_base_path": str(KNOWLEDGE_BASE_DIR),
        "chroma_db_path": str(CHROMA_DB_DIR),
        "total_vectors": collection.count(),
        "collection_name": COLLECTION_NAME
    }


# =====================================================
# REBUILD KNOWLEDGE BASE
# =====================================================

def rebuild_knowledge_base():

    global _index
    global _bm25_retriever

    _index = None
    _bm25_retriever = None

    return build_index(
        force_rebuild=True
    )
