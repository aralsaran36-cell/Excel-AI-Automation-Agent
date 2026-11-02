import sys
from pathlib import Path

# ----------------------------------
# Project Root Path
# ----------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))


# ----------------------------------
# Imports
# ----------------------------------

import streamlit as st

try:
    from backend.orchestrator import process_user_request as process_message
except ImportError as e:
    st.error(f"Orchestrator Import Error: {e}")
    st.stop()


# Ensure data directory exists
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)


# ----------------------------------
# Page Configuration
# ----------------------------------

st.set_page_config(
    page_title="Excel AI Agent",
    page_icon="📊",
    layout="centered"
)


# ----------------------------------
# Session State
# ----------------------------------

if "messages" not in st.session_state:
    st.session_state.messages = []

if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []


# ----------------------------------
# Application Header
# ----------------------------------

st.title("📊 Excel AI Agent")

st.caption(
    "Your AI-powered Excel Assistant"
)

st.divider()


# ----------------------------------
# Display Chat History
# ----------------------------------

for message in st.session_state.messages:

    with st.chat_message(message["role"]):

        st.markdown(message["content"])


# ----------------------------------
# File Upload
# ----------------------------------

with st.popover("＋ Upload Files"):

    uploaded_files = st.file_uploader(
        "Upload Excel or CSV files",
        type=[
            "xlsx",
            "xls",
            "xlsm",
            "csv"
        ],
        accept_multiple_files=True
    )

    if uploaded_files:
        saved_paths = []
        for uploaded_file in uploaded_files:
            file_path = DATA_DIR / uploaded_file.name
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            saved_paths.append(str(file_path))

        st.session_state.uploaded_files = saved_paths

        st.success(
            f"{len(uploaded_files)} file(s) saved successfully!"
        )


# ----------------------------------
# Chat Input
# ----------------------------------

user_input = st.chat_input(
    "Ask anything about your Excel data..."
)


# ----------------------------------
# Backend Processing
# ----------------------------------

if user_input:

    # Store user message
    st.session_state.messages.append({
        "role": "user",
        "content": user_input
    })

    # Display user message
    with st.chat_message("user"):

        st.markdown(user_input)

    # Get AI response
    with st.chat_message("assistant"):

        with st.spinner("Processing your request..."):

            try:

                response = process_message(user_input)

                if response is None:
                    response = "Request completed."

                response = str(response)

                st.markdown(response)

                st.session_state.messages.append({
                    "role": "assistant",
                    "content": response
                })

            except Exception as e:

                error_message = f"Error: {str(e)}"

                st.error(error_message)

                st.session_state.messages.append({
                    "role": "assistant",
                    "content": error_message
                })