"""Streamlit interface for the IEEE Agentic Formatter."""

import os
from io import BytesIO, StringIO

import streamlit as st
from docx import Document as DocxDocument
from dotenv import load_dotenv

from agent.text_parser import parse_raw_text, parse_raw_text_local
from formatter.docx_engine import generate_ieee_docx

load_dotenv()

st.set_page_config(
    page_title="IEEE Agentic Formatter",
    page_icon="📄",
    layout="wide",
)

st.title("IEEE Agentic Formatter")
st.caption(
    "Upload raw text, research notes, or a file and receive a "
    "strictly compliant IEEE conference paper in .docx format."
)

# --- Sidebar ---
with st.sidebar:
    st.header("Configuration")

    backend = st.radio(
        "Backend",
        ["Local (Ollama)", "Anthropic API"],
        index=0,
        help="Local uses Ollama — no API key needed. Anthropic uses claude-opus-4-8.",
    )

    if backend == "Anthropic API":
        api_key_input = st.text_input(
            "Anthropic API Key",
            type="password",
            value=os.getenv("ANTHROPIC_API_KEY", ""),
            help="Required. Get yours at console.anthropic.com",
        )
    else:
        api_key_input = ""
        ollama_url = st.text_input(
            "Ollama URL",
            value="http://localhost:11434",
            help="Default Ollama address.",
        )
        local_model = st.selectbox(
            "Local Model",
            ["qwen2.5:7b", "llama3:latest", "mistral:latest"],
            index=0,
            help="qwen2.5:7b gives the best structured JSON output.",
        )

    st.divider()
    if backend == "Local (Ollama)":
        st.markdown(
            "**Mode:** Local (no API key)\n\n"
            f"**Model:** `{local_model}`\n\n"
            "**Output:** IEEE 2-column, Letter size, Times New Roman\n\n"
            "Powered by Sunil Gentyala + python-docx"
        )
    else:
        st.markdown(
            "**Mode:** Anthropic API\n\n"
            "**Model:** `claude-opus-4-8`\n\n"
            "**Output:** IEEE 2-column, Letter size, Times New Roman\n\n"
            "Powered by Sunil Gentyala + python-docx"
        )

# --- Input tabs ---
tab_text, tab_file = st.tabs(["Paste Text", "Upload File"])

raw_text: str = ""

with tab_text:
    raw_text_input = st.text_area(
        "Paste your raw content here",
        height=400,
        placeholder=(
            "Research notes, a draft abstract, bullet points, loose paragraphs "
            "-- anything. The AI will structure it into a full IEEE paper."
        ),
    )
    if raw_text_input:
        raw_text = raw_text_input

with tab_file:
    uploaded = st.file_uploader(
        "Upload a .txt, .md, or .docx file",
        type=["txt", "md", "docx"],
        help="Plain text, Markdown, or Word (.docx) files up to 200 KB.",
    )
    if uploaded is not None:
        if uploaded.name.lower().endswith(".docx"):
            doc = DocxDocument(BytesIO(uploaded.read()))
            file_text = "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
        else:
            file_text = StringIO(uploaded.read().decode("utf-8", errors="ignore")).read()
        st.text_area("File preview", value=file_text[:3000], height=300, disabled=True)
        raw_text = file_text

# --- Generate ---
st.divider()
generate_btn = st.button("Generate IEEE Paper", type="primary", disabled=not raw_text.strip())

if generate_btn and raw_text.strip():

    # --- Validate config ---
    if backend == "Anthropic API":
        effective_key = api_key_input.strip() or os.getenv("ANTHROPIC_API_KEY", "")
        if not effective_key:
            st.error("Please provide an Anthropic API key in the sidebar or in your .env file.")
            st.stop()

    col1, col2 = st.columns(2)

    # --- Step 1: Parse ---
    model_label = local_model if backend == "Local (Ollama)" else "claude-opus-4-8"
    with st.spinner(f"Step 1/2 — Parsing with {model_label}..."):
        try:
            if backend == "Local (Ollama)":
                paper = parse_raw_text_local(
                    raw_text,
                    model=local_model,
                    ollama_url=ollama_url,
                )
            else:
                import anthropic
                client = anthropic.Anthropic(api_key=effective_key)
                paper = parse_raw_text(raw_text, client=client)
        except Exception as exc:
            st.error(f"Parsing failed: {exc}")
            st.stop()

    with col1:
        st.success("Parsing complete")
        st.subheader("Extracted Structure")
        st.markdown(f"**Title:** {paper.title}")
        st.markdown(f"**Authors:** {', '.join(a.name for a in paper.authors)}")
        st.markdown(f"**Keywords:** {', '.join(paper.keywords)}")
        st.markdown(f"**Sections:** {len(paper.sections)}")
        st.markdown(f"**References:** {len(paper.references)}")

    # --- Step 2: Format ---
    with st.spinner("Step 2/2 — Generating IEEE .docx..."):
        try:
            docx_bytes = generate_ieee_docx(paper)
        except Exception as exc:
            st.error(f"Formatting failed: {exc}")
            st.stop()

    with col2:
        st.success("Formatting complete")
        safe_title = "".join(c if c.isalnum() or c in " _-" else "_" for c in paper.title)[:60]
        filename = f"{safe_title.strip().replace(' ', '_')}_IEEE.docx"
        st.download_button(
            label="Download IEEE Paper (.docx)",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    with st.expander("Abstract preview"):
        st.write(paper.abstract)

    with st.expander("Section outline"):
        for s in paper.sections:
            prefix = "   " if s.level == 2 else ""
            st.markdown(f"{prefix}- **{s.heading}** (level {s.level})")
