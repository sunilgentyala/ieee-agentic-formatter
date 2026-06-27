"""Streamlit interface for the IEEE Agentic Formatter."""

import os
from io import BytesIO, StringIO

import streamlit as st
from docx import Document as DocxDocument
from dotenv import load_dotenv

from agent.text_parser import parse_raw_text
from formatter.docx_engine import generate_ieee_docx

load_dotenv()

st.set_page_config(
    page_title="IEEE Agentic Formatter",
    page_icon="📄",
    layout="wide",
)

st.title("IEEE Agentic Formatter")
st.caption(
    "Upload raw text, research notes, or a Markdown file and receive a "
    "strictly compliant IEEE conference paper in .docx format."
)

# --- Sidebar ---
with st.sidebar:
    st.header("Configuration")
    api_key_input = st.text_input(
        "Anthropic API Key",
        type="password",
        value=os.getenv("ANTHROPIC_API_KEY", ""),
        help="Required if ANTHROPIC_API_KEY is not set in your .env file.",
    )
    st.divider()
    st.markdown(
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
            "— anything. Claude will structure it into a full IEEE paper."
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
    effective_key = api_key_input.strip() or os.getenv("ANTHROPIC_API_KEY", "")
    if not effective_key:
        st.error("Please provide an Anthropic API key in the sidebar or in your .env file.")
        st.stop()

    import anthropic

    client = anthropic.Anthropic(api_key=effective_key)

    col1, col2 = st.columns(2)

    with st.spinner("Step 1/2 — Claude is parsing and structuring your content..."):
        try:
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
