<div align="center">

<img src="docs/favicon.svg" width="72" height="72" alt="IEEE Agentic Formatter logo">

# IEEE Agentic Formatter

**Transform raw research notes, drafts, or Markdown into a strictly compliant,  
two-column IEEE Conference Paper (.docx) — in seconds.**

[![Version](https://img.shields.io/badge/version-1.1.0-4fd1c5?style=flat-square)](https://github.com/sunilgentyala/ieee-agentic-formatter/releases)
[![Python](https://img.shields.io/badge/python-3.10%2B-7c9eff?style=flat-square)](https://python.org)
[![Ollama](https://img.shields.io/badge/local-Ollama-4fd1c5?style=flat-square)](https://ollama.com)
[![Claude](https://img.shields.io/badge/cloud-claude--opus--4--8-7c9eff?style=flat-square)](https://anthropic.com)
[![License](https://img.shields.io/badge/license-MIT-9aa7b4?style=flat-square)](LICENSE)
[![Pages](https://img.shields.io/badge/site-live-4fd1c5?style=flat-square)](https://sunilgentyala.github.io/ieee-agentic-formatter/)

[**Live Site**](https://sunilgentyala.github.io/ieee-agentic-formatter/) &nbsp;|&nbsp;
[**Quick Start**](#quick-start) &nbsp;|&nbsp;
[**Local Mode**](#local-mode-no-api-key) &nbsp;|&nbsp;
[**Architecture**](#architecture) &nbsp;|&nbsp;
[**IEEE Spec**](#ieee-formatting-spec)

</div>

---

## What it does

One pipeline — from unstructured text to publication-ready IEEE paper:

1. **Paste** raw notes, a rough draft, or bullet points — or **upload** a `.txt`, `.md`, or `.docx` file
2. **AI parses** the content (local Ollama or Anthropic cloud) and extracts the full paper structure
3. **python-docx + OOXML** renders a pixel-perfect IEEE two-column layout
4. **Download** the `.docx` instantly from the browser

> **No API key required.** Run fully offline with a local Ollama model.

---

## Quick Start

```bash
# 1. Clone
git clone https://github.com/sunilgentyala/ieee-agentic-formatter.git
cd ieee-agentic-formatter

# 2. Install
pip install -r requirements.txt

# 3. Run  (no API key needed — uses local Ollama by default)
streamlit run app.py
```

Open **http://localhost:8501**, select **Local (Ollama)** in the sidebar, paste or upload your content, and click **Generate IEEE Paper**.

---

## Local Mode (No API Key)

The app defaults to **Local (Ollama)** — zero cost, fully offline, no account needed.

### Prerequisites

Install [Ollama](https://ollama.com) and pull a model:

```bash
ollama pull qwen2.5:7b    # recommended — best structured JSON output
# or
ollama pull llama3        # alternative
```

### Sidebar options

| Setting | Default | Notes |
|---------|---------|-------|
| Backend | `Local (Ollama)` | Switch to `Anthropic API` for cloud |
| Local Model | `qwen2.5:7b` | Also supports `llama3`, `mistral` |
| Ollama URL | `http://localhost:11434` | Change if Ollama runs on another port |

### How it works

`parse_raw_text_local()` sends the raw content to Ollama's OpenAI-compatible endpoint (`/v1/chat/completions`) with a strict JSON-schema system prompt. The response is parsed, validated by Pydantic v2, and fed directly into the python-docx formatter.

---

## Cloud Mode (Anthropic API)

Switch the sidebar to **Anthropic API** for `claude-opus-4-8` with adaptive thinking and forced tool use — highest quality output for complex or ambiguous inputs.

```bash
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env
streamlit run app.py
```

Or enter the key directly in the sidebar at runtime.

---

## Architecture

```
ieee-agentic-formatter/
├── agent/
│   ├── models.py        # Pydantic v2: IEEEPaper, Author, Section
│   └── text_parser.py   # parse_raw_text() — Anthropic
│                        # parse_raw_text_local() — Ollama
├── formatter/
│   └── docx_engine.py   # python-docx + OOXML IEEE layout engine
├── docs/
│   ├── index.html       # GitHub Pages site
│   ├── favicon.svg      # Address-bar logo
│   └── og-image.svg     # Social preview banner
├── app.py               # Streamlit UI (backend selector)
└── requirements.txt
```

| Module | Responsibility |
|--------|---------------|
| `agent/models.py` | Pydantic v2 schema — `IEEEPaper`, `Author`, `Section` |
| `agent/text_parser.py` | **Anthropic**: streaming + `tool_choice` forced to `structure_ieee_paper`. **Local**: Ollama OpenAI-compat endpoint + JSON-mode prompting |
| `formatter/docx_engine.py` | python-docx + `OxmlElement` injection for `w:mirrorMargins`, two-column `w:cols`, exact-leading paragraphs |
| `app.py` | Streamlit: backend radio, model selector, text paste / file upload, live preview, download |

---

## IEEE Formatting Spec

| Property | Value |
|----------|-------|
| Page size | Letter (8.5 x 11 in) |
| Margins | 1 in all sides, mirrored (odd/even pages) |
| Title | Times New Roman 24 pt, centered |
| Author block | 10 pt name / 10 pt italic affiliation / 9 pt email |
| Abstract | 9 pt bold-italic label (`Abstract—`) + 9 pt italic body |
| Keywords | 9 pt bold-italic label (`Index Terms—`) + 9 pt italic body |
| Body columns | 2 equal columns, 0.32 in gap, continuous section break |
| Body text | Times New Roman 10 pt, justified, 0.15 in first-line indent, 12 pt exact leading |
| Section headings | Level 1: `I. INTRODUCTION` centered bold; Level 2: `A. Background` left italic |
| References | 8 pt Times New Roman, hanging indent, IEEE bracket style |

---

## Stack

| Layer | Technology |
|-------|-----------|
| Local LLM | Ollama (`qwen2.5:7b`, `llama3`, `mistral`) via OpenAI-compatible API |
| Cloud LLM | `claude-opus-4-8` with adaptive thinking + `tool_choice` |
| Structured output | Pydantic v2 + JSON-mode prompting / Anthropic tool use |
| Document generation | python-docx 1.1+ with raw OOXML injection |
| UI | Streamlit 1.40+ |
| Config | python-dotenv |

---

## Releases

### v1.1.0 — Local Ollama Backend *(2026-06-26)*

- **Local mode**: run fully offline via Ollama — no API key required
- `parse_raw_text_local()`: OpenAI-compatible Ollama endpoint + JSON-schema prompting
- Streamlit sidebar: backend radio (`Local / Anthropic`), model selector, Ollama URL
- Tested with `qwen2.5:7b` and `llama3` on multi-section research drafts
- `openai>=1.0.0` added to requirements for Ollama client

### v1.0.0 — Initial Release *(2026-06-26)*

- Agentic parsing pipeline: Claude `claude-opus-4-8` + forced `tool_use` + Pydantic v2 validation
- Strict IEEE two-column `.docx` output via python-docx and OOXML
- Multi-format input: plain text, `.txt`, `.md`, `.docx` upload
- Streamlit UI with live structure preview and one-click download
- GitHub Pages site with SVG favicon and social preview banner

---

## Author

<div align="center">

### **Sunil Gentyala**

**AI/ML Researcher &nbsp;|&nbsp; IEEE Member &nbsp;|&nbsp; HCLTech (HCL America Inc.)**

*Independent researcher specialising in agentic AI systems, multi-agent security frameworks,  
federated learning, and quantum-safe protocols — with 8+ accepted IEEE publications.*

[![IEEE](https://img.shields.io/badge/IEEE-sunil.gentyala%40ieee.org-00629B?style=flat-square&logo=ieee&logoColor=white)](mailto:sunil.gentyala@ieee.org)
[![GitHub](https://img.shields.io/badge/GitHub-sunilgentyala-181717?style=flat-square&logo=github&logoColor=white)](https://github.com/sunilgentyala)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Sunil%20Gentyala-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](https://linkedin.com/in/sunilgentyala)

</div>

---

## Contributors

<table>
<tr>
  <td align="center">
    <a href="https://github.com/sunilgentyala">
      <img src="https://github.com/sunilgentyala.png" width="80" height="80" style="border-radius:50%" alt="Sunil Gentyala"><br>
      <b>Sunil Gentyala</b>
    </a><br>
    <sub>Creator &amp; Maintainer</sub>
  </td>
</tr>
</table>

---

## License

MIT &copy; 2026 Sunil Gentyala
