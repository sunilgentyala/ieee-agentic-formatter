<div align="center">

<img src="docs/favicon.svg" width="72" height="72" alt="IEEE Agentic Formatter logo">

# IEEE Agentic Formatter

**Transform raw research notes, drafts, or Markdown into a strictly compliant,  
two-column IEEE Conference Paper (.docx) — in seconds.**

[![Version](https://img.shields.io/badge/version-1.0.0-4fd1c5?style=flat-square)](https://github.com/sunilgentyala/ieee-agentic-formatter/releases)
[![Python](https://img.shields.io/badge/python-3.10%2B-7c9eff?style=flat-square)](https://python.org)
[![Model](https://img.shields.io/badge/claude-opus--4--8-4fd1c5?style=flat-square)](https://anthropic.com)
[![License](https://img.shields.io/badge/license-MIT-9aa7b4?style=flat-square)](LICENSE)
[![Pages](https://img.shields.io/badge/site-live-4fd1c5?style=flat-square)](https://sunilgentyala.github.io/ieee-agentic-formatter/)

[**Live Site**](https://sunilgentyala.github.io/ieee-agentic-formatter/) &nbsp;|&nbsp;
[**Quick Start**](#quick-start) &nbsp;|&nbsp;
[**Architecture**](#architecture) &nbsp;|&nbsp;
[**IEEE Spec**](#ieee-formatting-spec)

</div>

---

## What it does

One pipeline — from unstructured text to publication-ready IEEE paper:

1. **Paste** raw notes, a rough draft, or bullet points — or **upload** a `.txt`, `.md`, or `.docx` file
2. **Claude** (`claude-opus-4-8`, adaptive thinking, forced tool use) extracts and infers the full paper structure
3. **python-docx + OOXML** renders a pixel-perfect IEEE two-column layout
4. **Download** the `.docx` instantly from the browser

---

## Quick Start

```bash
# 1. Clone
git clone https://github.com/sunilgentyala/ieee-agentic-formatter.git
cd ieee-agentic-formatter

# 2. Install
pip install -r requirements.txt

# 3. Configure
echo "ANTHROPIC_API_KEY=sk-ant-..." > .env

# 4. Run
streamlit run app.py
```

Open **http://localhost:8501**, paste or upload content, click **Generate IEEE Paper**, download your `.docx`.

---

## Architecture

```
ieee-agentic-formatter/
├── agent/
│   ├── models.py        # Pydantic v2: IEEEPaper, Author, Section
│   └── text_parser.py   # Claude streaming + forced tool_use extraction
├── formatter/
│   └── docx_engine.py   # python-docx + raw OOXML IEEE layout engine
├── docs/
│   ├── index.html       # GitHub Pages site
│   └── favicon.svg      # Address-bar logo
├── app.py               # Streamlit UI
└── requirements.txt
```

| Module | Responsibility |
|--------|---------------|
| `agent/models.py` | Pydantic v2 schema — `IEEEPaper`, `Author`, `Section` with field-level descriptions fed to the tool schema |
| `agent/text_parser.py` | Anthropic streaming API, `tool_choice` forced to `structure_ieee_paper`, adaptive thinking enabled |
| `formatter/docx_engine.py` | python-docx + `OxmlElement` injection for `w:mirrorMargins`, two-column `w:cols`, exact-leading paragraphs |
| `app.py` | Streamlit: text paste, `.txt` / `.md` / `.docx` upload, live structure preview, in-browser download |

---

## IEEE Formatting Spec

| Property | Value |
|----------|-------|
| Page size | Letter (8.5 × 11 in) |
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
| LLM | `claude-opus-4-8` with adaptive thinking + `tool_choice` |
| Structured output | Pydantic v2 + Anthropic tool use |
| Document generation | python-docx 1.1+ with raw OOXML injection |
| UI | Streamlit 1.40+ |
| Config | python-dotenv |

---

## Releases

### v1.0.0 — Initial Release *(2026-06-26)*

- Agentic parsing pipeline: Claude `claude-opus-4-8` + forced `tool_use` + Pydantic v2 validation
- Strict IEEE two-column `.docx` output via python-docx and OOXML
- Multi-format input: plain text, `.txt`, `.md`, `.docx` upload
- Streamlit UI with live structure preview and one-click download
- GitHub Pages site with SVG favicon

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
