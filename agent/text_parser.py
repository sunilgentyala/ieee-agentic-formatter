"""LLM-powered parser — supports Anthropic API and local Ollama backends."""

import json
import os
import re
from typing import Optional

from dotenv import load_dotenv

from .models import Author, IEEEPaper, Section

load_dotenv()

# ---------------------------------------------------------------------------
# Shared schema / prompts
# ---------------------------------------------------------------------------

_IEEE_TOOL = {
    "name": "structure_ieee_paper",
    "description": (
        "Extract and structure the content of a research paper into the canonical "
        "IEEE conference paper format. Infer and fill any missing sections from "
        "context. Generate Roman-numeral main sections and letter-keyed subsections."
    ),
    "input_schema": {
        "type": "object",
        "properties": {
            "title": {"type": "string", "description": "Full paper title in title case"},
            "authors": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "name": {"type": "string"},
                        "affiliation": {"type": "string"},
                        "email": {"type": "string"},
                    },
                    "required": ["name", "affiliation"],
                },
                "description": "List of authors with affiliation and optional email",
            },
            "abstract": {
                "type": "string",
                "description": "150-250 word abstract without the 'Abstract' label",
            },
            "keywords": {
                "type": "array",
                "items": {"type": "string"},
                "description": "4-6 index terms / keywords",
            },
            "sections": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "heading": {
                            "type": "string",
                            "description": "Heading text without numbering prefix",
                        },
                        "level": {
                            "type": "integer",
                            "enum": [1, 2],
                            "description": "1=main section (Roman numeral), 2=subsection (letter)",
                        },
                        "content": {
                            "type": "string",
                            "description": "Section body text; separate paragraphs with \\n\\n",
                        },
                    },
                    "required": ["heading", "level", "content"],
                },
                "description": (
                    "Ordered list of sections. Typical IEEE order: Introduction, "
                    "Related Work, Methodology, Results, Conclusion."
                ),
            },
            "references": {
                "type": "array",
                "items": {"type": "string"},
                "description": "IEEE-formatted reference strings (e.g. 'A. Smith, ..., 2024.')",
            },
        },
        "required": ["title", "authors", "abstract", "keywords", "sections", "references"],
    },
}

_ANTHROPIC_SYSTEM = (
    "You are an expert IEEE conference paper editor. "
    "When given raw text, research notes, or poorly formatted content, you extract "
    "and infer the complete paper structure. You MUST call the structure_ieee_paper "
    "tool with a fully populated result. Infer missing fields from context; never "
    "leave required fields empty. Produce publication-ready academic prose."
)

_LOCAL_SYSTEM = """You are an expert IEEE conference paper editor.
Extract and structure the given content into a complete IEEE conference paper.
Infer missing fields from context. Never leave required fields empty.

Respond with ONLY a valid JSON object — no markdown fences, no explanation — matching this schema exactly:
{
  "title": "string",
  "authors": [{"name": "string", "affiliation": "string", "email": "string"}],
  "abstract": "string (150-250 words, no label)",
  "keywords": ["string", ...],
  "sections": [
    {"heading": "string (no number prefix)", "level": 1, "content": "string (paragraphs separated by \\n\\n)"},
    ...
  ],
  "references": ["IEEE-formatted string", ...]
}
level 1 = main section (Roman numeral heading), level 2 = subsection (letter heading).
Typical section order: Introduction, Related Work, Methodology/System Design, Results/Evaluation, Conclusion."""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _build_paper(data: dict) -> IEEEPaper:
    authors = [
        Author(
            name=a["name"],
            affiliation=a["affiliation"],
            email=a.get("email", ""),
        )
        for a in data["authors"]
    ]
    sections = [
        Section(
            heading=s["heading"],
            level=s["level"],
            content=s["content"],
        )
        for s in data["sections"]
    ]
    return IEEEPaper(
        title=data["title"],
        authors=authors,
        abstract=data["abstract"],
        keywords=data["keywords"],
        sections=sections,
        references=data["references"],
    )


def _extract_json(text: str) -> dict:
    """Strip markdown fences and parse the first JSON object found."""
    text = re.sub(r"^```(?:json)?\s*", "", text.strip(), flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text.strip())
    # Find the outermost JSON object
    start = text.find("{")
    if start == -1:
        raise ValueError("No JSON object found in model response.")
    depth = 0
    for i, ch in enumerate(text[start:], start):
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                return json.loads(text[start : i + 1])
    raise ValueError("Malformed JSON in model response.")


# ---------------------------------------------------------------------------
# Anthropic backend
# ---------------------------------------------------------------------------

def parse_raw_text(raw_text: str, client=None) -> IEEEPaper:
    """Parse via Anthropic claude-opus-4-8 with forced tool use."""
    import anthropic as _anthropic

    if client is None:
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY not set. Use local mode or add key to .env.")
        client = _anthropic.Anthropic(api_key=api_key)

    with client.messages.stream(
        model="claude-opus-4-8",
        max_tokens=8192,
        thinking={"type": "adaptive"},
        system=_ANTHROPIC_SYSTEM,
        tools=[_IEEE_TOOL],
        tool_choice={"type": "tool", "name": "structure_ieee_paper"},
        messages=[
            {
                "role": "user",
                "content": (
                    "Structure the following content as a complete IEEE conference paper:\n\n"
                    + raw_text
                ),
            }
        ],
    ) as stream:
        response = stream.get_final_message()

    tool_block = next((b for b in response.content if b.type == "tool_use"), None)
    if tool_block is None:
        raise RuntimeError("Claude did not return a tool_use block.")

    return _build_paper(tool_block.input)


# ---------------------------------------------------------------------------
# Local Ollama backend
# ---------------------------------------------------------------------------

def parse_raw_text_local(
    raw_text: str,
    model: str = "qwen2.5:7b",
    ollama_url: str = "http://localhost:11434",
) -> IEEEPaper:
    """Parse via a local Ollama model using JSON-mode prompting."""
    try:
        from openai import OpenAI
    except ImportError:
        raise ImportError("Run: pip install openai  (needed for Ollama's OpenAI-compatible API)")

    client = OpenAI(base_url=f"{ollama_url}/v1", api_key="ollama")

    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": _LOCAL_SYSTEM},
            {
                "role": "user",
                "content": (
                    "Structure the following content as a complete IEEE conference paper "
                    "and return ONLY the JSON object:\n\n" + raw_text
                ),
            },
        ],
        temperature=0.2,
        max_tokens=6000,
    )

    raw_json = response.choices[0].message.content
    data = _extract_json(raw_json)
    return _build_paper(data)
