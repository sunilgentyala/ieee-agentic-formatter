"""LLM-powered parser that extracts IEEE paper structure from raw text."""

import json
import os
from typing import Optional

import anthropic
from dotenv import load_dotenv

from .models import Author, IEEEPaper, Section

load_dotenv()

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

_SYSTEM_PROMPT = (
    "You are an expert IEEE conference paper editor. "
    "When given raw text, research notes, or poorly formatted content, you extract "
    "and infer the complete paper structure. You MUST call the structure_ieee_paper "
    "tool with a fully populated result. Infer missing fields from context; never "
    "leave required fields empty. Produce publication-ready academic prose."
)


def parse_raw_text(raw_text: str, client: Optional[anthropic.Anthropic] = None) -> IEEEPaper:
    """Parse raw text or research notes into a validated IEEEPaper using Claude."""
    if client is None:
        api_key = os.getenv("ANTHROPIC_API_KEY")
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY not set. Add it to .env or pass a client.")
        client = anthropic.Anthropic(api_key=api_key)

    with client.messages.stream(
        model="claude-opus-4-8",
        max_tokens=8192,
        thinking={"type": "adaptive"},
        system=_SYSTEM_PROMPT,
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

    tool_use_block = next(
        (b for b in response.content if b.type == "tool_use"), None
    )
    if tool_use_block is None:
        raise RuntimeError("Claude did not return a tool_use block. Response: " + str(response))

    data = tool_use_block.input

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
