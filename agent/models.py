"""Pydantic data models for IEEE paper structure."""

from typing import List, Optional
from pydantic import BaseModel, Field


class Author(BaseModel):
    name: str
    affiliation: str
    email: str = ""


class Section(BaseModel):
    heading: str = Field(description="Section heading without numbering prefix")
    level: int = Field(description="1 for main section (Roman numeral), 2 for subsection (letter)")
    content: str = Field(description="Section prose, paragraphs separated by double newlines")


class IEEEPaper(BaseModel):
    title: str
    authors: List[Author]
    abstract: str
    keywords: List[str]
    sections: List[Section]
    references: List[str]
