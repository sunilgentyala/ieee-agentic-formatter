"""python-docx engine that produces a strictly IEEE-compliant two-column .docx."""

import re
from io import BytesIO
from typing import Optional

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

from agent.models import IEEEPaper

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_ROMAN = [
    (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
    (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
    (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
]


def _to_roman(num: int) -> str:
    result = ""
    for value, numeral in _ROMAN:
        while num >= value:
            result += numeral
            num -= value
    return result


def _to_letter(num: int) -> str:
    return chr(ord("A") + num - 1)


def _set_font(run, name: str, size_pt: float, bold: bool = False, italic: bool = False):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic


def _para_spacing(para, space_before: float = 0, space_after: float = 0):
    pPr = para._p.get_or_add_pPr()
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), str(int(space_before * 20)))
    spacing.set(qn("w:after"), str(int(space_after * 20)))
    pPr.append(spacing)


def _set_line_spacing_exact(para, pts: float):
    pPr = para._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:line"), str(int(pts * 20)))
    spacing.set(qn("w:lineRule"), "exact")


def _inject_cols(section_element, num_cols: int, space_twips: int = 460):
    """Inject w:cols into a section's sectPr."""
    sectPr = section_element._sectPr
    cols = OxmlElement("w:cols")
    cols.set(qn("w:num"), str(num_cols))
    cols.set(qn("w:space"), str(space_twips))
    cols.set(qn("w:equalWidth"), "1")
    sectPr.append(cols)


def _inject_mirror_margins(doc: Document):
    """Add w:mirrorMargins to the document-level settings."""
    settings_part = doc.settings.element
    mirror = OxmlElement("w:mirrorMargins")
    settings_part.append(mirror)


def _set_page_margins(section, top=1.0, bottom=1.0, left=1.0, right=1.0):
    section.top_margin = Inches(top)
    section.bottom_margin = Inches(bottom)
    section.left_margin = Inches(left)
    section.right_margin = Inches(right)


def _set_page_size_letter(section):
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)


# ---------------------------------------------------------------------------
# Section 1 — single-column title block
# ---------------------------------------------------------------------------

def _add_title(doc: Document, title: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _para_spacing(p, space_before=0, space_after=6)
    run = p.add_run(title)
    _set_font(run, "Times New Roman", 24, bold=False)


def _add_author_block(doc: Document, paper: IEEEPaper):
    """Render authors as a centred block: name / affiliation / email."""
    for author in paper.authors:
        p_name = doc.add_paragraph()
        p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _para_spacing(p_name, space_before=0, space_after=0)
        run = p_name.add_run(author.name)
        _set_font(run, "Times New Roman", 10, bold=False)

        p_aff = doc.add_paragraph()
        p_aff.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _para_spacing(p_aff, space_before=0, space_after=0)
        run = p_aff.add_run(author.affiliation)
        _set_font(run, "Times New Roman", 10, italic=True)

        if author.email:
            p_email = doc.add_paragraph()
            p_email.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _para_spacing(p_email, space_before=0, space_after=4)
            run = p_email.add_run(author.email)
            _set_font(run, "Times New Roman", 9)


def _add_abstract_block(doc: Document, abstract: str, keywords: list):
    """Render Abstract and Keywords in 9pt bold-italic label style."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _para_spacing(p, space_before=6, space_after=0)

    label = p.add_run("Abstract—")
    _set_font(label, "Times New Roman", 9, bold=True, italic=True)

    body = p.add_run(abstract)
    _set_font(body, "Times New Roman", 9, italic=True)

    kw_text = ", ".join(keywords)
    pk = doc.add_paragraph()
    pk.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _para_spacing(pk, space_before=4, space_after=6)

    kw_label = pk.add_run("Index Terms—")
    _set_font(kw_label, "Times New Roman", 9, bold=True, italic=True)

    kw_body = pk.add_run(kw_text)
    _set_font(kw_body, "Times New Roman", 9, italic=True)


# ---------------------------------------------------------------------------
# Section 2 — two-column body
# ---------------------------------------------------------------------------

def _add_section_heading(doc: Document, text: str, level: int, counter: int):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    _para_spacing(p, space_before=6, space_after=3)

    if level == 1:
        label = f"{_to_roman(counter)}. {text.upper()}"
        run = p.add_run(label)
        _set_font(run, "Times New Roman", 10, bold=True)
    else:
        label = f"{_to_letter(counter)}. {text}"
        run = p.add_run(label)
        _set_font(run, "Times New Roman", 10, italic=True)


def _add_body_paragraph(doc: Document, text: str, first_indent: bool = True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _para_spacing(p, space_before=0, space_after=0)
    _set_line_spacing_exact(p, 12)

    if first_indent:
        pPr = p._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:firstLine"), str(int(0.15 * 1440)))
        pPr.append(ind)

    run = p.add_run(text.strip())
    _set_font(run, "Times New Roman", 10)


def _add_references(doc: Document, references: list):
    heading_p = doc.add_paragraph()
    heading_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _para_spacing(heading_p, space_before=6, space_after=3)
    run = heading_p.add_run("REFERENCES")
    _set_font(run, "Times New Roman", 10, bold=True)

    for idx, ref in enumerate(references, start=1):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _para_spacing(p, space_before=0, space_after=1)
        _set_line_spacing_exact(p, 10)

        pPr = p._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(int(0.25 * 1440)))
        ind.set(qn("w:hanging"), str(int(0.25 * 1440)))
        pPr.append(ind)

        num_run = p.add_run(f"[{idx}] ")
        _set_font(num_run, "Times New Roman", 8, bold=True)

        ref_run = p.add_run(ref)
        _set_font(ref_run, "Times New Roman", 8)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def generate_ieee_docx(paper: IEEEPaper) -> bytes:
    """Return the bytes of a complete IEEE-formatted .docx for the given paper."""
    doc = Document()

    # --- Page setup (first / only section at this point = single column) ---
    sec1 = doc.sections[0]
    _set_page_size_letter(sec1)
    _set_page_margins(sec1, top=1.0, bottom=1.0, left=1.0, right=1.0)

    _inject_mirror_margins(doc)

    # Remove default paragraph spacing for the whole document
    doc.styles["Normal"].paragraph_format.space_before = Pt(0)
    doc.styles["Normal"].paragraph_format.space_after = Pt(0)

    # --- Title block (single column) ---
    _add_title(doc, paper.title)
    _add_author_block(doc, paper)
    _add_abstract_block(doc, paper.abstract, paper.keywords)

    # --- Switch to two-column layout via a CONTINUOUS section break ---
    p_break = doc.add_paragraph()
    sec2 = p_break._p.get_or_add_pPr()

    # Insert sectPr with 2-col layout into the paragraph that triggers the break
    sectPr = OxmlElement("w:sectPr")
    pg_sz = OxmlElement("w:pgSz")
    pg_sz.set(qn("w:w"), str(int(8.5 * 1440)))
    pg_sz.set(qn("w:h"), str(int(11 * 1440)))
    sectPr.append(pg_sz)

    pg_mar = OxmlElement("w:pgMar")
    pg_mar.set(qn("w:top"), str(int(1.0 * 1440)))
    pg_mar.set(qn("w:right"), str(int(1.0 * 1440)))
    pg_mar.set(qn("w:bottom"), str(int(1.0 * 1440)))
    pg_mar.set(qn("w:left"), str(int(1.0 * 1440)))
    pg_mar.set(qn("w:header"), "720")
    pg_mar.set(qn("w:footer"), "720")
    sectPr.append(pg_mar)

    cols_elem = OxmlElement("w:cols")
    cols_elem.set(qn("w:num"), "1")
    sectPr.append(cols_elem)

    typ = OxmlElement("w:type")
    typ.set(qn("w:val"), "continuous")
    sectPr.append(typ)

    sec2.append(sectPr)

    # Now add body sections; the document's final sectPr will be 2-column
    main_section_counter = 0
    sub_section_counter = 0

    for section in paper.sections:
        if section.level == 1:
            main_section_counter += 1
            sub_section_counter = 0
            _add_section_heading(doc, section.heading, level=1, counter=main_section_counter)
        else:
            sub_section_counter += 1
            _add_section_heading(doc, section.heading, level=2, counter=sub_section_counter)

        paragraphs = re.split(r"\n\n+", section.content.strip())
        for i, para_text in enumerate(paragraphs):
            if para_text.strip():
                _add_body_paragraph(doc, para_text, first_indent=(i == 0))

    _add_references(doc, paper.references)

    # --- Apply 2-column layout to the document-level sectPr ---
    body = doc.element.body
    final_sectPr = body.find(qn("w:sectPr"))
    if final_sectPr is None:
        final_sectPr = OxmlElement("w:sectPr")
        body.append(final_sectPr)

    existing_cols = final_sectPr.find(qn("w:cols"))
    if existing_cols is not None:
        final_sectPr.remove(existing_cols)

    two_cols = OxmlElement("w:cols")
    two_cols.set(qn("w:num"), "2")
    two_cols.set(qn("w:space"), "460")
    two_cols.set(qn("w:equalWidth"), "1")
    final_sectPr.append(two_cols)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
