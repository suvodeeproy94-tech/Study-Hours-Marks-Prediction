"""
Create a publication PDF from the final IEEE-style Word paper.

This builder uses the verified DOCX as its content source and makes a
two-column PDF for Zenodo. The PDF can be checked without Microsoft Word or
LibreOffice.
"""

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    FrameBreak,
    Image,
    KeepTogether,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SOURCE_DOCX = PROJECT_ROOT / "Study_Hours_Marks_Prediction_IEEE_Research_Paper.docx"
OUTPUT_PDF = PROJECT_ROOT / "Study_Hours_Marks_Prediction_IEEE_Preprint_v1.0.pdf"
GRAPH_FILE = PROJECT_ROOT / "images" / "study_hours_marks_graph.png"

PAGE_WIDTH, PAGE_HEIGHT = letter
SIDE_MARGIN = 0.625 * inch
BOTTOM_MARGIN = 0.65 * inch
COLUMN_GAP = 0.25 * inch
COLUMN_WIDTH = (PAGE_WIDTH - (2 * SIDE_MARGIN) - COLUMN_GAP) / 2
FIRST_BODY_TOP = PAGE_HEIGHT - 1.70 * inch
LATER_BODY_TOP = PAGE_HEIGHT - 0.55 * inch


def draw_first_page(canvas, document):
    """Draw the full-width title and publication data on page one."""

    canvas.saveState()
    canvas.setTitle(
        "Study Hours and Marks Prediction Using Linear Regression: "
        "A Reproducible Beginner Study"
    )
    canvas.setAuthor("Suvodeep Roy")
    canvas.setSubject("Version 1.0 educational machine learning preprint")
    canvas.setKeywords(
        "educational data mining, linear regression, marks prediction, "
        "study hours, synthetic dataset"
    )

    canvas.setFont("Times-Roman", 17)
    canvas.drawCentredString(
        PAGE_WIDTH / 2,
        PAGE_HEIGHT - 0.58 * inch,
        "Study Hours and Marks Prediction Using",
    )
    canvas.drawCentredString(
        PAGE_WIDTH / 2,
        PAGE_HEIGHT - 0.83 * inch,
        "Linear Regression: A Reproducible Beginner Study",
    )

    canvas.setFont("Times-Roman", 11)
    canvas.drawCentredString(
        PAGE_WIDTH / 2,
        PAGE_HEIGHT - 1.08 * inch,
        "Suvodeep Roy",
    )
    canvas.setFont("Times-Italic", 9)
    canvas.drawCentredString(
        PAGE_WIDTH / 2,
        PAGE_HEIGHT - 1.25 * inch,
        "Independent Student Researcher",
    )
    canvas.setFont("Times-Bold", 8)
    canvas.drawCentredString(
        PAGE_WIDTH / 2,
        PAGE_HEIGHT - 1.42 * inch,
        "PREPRINT | Version 1.0 | June 25, 2026 | Not peer reviewed",
    )
    canvas.setFont("Times-Italic", 7.5)
    canvas.drawCentredString(
        PAGE_WIDTH / 2,
        PAGE_HEIGHT - 1.57 * inch,
        "Creative Commons Attribution 4.0 International (CC BY 4.0)",
    )
    canvas.setFont("Times-Roman", 8)
    canvas.drawRightString(
        PAGE_WIDTH - SIDE_MARGIN,
        0.35 * inch,
        "1",
    )
    canvas.restoreState()


def draw_later_page(canvas, document):
    """Draw a quiet preprint header and page number."""

    canvas.saveState()
    canvas.setFont("Times-Italic", 7.5)
    canvas.drawString(
        SIDE_MARGIN,
        PAGE_HEIGHT - 0.32 * inch,
        "PREPRINT - Study Hours and Marks Prediction",
    )
    canvas.setStrokeColor(colors.grey)
    canvas.setLineWidth(0.3)
    canvas.line(
        SIDE_MARGIN,
        PAGE_HEIGHT - 0.38 * inch,
        PAGE_WIDTH - SIDE_MARGIN,
        PAGE_HEIGHT - 0.38 * inch,
    )
    canvas.setFillColor(colors.black)
    canvas.setFont("Times-Roman", 8)
    canvas.drawRightString(
        PAGE_WIDTH - SIDE_MARGIN,
        0.35 * inch,
        str(document.page),
    )
    canvas.restoreState()


def build_styles():
    """Create compact styles for a two-column technical paper."""

    styles = getSampleStyleSheet()

    return {
        "body": ParagraphStyle(
            "PaperBody",
            parent=styles["BodyText"],
            fontName="Times-Roman",
            fontSize=9,
            leading=10.4,
            alignment=TA_JUSTIFY,
            firstLineIndent=10,
            spaceAfter=3,
            allowWidows=0,
            allowOrphans=0,
        ),
        "abstract": ParagraphStyle(
            "Abstract",
            parent=styles["BodyText"],
            fontName="Times-Italic",
            fontSize=8.2,
            leading=9.5,
            alignment=TA_JUSTIFY,
            spaceAfter=4,
        ),
        "heading": ParagraphStyle(
            "Heading",
            parent=styles["Heading2"],
            fontName="Times-Roman",
            fontSize=9.2,
            leading=10.5,
            alignment=TA_CENTER,
            spaceBefore=6,
            spaceAfter=3,
            keepWithNext=True,
        ),
        "subheading": ParagraphStyle(
            "Subheading",
            parent=styles["Heading3"],
            fontName="Times-Italic",
            fontSize=9,
            leading=10,
            alignment=TA_LEFT,
            spaceBefore=4,
            spaceAfter=1,
            keepWithNext=True,
        ),
        "equation": ParagraphStyle(
            "Equation",
            parent=styles["BodyText"],
            fontName="Times-Italic",
            fontSize=9,
            leading=10,
            alignment=TA_CENTER,
            spaceBefore=3,
            spaceAfter=3,
        ),
        "caption": ParagraphStyle(
            "Caption",
            parent=styles["BodyText"],
            fontName="Times-Roman",
            fontSize=7.2,
            leading=8,
            alignment=TA_CENTER,
            spaceBefore=3,
            spaceAfter=3,
            keepWithNext=True,
        ),
        "reference": ParagraphStyle(
            "Reference",
            parent=styles["BodyText"],
            fontName="Times-Roman",
            fontSize=7.2,
            leading=8.2,
            alignment=TA_JUSTIFY,
            leftIndent=10,
            firstLineIndent=-10,
            spaceAfter=2,
        ),
    }


def iter_document_blocks(document):
    """Yield paragraphs and tables in their real document order."""

    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield DocxParagraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield DocxTable(child, document)


def clean_text(text):
    """Escape characters that ReportLab treats as markup."""

    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("—", "&mdash;")
    )


def convert_docx_table(docx_table):
    """Convert one DOCX table into a compact PDF table."""

    values = [
        [cell.text.strip() for cell in row.cells]
        for row in docx_table.rows
    ]
    column_count = len(values[0])
    column_widths = [COLUMN_WIDTH / column_count] * column_count

    pdf_table = Table(
        values,
        colWidths=column_widths,
        repeatRows=1,
        hAlign="LEFT",
    )
    pdf_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Times-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Times-Roman"),
                ("FONTSIZE", (0, 0), (-1, -1), 6.8),
                ("LEADING", (0, 0), (-1, -1), 7.8),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("LEFTPADDING", (0, 0), (-1, -1), 2),
                ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    return pdf_table


def build_story(document):
    """Convert the DOCX body into ReportLab flow items."""

    styles = build_styles()
    story = []
    content_started = False
    graph_added = False
    reference_count = 0

    skip_text = {
        "Study Hours and Marks Prediction Using Linear Regression: "
        "A Reproducible Beginner Study",
        "Suvodeep Roy",
        "Independent Student Researcher",
        "PREPRINT | Version 1.0 | June 25, 2026 | Not peer reviewed",
        "Licensed under Creative Commons Attribution 4.0 International "
        "(CC BY 4.0)",
    }

    for block in iter_document_blocks(document):
        if isinstance(block, DocxTable):
            story.append(Spacer(1, 2))
            story.append(convert_docx_table(block))
            story.append(Spacer(1, 3))
            continue

        text = block.text.strip()

        if text in skip_text:
            continue

        if not content_started:
            if text.startswith("Abstract"):
                content_started = True
            else:
                continue

        has_picture = bool(block._p.xpath(".//w:drawing"))
        if has_picture and GRAPH_FILE.exists() and not graph_added:
            image = Image(str(GRAPH_FILE), width=COLUMN_WIDTH, height=2.15 * inch)
            image.hAlign = "CENTER"
            story.append(Spacer(1, 4))
            story.append(image)
            graph_added = True
            continue

        if not text:
            continue

        safe_text = clean_text(text)
        style_name = block.style.name if block.style else ""

        if text.startswith("Abstract") or text.startswith("Index Terms"):
            label, remaining_text = text.split("—", 1)
            story.append(
                Paragraph(
                    f"<b>{clean_text(label)}—</b>{clean_text(remaining_text)}",
                    styles["abstract"],
                )
            )
        elif style_name == "Heading 1":
            story.append(Paragraph(safe_text, styles["heading"]))
        elif style_name == "Heading 2":
            story.append(Paragraph(safe_text, styles["subheading"]))
        elif text.startswith("TABLE ") or text.startswith("Fig. "):
            story.append(Paragraph(safe_text.replace("\n", "<br/>"), styles["caption"]))
        elif text.startswith("[") and text[1:2].isdigit():
            story.append(Paragraph(safe_text, styles["reference"]))
            reference_count += 1
            if reference_count == 5:
                story.append(FrameBreak())
        elif block.alignment == 1 and text.endswith(")"):
            story.append(Paragraph(safe_text, styles["equation"]))
        else:
            story.append(Paragraph(safe_text, styles["body"]))

    return story


def build_pdf():
    """Build the final publication PDF."""

    source_document = Document(SOURCE_DOCX)

    first_frames = [
        Frame(
            SIDE_MARGIN,
            BOTTOM_MARGIN,
            COLUMN_WIDTH,
            FIRST_BODY_TOP - BOTTOM_MARGIN,
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
            id="first_left",
        ),
        Frame(
            SIDE_MARGIN + COLUMN_WIDTH + COLUMN_GAP,
            BOTTOM_MARGIN,
            COLUMN_WIDTH,
            FIRST_BODY_TOP - BOTTOM_MARGIN,
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
            id="first_right",
        ),
    ]

    later_frames = [
        Frame(
            SIDE_MARGIN,
            BOTTOM_MARGIN,
            COLUMN_WIDTH,
            LATER_BODY_TOP - BOTTOM_MARGIN,
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
            id="later_left",
        ),
        Frame(
            SIDE_MARGIN + COLUMN_WIDTH + COLUMN_GAP,
            BOTTOM_MARGIN,
            COLUMN_WIDTH,
            LATER_BODY_TOP - BOTTOM_MARGIN,
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
            id="later_right",
        ),
    ]

    pdf = BaseDocTemplate(
        str(OUTPUT_PDF),
        pagesize=letter,
        leftMargin=SIDE_MARGIN,
        rightMargin=SIDE_MARGIN,
        topMargin=0.55 * inch,
        bottomMargin=BOTTOM_MARGIN,
        title=(
            "Study Hours and Marks Prediction Using Linear Regression: "
            "A Reproducible Beginner Study"
        ),
        author="Suvodeep Roy",
        subject="Version 1.0 educational machine learning preprint",
    )

    pdf.addPageTemplates(
        [
            PageTemplate(
                id="FirstPage",
                frames=first_frames,
                onPage=draw_first_page,
                autoNextPageTemplate="LaterPages",
            ),
            PageTemplate(
                id="LaterPages",
                frames=later_frames,
                onPage=draw_later_page,
            ),
        ]
    )

    pdf.build(build_story(source_document))
    print(OUTPUT_PDF)


if __name__ == "__main__":
    build_pdf()
