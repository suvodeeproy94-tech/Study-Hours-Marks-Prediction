"""
Render each PDF page as a PNG image for visual quality checks.
"""

from pathlib import Path

import fitz


PROJECT_ROOT = Path(__file__).resolve().parent.parent
PDF_FILE = PROJECT_ROOT / "Study_Hours_Marks_Prediction_IEEE_Preprint_v1.0.pdf"
OUTPUT_FOLDER = PROJECT_ROOT / ".zenodo_pdf_review"


def render_pages():
    """Render the PDF at a clear review size."""

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    pdf_document = fitz.open(PDF_FILE)

    for page_number, page in enumerate(pdf_document, start=1):
        image = page.get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
        output_file = OUTPUT_FOLDER / f"page-{page_number}.png"
        image.save(output_file)
        print(output_file)

    print(f"pages={pdf_document.page_count}")


if __name__ == "__main__":
    render_pages()
