"""Generates the committed .docx fixture corpus for the ooxml/export.js test suite.

Run: pip install -r requirements.txt && python3 fixtures/generate.py
Each fixture is documented at its call site below with what it exercises and why.
The M1 export tests read the committed .docx files directly -- this script is a
generator, not a build step, so its output must be regenerated and re-committed
whenever a fixture's construction changes.
"""

import os

from docx import Document

FIXTURES_DIR = os.path.dirname(os.path.abspath(__file__))


def plain_paragraphs():
    """Baseline fixture: a few ordinary paragraphs, no tracked changes, no
    comments, no special formatting. Proves the export pipeline handles the
    simplest possible document and gives a byte-identical-to-reference-impl
    baseline to diff every other fixture against."""
    doc = Document()
    doc.add_paragraph("This is the first paragraph of a plain document.")
    doc.add_paragraph("This is the second paragraph, with nothing tracked.")
    doc.add_paragraph(
        "A third paragraph exists so table/list neighbors in later fixtures "
        "have a plain-paragraph baseline to compare against."
    )
    doc.save(os.path.join(FIXTURES_DIR, "plain-paragraphs.docx"))


FIXTURES = [
    plain_paragraphs,
]


def main():
    for fixture in FIXTURES:
        fixture()
        print(f"generated: {fixture.__name__}")


if __name__ == "__main__":
    main()
