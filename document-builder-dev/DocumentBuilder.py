"""
DocumentBuilder.py
Document builder: creates DocumentStyles, calls chapters from chapters/ and saves the DOCX (and PDF).

Run:  python DocumentBuilder.py [name.docx] [--pdf]
"""

from __future__ import annotations

import os
import sys

current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from DocumentStyles import DocumentStyles, __version__
from chapters.example_lab import render as render_lab


def build(output_filename: str = "document.docx", pdf: bool = False) -> str:
    print(f"document-builder {__version__}: building the document")
    doc = DocumentStyles()

    # For a lab report: the title page and sections in a single chapter file.
    render_lab(doc)

    # For a large document: TOC and separate chapters chapters/chapterN.py
    # doc.add_toc()
    # from chapters.chapter1 import render as render_ch1
    # render_ch1(doc)

    return doc.save(os.path.join(current_dir, output_filename), pdf=pdf)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    build(args[0] if args else "document.docx", pdf="--pdf" in sys.argv)
