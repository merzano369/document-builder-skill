"""
DOCX -> PDF separately from the build (for example, after manual edits in Word).

    python tools/export_pdf.py <report.docx> [<report.pdf>]

Word via PowerShell COM (updates TOC and fields), then LibreOffice, then docx2pdf.
Checks that the PDF was really created anew and prints the page count (if PyMuPDF is available).
"""

import os
import sys
import time

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from DocumentStyles import convert_to_pdf  # noqa: E402


def main():
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    docx = os.path.abspath(sys.argv[1])
    if not os.path.isfile(docx):
        sys.exit(f"File not found: {docx}")
    pdf = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.splitext(docx)[0] + ".pdf"
    started = time.time()
    result = convert_to_pdf(docx, pdf)
    if not result:
        sys.exit(1)
    if os.path.getmtime(result) < started - 1:
        sys.exit(f"[ERROR] PDF was not updated (it may be open in a viewer): {result}")
    try:
        try:
            import pymupdf as fitz
        except ImportError:
            import fitz
        with fitz.open(result) as d:
            print(f"Pages: {len(d)}")
    except ImportError:
        pass


if __name__ == "__main__":
    main()
