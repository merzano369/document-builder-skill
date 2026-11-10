"""
report_pdf.py
DOCX -> PDF conversion: Word via PowerShell COM (updates TOC and fields), then LibreOffice,
then docx2pdf. Does not depend on DocumentStyles.
"""

from __future__ import annotations

import base64
import os
import shutil
import subprocess
from typing import Optional, Tuple

# ------------------------------------------------------------------------------
# DOCX -> PDF conversion
# ------------------------------------------------------------------------------

def _pdf_via_word(docx: str, pdf: str) -> Tuple[bool, str]:
    if os.name != "nt" or not shutil.which("powershell"):
        return False, "Word COM is available only on Windows"
    q = lambda s: s.replace("'", "''")
    script = (
        "$ErrorActionPreference='Stop'; $w=$null; "
        "try { $w=New-Object -ComObject Word.Application; $w.Visible=$false; $w.DisplayAlerts=0; "
        f"$d=$w.Documents.Open('{q(docx)}', $false, $false); "
        "foreach($t in $d.TablesOfContents){ $t.Update() | Out-Null }; "
        "$d.Fields.Update() | Out-Null; "
        "foreach($t in $d.TablesOfContents){ $t.UpdatePageNumbers() | Out-Null }; "
        "$d.Save(); "
        f"$d.SaveAs2('{q(pdf)}', 17); $d.Close($false) "
        "} finally { if($w){ $w.Quit() } }"
    )
    enc = base64.b64encode(script.encode("utf-16-le")).decode("ascii")
    try:
        res = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", enc],
                             capture_output=True, text=True, timeout=300)
    except Exception as e:
        return False, str(e)
    if res.returncode == 0 and os.path.isfile(pdf):
        return True, ""
    return False, (res.stderr or res.stdout or "unknown error").strip().splitlines()[0][:200]


def _pdf_via_soffice(docx: str, pdf: str) -> Tuple[bool, str]:
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return False, "soffice not found in PATH"
    out_dir = os.path.dirname(os.path.abspath(pdf))
    try:
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx],
                       capture_output=True, timeout=300)
    except Exception as e:
        return False, str(e)
    produced = os.path.join(out_dir, os.path.splitext(os.path.basename(docx))[0] + ".pdf")
    if os.path.isfile(produced):
        if os.path.abspath(produced) != os.path.abspath(pdf):
            shutil.move(produced, pdf)
        return True, "TOC is not updated automatically"
    return False, "soffice did not create a PDF"


def _pdf_via_docx2pdf(docx: str, pdf: str) -> Tuple[bool, str]:
    try:
        from docx2pdf import convert
    except ImportError:
        return False, "docx2pdf package is not installed"
    try:
        convert(docx, pdf)
        return os.path.isfile(pdf), ""
    except Exception as e:
        return False, str(e)


def convert_to_pdf(docx_path: str, pdf_path: Optional[str] = None) -> Optional[str]:
    """DOCX -> PDF: Word via PowerShell COM (with field update), then soffice, then docx2pdf."""
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path or os.path.splitext(docx_path)[0] + ".pdf")
    errors = []
    for name, fn in (("Word", _pdf_via_word), ("LibreOffice", _pdf_via_soffice), ("docx2pdf", _pdf_via_docx2pdf)):
        ok, msg = fn(docx_path, pdf_path)
        if ok:
            print(f"[OK] PDF ({name}): {pdf_path}" + (f" [{msg}]" if msg else ""))
            return pdf_path
        errors.append(f"{name}: {msg}")
    print("[ERROR] PDF was not created. Attempts:\n  " + "\n  ".join(errors))
    print("        Install Microsoft Word or LibreOffice (soffice in PATH), or: pip install docx2pdf")
    return None
