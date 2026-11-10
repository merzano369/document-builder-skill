"""
Check of a finished PDF report (requires PyMuPDF: pip install pymupdf).

    python tools/check_report.py <report.pdf> [--out <folder for PNG>] [--cols 4]

What it does:
  1. Builds PNG contact sheets (8 pages per sheet) for a quick look at the layout.
  2. Finds "Таблиця"/"Лістинг" captions left as the last line of a page.
  3. Checks figure, table and listing numbering (gaps, duplicates).
  4. Finds placeholders: [ЗАГЛУШКА..., [ЗОБРАЖЕННЯ ВІДСУТНЄ..., [СКРИНШОТ..., TODO.
  5. Finds pages filled less than a third (except the first and the last).
Exit code 1 if problems are found.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict

try:
    import pymupdf as fitz
except ImportError:
    try:
        import fitz
    except ImportError:
        sys.exit("PyMuPDF is required: pip install pymupdf")

CAPTION_RE = re.compile(r"^(Рисунок|Таблиця|Лістинг)\s+(\d+(?:\.\d+)?)\s*[—–-]\s")
PLACEHOLDER_RE = re.compile(r"\[(ЗАГЛУШКА|ЗОБРАЖЕННЯ ВІДСУТНЄ|СКРИНШОТ|TODO)[^\]]*\]?|\bTODO\b|Зображення не знайдено")
PAGE_NUM_RE = re.compile(r"^\d{1,3}$")


def page_lines(page):
    """Page text lines top to bottom: (y0, y1, text). The page number is dropped."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            text = "".join(s["text"] for s in l["spans"]).strip()
            if text:
                out.append((l["bbox"][1], l["bbox"][3], text))
    out.sort()
    if out and PAGE_NUM_RE.match(out[-1][2]) and out[-1][0] > page.rect.height * 0.9:
        out.pop()
    return out


def contact_sheets(doc, out_dir, cols=4, rows=2, zoom=0.35):
    from PIL import Image
    os.makedirs(out_dir, exist_ok=True)
    per = cols * rows
    paths = []
    for start in range(0, len(doc), per):
        thumbs = []
        for i in range(start, min(start + per, len(doc))):
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            thumbs.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
        w, h = thumbs[0].size
        sheet = Image.new("RGB", (cols * (w + 10) + 10, rows * (h + 10) + 10), "#888888")
        for k, im in enumerate(thumbs):
            sheet.paste(im, (10 + (k % cols) * (w + 10), 10 + (k // cols) * (h + 10)))
        p = os.path.join(out_dir, f"sheet_{start // per + 1:02d}.png")
        sheet.save(p)
        paths.append(p)
    return paths


def check(pdf_path, out_dir=None, cols=4):
    doc = fitz.open(pdf_path)
    problems = []
    numbers = defaultdict(list)

    for idx, page in enumerate(doc):
        pno = idx + 1
        lines = page_lines(page)
        for _, _, text in lines:
            m = CAPTION_RE.match(text)
            if m:
                numbers[m.group(1)].append((m.group(2), pno))
            if PLACEHOLDER_RE.search(text):
                problems.append(f"p. {pno}: placeholder: {text[:80]}")
        if lines:
            last = lines[-1][2]
            if re.match(r"^(Таблиця|Лістинг)\s+\d", last):
                problems.append(f"p. {pno}: caption separated from its content (last line): {last[:60]}")
        if 0 < idx < len(doc) - 1:
            usable = page.rect.height * 0.86
            bottom = max((l[1] for l in lines), default=0)
            images = page.get_image_info()
            if images:
                bottom = max(bottom, max(i["bbox"][3] for i in images))
            if bottom - page.rect.height * 0.07 < usable / 3:
                problems.append(f"p. {pno}: page is filled less than a third")

    for kind, items in numbers.items():
        seen = {}
        for num, pno in items:
            if num in seen:
                problems.append(f"{kind} {num}: duplicate (p. {seen[num]} and p. {pno})")
            seen[num] = pno
        groups = defaultdict(list)
        for num in seen:
            head, _, tail = num.rpartition(".")
            groups[head].append(int(tail))
        for head, ns in groups.items():
            ns.sort()
            expected = list(range(1, ns[-1] + 1))
            missing = sorted(set(expected) - set(ns))
            if missing:
                pre = f"{head}." if head else ""
                problems.append(f"{kind}: missing numbers " + ", ".join(pre + str(m) for m in missing))
            if [int(n.rpartition(".")[2]) for n, _ in items if n.rpartition(".")[0] == head] != sorted(ns):
                problems.append(f"{kind}: numbers in chapter '{head or '-'}' are not increasing")

    print(f"File: {pdf_path} ({len(doc)} pages)")
    print("Captions: " + ", ".join(f"{k}: {len(v)}" for k, v in numbers.items()) if numbers else "No captions found")
    if out_dir:
        for p in contact_sheets(doc, out_dir, cols=cols):
            print(f"Contact sheet: {p}")
    if problems:
        print(f"Problems found: {len(problems)}")
        for p in problems:
            print("  - " + p)
        return 1
    print("No problems found")
    return 0


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="PDF report layout check")
    ap.add_argument("pdf")
    ap.add_argument("--out", help="folder for contact sheet PNGs (default <pdf>_preview)")
    ap.add_argument("--cols", type=int, default=4)
    a = ap.parse_args()
    out = a.out or os.path.splitext(a.pdf)[0] + "_preview"
    sys.exit(check(a.pdf, out, a.cols))
