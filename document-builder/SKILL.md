---
name: document-builder
description: Generator of academic and technical DOCX/PDF reports (1 to 100+ pages) via simple doc.add_* calls, from scratch or by filling a ready .docx template — title page, auto-numbered figures/tables/listings with references, listings from files without comments, images at real size, native LaTeX formulas, PDF conversion and layout check. All formatting lives in styles_config.json. Use for звіт, лабораторна, практична, курсова, титулка, Word/PDF документ.
---

# Document Builder (version 2.2.1)

**Document language.** Write all document content (headings, paragraphs, captions, list items, table cells, title page) in Ukrainian unless the user explicitly asks for another language. The examples below are in Ukrainian on purpose. Keep them that way.

The `DocumentStyles.py` engine builds `.docx` via python-docx. Chapter code contains only `doc.add_*` calls; fonts, indents, captions and numbering are set in `styles_config.json`.

References (open when needed):
- `reference/config.md`: all `styles_config.json` keys and title page fields;
- `reference/template.md`: filling a teacher's template (`doc_path`, `clear_section`, `insert_at`, `replace_text`, `fill_table`);
- `chapters/example_lab.py`: reference structure of a lab report.

## Dependencies

Before the first build in a session, check the dependencies with the same interpreter you will build with:

```
python -c "import docx, lxml, latex2mathml, matplotlib, PIL"
```

If the command fails with `ModuleNotFoundError`, install the dependencies yourself without asking the user, then check again:

```
python -m pip install -r <skill>/requirements.txt
```

- The check is mandatory: without `lxml`, `latex2mathml`, `matplotlib` or `Pillow` the engine does not fail but silently degrades the document (formulas, image sizes).
- `tools/check_report.py` also needs `pymupdf`: if it prints "PyMuPDF is required", run `python -m pip install pymupdf`.
- If `pip` refuses because of permissions or "externally-managed-environment", try `python -m pip install --user -r <skill>/requirements.txt`; if that fails too, show the user the error and the command.
- Word and LibreOffice cannot be installed via pip. If the PDF was not created, deliver the DOCX and tell the user that one of them is needed for PDF.

## Using it in a lab

Carry only `#TOOLS/styles_config.json` between labs. Take the engine from the skill; do not copy `DocumentStyles.py` into `#TOOLS/`.

```python
import os, sys
TOOLS = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.dirname(TOOLS)
sys.path.insert(0, os.environ.get("DOCUMENT_BUILDER_PATH") or os.path.expanduser("~/.claude/skills/document-builder"))
sys.path.insert(0, TOOLS)
from DocumentStyles import DocumentStyles
from chapters.lab import render

doc = DocumentStyles()           # finds #TOOLS/styles_config.json and prints which file was used
render(doc)
doc.save(os.path.join(LAB, "Звіт.docx"), pdf=True)
```

- Config lookup order: explicit `config_path` -> `./#TOOLS/styles_config.json` -> `./styles_config.json` -> main script folder -> skill config.
- Images are looked up in `images/` relative to the current folder, the config folder and its parent (lab root), the script folder and its parent.

## Rules

1. No raw `oxml`, `twips`, `_apply_run_font` in chapter code. Missing a feature — look for a key in `reference/config.md` or a method parameter. Whatever exists nowhere is better added to the skill than worked around in the report.
2. Do not write numbers into captions: `caption="Схема БД"` becomes "Рисунок 5 — Схема БД". The number for the text comes from `doc.ref("fig", "schema")`, even before the figure is inserted. Reference objects in the order they appear in the document.
3. A missing screenshot does not stop the build: a red placeholder appears and `save()` prints the list. Mark deliberately unfilled places with `doc.add_placeholder("що тут має бути")`.
4. Take listings from files via `add_listing_from_file`; do not copy code into Python strings.
5. Do not ignore `save()` warnings. After building the PDF, run `tools/check_report.py` and look through the contact sheet PNGs.

## API

| Method | What it does |
|---|---|
| `DocumentStyles(config_path=None, doc_path=None, verbose=True, keep_template_styles=None)` | New document, or content into the `doc_path` template |
| `add_title_page(**fields)` | Title page in its own section without a page number. Fields in `reference/config.md`, title section. Call it first |
| `add_heading(text, level=1, page_break_before=None, align=None)` | H1–H3; H1 starts on a new page, except the first one and right after a break or the title page |
| `add_paragraph(text, bold, italic, align, indent=True, space_after_pt=0, parse_markdown=True, keep_with_next=False)` | Paragraph with inline Markdown |
| `add_line(text, bold, italic, size, align="center", space_before, space_after, keep_with_next)` | Line without first-line indent |
| `add_bullet(text, align=None)` | Bulleted item |
| `add_numbered(text, align=None, restart=None)` | Numbered item; a new list starts from 1 by itself if the previous document element is not a list item |
| `add_image(name, caption, width_cm, caption_position, caption_align, max_width_cm, max_height_cm, key)` | Image at real size (96 dpi), no wider than the text and no taller than `image_max_height_cm` |
| `add_table(headers, rows, caption, col_widths_cm, caption_position, caption_align, col_align, font_size, key)` | Table; `col_align=["left","center","right"]` for data |
| `add_code_block(code, caption, key, keep_together=None, keep_with_next=False, caption_position=None)` | Listing; `keep_with_next=True` keeps it together with the figure below |
| `add_listing_from_file(path, start=None, end=None, strip_comments=True, caption=None, lang=None, **kw)` | Listing from a file: fragment from the line with `start` to the line with `end`, without comments (SQL, Python, C-like, shell, PowerShell, CSS) |
| `add_formula(latex, number=None)` | Native Word formula, number `(1.1)` on the right |
| `add_placeholder(text)` | Red placeholder `[ЗАГЛУШКА: ...]` |
| `add_toc(title="ЗМІСТ")` | TOC field; filled during `save(pdf=True)` |
| `add_page_break()` | Page break (consecutive repeats are ignored) |
| `ref(kind, key)` | Object number: `kind` = `"fig"`, `"tab"`, `"lst"` |
| `check()` | List of document problems; `save()` calls it itself |
| `save(path, pdf=False)` | DOCX (a file open in Word -> `*_new.docx`) and a summary; `pdf=True`: Word -> LibreOffice -> docx2pdf, path in `doc.pdf_path` |

After the build: `doc.missing` (missing images), `doc.todo` (placeholders), `doc.warnings`. Settings can be changed from code: `doc.config["inline_math"] = False`.

**Inline Markdown** in paragraphs, lists, cells and captions: `**bold**`, `*italic*`, `***both***`, `` `code` ``, `$x_i$` (formula). Escapes: `\*`, `\$`, `` \` ``, `\\`.

## Chapter example

```python
def render(doc):
    doc.add_title_page(number="2", discipline="Бази даних", topic="Запити SELECT", variant="10")
    doc.add_heading("Мета роботи", level=1)
    doc.add_paragraph("Навчитися ...")
    doc.add_heading("Хід роботи", level=1, page_break_before=False)
    doc.add_paragraph(f"Запит наведено в лістингу {doc.ref('lst', 'q1')}, результат на рисунку {doc.ref('fig', 'q1')}.")
    doc.add_listing_from_file("../sql/queries.sql", start="-- [Q1]", end="-- [Q2]",
                              caption="Вибірка клієнтів", key="q1", keep_with_next=True)
    doc.add_image("q1.png", caption="Результат запиту", key="q1")
    doc.add_paragraph("Запит повернув 12 рядків ...")
    doc.add_heading("Висновок", level=1, page_break_before=False)
    doc.add_paragraph("Я ...")
```

## Checking the result

```
python <skill>/tools/check_report.py Звіт.pdf [--out folder]   # contact sheet PNGs + layout problems, exit code 1 if any
python <skill>/tools/export_pdf.py Звіт.docx [Звіт.pdf]        # conversion only, prints the page count
```

`check_report.py` finds "Таблиця/Лістинг" captions on the last line of a page, gaps and duplicates in numbering, placeholders and pages filled less than a third. Put the contact sheet PNGs outside the lab root (for example, in `#TOOLS/preview/`).
