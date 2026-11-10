# Filling a .docx template

When the teacher provides a ready template, the engine adds content to it without touching the formatting.

## Template mode

```python
doc = DocumentStyles(doc_path=os.path.join(LAB, "#INPUT", "шаблон.docx"))
```

- By default `keep_template_styles=True`: styles, margins, page size, headers/footers and page numbering come from the template. The config affects only new elements (line font, captions, tables).
- `keep_template_styles=False` applies the config styles on top of the template.
- Text width for figures and tables is computed from the template margins.
- If the template has no `List Bullet` / `List Number` styles, items become paragraphs with «–» or «1)» and a hanging indent. If there is no `Heading N`, a heading becomes a paragraph with an outline level, so it still gets into the TOC.
- The template listing style is set by the `code_style` key (e.g. `"Код исх"`). The engine zeroes the spacing between code lines itself. If the style is missing, `code_font_name`/`code_font_size` are used and a warning is printed.

## Methods

| Method | What it does |
|---|---|
| `find_paragraph(text, style_prefix=None, exact=False)` | First body paragraph with this text; `style_prefix="Heading"` searches only among headings. Not found: `LookupError` |
| `clear_section(start, end=None)` | Deletes content between headings (the headings stay). Without `end` it goes to the next heading of the same or higher level. Does not delete section breaks. Returns the boundary paragraph, or `None` if the document ends there |
| `insert_at(anchor, after=False)` | Context manager: all `add_*` inside are inserted before `anchor` (or after it) instead of at the end. `anchor`: text, paragraph, table or `None` (end of document) |
| `replace_text(old, new)` | Replacement in the body, tables, text boxes and headers/footers, including when Word split the text into several runs. Returns the number of replacements |
| `fill_table(table, rows, start_row=1, clear_highlight=True, font_size=None, col_align=None)` | Fills an existing table (object or index from 0); missing rows are copied from the last one; removes highlighting |
| `check()` | Finds leftover placeholders (`template_placeholders`), highlighted text, consecutive empty paragraphs, a caption without a table, tables wider than the text area. `save()` calls it itself |

## Example

```python
doc = DocumentStyles(doc_path=os.path.join(LAB, "#INPUT", "шаблон.docx"))
doc.config["code_style"] = "Код исх"
doc.replace_text("{ПІБ}", "Іваненко І. І.")
with doc.insert_at(doc.clear_section("Хід роботи")):
    doc.add_paragraph(f"Код наведено в лістингу {doc.ref('lst', 'main')}.")
    doc.add_listing_from_file("../src/main.py", caption="Основний модуль", key="main")
doc.fill_table(0, [["Час", "12 мс"], ["Пам'ять", "3 МБ"]])
doc.save(os.path.join(LAB, "Звіт.docx"), pdf=True)   # check() warns about leftover placeholders and highlighting
```

## Limits

- Do not add a title page (`add_title_page`) or a TOC if the template already has them; `add_title_page` inside `insert_at` is not supported.
- The first `add_numbered` inside `insert_at` continues the list if the anchor is preceded by an item of the same list, and starts a new one otherwise.
