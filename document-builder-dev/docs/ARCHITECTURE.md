# document-builder architecture

A document for agents and people who change the skill itself. Skill users only need `SKILL.md`.

## 1. Purpose and scope

The skill turns a sequence of `doc.add_*` calls into a DOCX/PDF report that meets academic requirements. The main value: the agent writes only the content, and the engine takes care of formatting, numbering and checks.

What the skill deliberately does **not** do: template engines (Jinja, docxtpl), a DSL on top of `add_*`, parsing whole Markdown files, class hierarchies for document types, two-pass builds. Each such proposal has already been rejected as complicating the skill without need.

## 2. File map

| File | For whom | What is inside |
|---|---|---|
| `SKILL.md` | user agent | Setup, rules, short API. Must stay short |
| `reference/config.md` | user agent | All config keys (a test checks it against the code) |
| `reference/template.md` | user agent | Template filling mode |
| `DocumentStyles.py` | engine | `DocumentStyles` class, `DEFAULT_CONFIG`, config lookup, Markdown |
| `code_comments.py` | engine | `strip_code_comments()`: pure functions, no python-docx |
| `report_pdf.py` | engine | `convert_to_pdf()`: Word COM -> LibreOffice -> docx2pdf |
| `styles_config.json` | both | Full config with all keys (a sample for `#TOOLS/`) |
| `MML2OMML.XSL` | engine | Microsoft XSLT: MathML -> Word formulas |
| `DocumentBuilder.py`, `chapters/` | examples | Builder template, reference report `example_lab.py` |
| `tools/check_report.py` | checks | Analysis of a finished PDF (PyMuPDF) |
| `tools/export_pdf.py` | checks | Standalone DOCX -> PDF conversion |
| `tests/test_build.py` | development | Fast tests without Word |
| `../document-builder-demo/` | development | Full testing with Word and comparison with the old v1 |

The modules are not a package: the lab builder adds the skill folder to `sys.path` and does `from DocumentStyles import DocumentStyles`. That is why helper module names are unique (`code_comments`, `report_pdf`), so that a file from the lab's `#TOOLS/` does not shadow them. `DocumentStyles` appends its own folder to the end of `sys.path`, so it also works when loaded by file path (`importlib`); `test_load_by_path` checks this. `DocumentStyles` re-exports `strip_code_comments` and `convert_to_pdf`.

## 3. Flow

```
DocumentStyles(config_path, doc_path)
  find_config()          explicit path -> ./#TOOLS/ -> ./ -> script folder -> skill
  _load_config()         DEFAULT_CONFIG + JSON (sections merged, title separate), warnings about unknown keys
  Document(doc_path)     empty document or template
  if not a template:     _setup_page_margins (A4, margins), _setup_default_styles, _setup_page_numbering

doc.add_*()              every method appends elements to the end of body (or to the insert_at block)
doc.save(path, pdf)      _drop_trailing_break -> write DOCX (locked -> *_new.docx)
                         -> summary: counters, warnings, check(), missing, todo
                         -> convert_to_pdf() if pdf=True
```

## 4. Key mechanisms

Sections of `DocumentStyles.py` are marked with heading comments (`# TABLES`, `# LISTINGS`, etc.); search by them, not by line numbers.

**Config: a single source.** All values live in `self.config`. Attributes like `doc.font_size_body` are not stored: `__getattr__` reads them from `self.config` via the `_CONFIG_ATTRS` table, so a change to `doc.config[...]` takes effect immediately. A new key is added in three places: `DEFAULT_CONFIG`, `styles_config.json`, `reference/config.md` (checked by `test_config_sync`). Adding a shortcut attribute is optional: in new code read `self.config["key"]`.

**Document state is determined by the document itself.** Decisions like "put a break before H1?", "avoid a duplicate break?", "start the list from 1?" are made from the previous element in `body` (`_prev_element()`, `_is_break()`, `_is_numbered_item()`), not from a "what was called last" flag. Inside `insert_at` new elements temporarily sit at the end of the document, so `_prev_element()` returns the anchor's neighbour for the first of them (`self._insert_ctx`). Explicit state remains only where it is really needed: caption counters, `_list_num_id`, `_manual_num`, `_h1_count`.

**Caption numbering.** `_number_for(kind, key)` returns the next number or the number previously reserved for `key`; `ref()` reserves a number before insertion. `_make_caption()` does not touch captions that already start with a label and a digit (compatibility with old reports) and records a warning if numbers do not increase. In `caption_numbering="chapter"` mode the counters reset at every H1.

**Lists.** Native list: for each new list a `w:num` with `startOverride=1` is created based on the abstract list of the `List Number` style (`_new_list_num_id()`). If the style is missing (template) or `numbered_mode="manual"`, the item becomes a paragraph "1)\t text" with a hanging indent (`_hanging_item()`).

**Title page.** Four paragraphs (header, title, signatures, city and year) with line breaks inside, then a section break. The title section has `vAlign="both"`, so Word itself spreads the blocks over the page height. `titlePg` removes the number from the title page; the next section is numbered from 2.

**Images.** `_find_image()` looks for the file in several folders; `_image_size_cm()` computes the real size at 96 dpi, limited by the text width and `image_max_height_cm`. A missing file gives a red placeholder and an entry in `doc.missing`.

**Listings.** Each code line is a separate paragraph with `noProof`. A short listing (up to `code_keep_together_lines`) is fully chained with `keep_with_next`; a long one chains the first three lines with the caption. `add_listing_from_file()` cuts the fragment between markers, removes comments (`code_comments.py`) and applies `dedent`.

**Template.** `clear_section()` deletes elements between headings without touching section breaks. `insert_at()` remembers the set of existing elements and on exit moves the new elements before the anchor (or after it). `replace_text()` treats all `w:t` of a paragraph as one string, so it finds text that Word split into several runs.

**Checks.** `check()` checks the DOCX (placeholders, highlighting, empty paragraphs, caption without a table, wide tables). `tools/check_report.py` checks the already laid-out PDF (separated captions, numbering, half-empty pages). These are different levels: the first sees the structure, the second sees Word's layout.

**PDF.** `report_pdf.py` tries Word via PowerShell COM: updates the TOC and fields, re-saves the DOCX and saves the PDF. The script is passed via `-EncodedCommand` (UTF-16LE base64) so that paths with Cyrillic and apostrophes do not break. Then it tries `soffice --headless`, then `docx2pdf`. If nothing works, it prints all reasons and returns `None`.

## 5. Invariants

Do not break them without an explicit decision of the skill owner:

1. The public API is backward compatible: new features are added as a parameter with a default value or a config key. Old calls (`add_code_block(code, caption)`, `_create_caption_paragraph`, `add_formatted_paragraph`, attributes `doc.font_size_body`) keep working.
2. Formatting is set by the config, not by chapter code.
3. As few new public methods as possible; first look for a parameter to an existing one.
4. No emoji in code or console output (a workspace CLAUDE.md requirement; besides, the Windows console does not always print them).
5. The build does not fail because of a missing screenshot or a failed PDF conversion: it warns (except with `strict_images`).
6. The engine is not tied to one university: everything specific to KPI lives in title page fields and the config.
7. The default format is A4; in template mode the engine does not change the template's styles, margins and headers/footers.
8. Language split: instructions, code comments and console messages are in English; everything that ends up in the document (default labels, title page strings, placeholders `[ЗАГЛУШКА: ...]`, examples in `SKILL.md` and `chapters/`) is in Ukrainian. `tools/check_report.py` searches for the Ukrainian labels and placeholders in the PDF, so they must not be translated.

## 6. How to make typical changes

**New config key:** add it to `DEFAULT_CONFIG` with a comment, to `styles_config.json` in the matching section, a row in `reference/config.md`; read it as `self.config["key"]`. `test_config_sync` will remind you if something is missing.

**New method parameter:** the default value must keep the old behaviour; update the method row in the `SKILL.md` API table.

**New `add_*` method:** create paragraphs via `_new_paragraph()` (it sets alignment, spacing and zero first-line indent), text via `_add_formatted_text()`, captions via `_make_caption()` and `_caption_paragraph()`. Do not introduce "what was last" flags: if behaviour depends on a neighbour, look at `_prev_element()`.

**New language for comment removal:** a row in `_LANGS` in `code_comments.py` (line comment, block comment, quotes, escape by doubling) and a check in `test_strip_comments`.

**New check:** if it is visible in the DOCX structure, add it to `check()`; if only in the layout, add it to `tools/check_report.py`.

## 7. python-docx and Word pitfalls

Each of them has already cost time:

- The default python-docx template is **Letter**, not A4. The format must be set explicitly (`page_width_cm`, `page_height_cm`).
- When re-saving, Word **drops redundant attributes**, e.g. `firstLine="0"` if the style has no indent anyway. A missing attribute in the DOCX after conversion is not an engine bug.
- `run.add_run("1)\t")` writes the tab as a separate `<w:tab/>`, so the text of the `w:t` elements alone does not contain it. Use `Paragraph(el, parent).text`.
- lxml proxy elements are temporary: **do not compare elements via `id()`**, because ids are reused. Keep the elements themselves in a set (as `insert_at` does).
- The order of child elements in XML matters: insert `numPr` via `get_or_add_numPr()`, put `vAlign` in `sectPr` before `titlePg`, `updateFields` in `settings` before `compat`/`rsids`, etc.
- A page break as a separate empty paragraph before H1 can spill onto the next page and produce an empty one. That is why H1 uses the `page_break_before` property and `add_page_break()` does not duplicate a break.
- After updating the TOC, Word leaves an empty field paragraph below it. It has zero height (`add_toc`), otherwise an empty page may appear.
- In a template code style, space after a paragraph is inherited from Normal; it must be zeroed between listing lines.
- Each listing line is a separate paragraph, so Word wraps a long line to the paragraph's left edge; the hanging indent fixes that.
- Patch scripts in a Bash heredoc break on backslashes and quotes. For large edits, write the script to a file.

## 8. Testing and readiness of a change

1. `python tests/test_build.py` (seconds, no Word): all tests must pass. For every fixed bug, first write a test that reproduces it.
2. `python ../document-builder-demo/run_all.py` (about 40 s, needs Word): comparison with the old version, PDF, template, edge cases. Look through `output/compare/` with your own eyes: tests do not see the layout fully.
3. Update `__version__` in `DocumentStyles.py` and the `SKILL.md` heading: patch for fixes, minor for new features.
4. If the API or keys changed, update `SKILL.md` and `reference/`.
