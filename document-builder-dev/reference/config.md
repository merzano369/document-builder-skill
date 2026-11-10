# styles_config.json: all keys

Defaults are set by `DEFAULT_CONFIG` in `DocumentStyles.py`. Sections in the JSON exist only for readability: the engine merges all keys into one dictionary (except `title`). A missing key takes its default; an unknown key (a typo) is printed as a warning.

Settings can also be changed from code after the document is created: `doc.config["paragraph_align"] = "left"` affects all subsequent elements. The exceptions are Word styles (Normal, Heading, List) and page margins: they are set up once when the document is created.

The test `tests/test_build.py::test_config_sync` checks that every `DEFAULT_CONFIG` key is present in `styles_config.json` and in this table.

## fonts

| Key | Default | Affects |
|---|---|---|
| `font_name` | `"Times New Roman"` | Font of all text |
| `font_size_body` | `14.0` | Body text, lists, captions |
| `font_size_h1` | `16.0` | Level-1 heading |
| `font_size_h2` | `14.0` | Level-2 heading |
| `font_size_h3` | `14.0` | Level-3 heading |

## alignment

Values: `"justify"`, `"left"`, `"center"`, `"right"`.

| Key | Default | Affects |
|---|---|---|
| `paragraph_align` | `"justify"` | Paragraphs |
| `bullet_align` | `"justify"` | Bulleted lists |
| `numbered_align` | `"justify"` | Numbered lists |
| `heading_align` | `"left"` | Headings, if no level key is set |
| `heading_h1_align` | `"left"` | Level-1 heading |
| `heading_h2_align` | `"left"` | Level-2 heading |
| `heading_h3_align` | `"left"` | Level-3 heading |

## paragraph_format

| Key | Default | Affects |
|---|---|---|
| `line_spacing` | `1.15` | Line spacing of text |
| `paragraph_indent_cm` | `1.25` | First-line indent and list indent |
| `h1_page_break_before` | `true` | Every H1 starts on a new page (except the first one and right after a break or the title page) |
| `numbered_mode` | `"native"` | `native`: Word list; `manual`: text prefix «1)» |

## page_margins

In template mode (`doc_path`) these keys are not applied: margins and page size come from the template.

| Key | Default | Affects |
|---|---|---|
| `margin_top_cm` | `2.0` | Top margin |
| `margin_bottom_cm` | `2.0` | Bottom margin |
| `margin_left_cm` | `3.0` | Left margin |
| `margin_right_cm` | `1.0` | Right margin |
| `page_width_cm` | `21.0` | Page width (A4). Without it python-docx creates Letter 21.59 × 27.94 cm |
| `page_height_cm` | `29.7` | Page height (A4) |

## page_numbering

| Key | Default | Affects |
|---|---|---|
| `page_numbering` | `true` | Page number in the footer |
| `page_number_align` | `"center"` | `center`, `right`, `left` |
| `page_number_size` | `12.0` | Number size |
| `page_number_diff_first` | `true` | No number on the first page (the title page never has one) |

## captions

| Key | Default | Affects |
|---|---|---|
| `caption_autonumber` | `true` | Caption auto-numbering. A caption already starting with «Рисунок 3» is left unchanged |
| `caption_numbering` | `"global"` | `global`: «Рисунок 5»; `chapter`: «Рисунок 2.3» (the counter resets at every H1) |
| `caption_separator` | `" — "` | Between the number and the title |
| `figure_label` | `"Рисунок"` | Figure caption label |
| `table_label` | `"Таблиця"` | Table caption label |
| `listing_label` | `"Лістинг"` | Listing caption label |
| `image_caption_italic` | `false` | Italic figure caption |
| `table_caption_italic` | `false` | Italic table caption |
| `listing_caption_italic` | `false` | Italic listing caption |

## tables

| Key | Default | Affects |
|---|---|---|
| `table_caption_position` | `"above"` | `above` / `below` |
| `table_caption_align` | `"center"` | Table caption alignment |
| `table_cell_align` | `"center"` | Data alignment if `col_align` is not given (the header is always centered) |
| `table_font_size` | `14.0` | Table font size |
| `table_line_spacing` | `1.0` | Spacing in cells |
| `table_space_after_pt` | `6` | Space after the table (if the caption is not below) |
| `table_border_color` | `"000000"` | Border color |
| `table_border_sz` | `"8"` | Border width in eighths of a point |

## images

| Key | Default | Affects |
|---|---|---|
| `images_dir` | `"images"` | Images folder (relative or absolute) |
| `image_caption_position` | `"below"` | `above` / `below` |
| `image_caption_align` | `"center"` | Figure caption alignment |
| `image_default_size` | `"real"` | `real`: pixels at 96 dpi; `width`: full text width |
| `image_max_width_cm` | `null` | Width limit (`null` = text width) |
| `image_max_height_cm` | `22.0` | Height limit so that a figure with its caption fits on a page |
| `strict_images` | `false` | `true`: a missing file raises an exception instead of a red placeholder |

## code

| Key | Default | Affects |
|---|---|---|
| `code_style` | `null` | Template style for listing lines (e.g. `"Код исх"`); takes priority over the font and indent below |
| `code_font_name` | `"Consolas"` | Font of listings and inline `code` |
| `code_font_size` | `10.0` | Listing font size |
| `code_hanging_indent_cm` | `0.8` | Hanging indent of wrapped lines |
| `code_keep_together_lines` | `15` | A listing up to this many lines is not split across pages |
| `listing_caption_position` | `"above"` | `above` / `below` |
| `listing_caption_align` | `"left"` | Listing caption alignment |

## markdown

| Key | Default | Affects |
|---|---|---|
| `inline_math` | `true` | Parsing `$...$` as formulas. In reports with prices or shell variables, turn it off or escape `\$` |

## check

| Key | Default | Affects |
|---|---|---|
| `template_placeholders` | `["\\bTODO\\b", "\\b[XХ]{3,}\\b"]` | Regular expressions for leftover placeholder text for `doc.check()` |
| `toc_update_on_open` | `true` | Word updates the TOC when the DOCX is opened |

## title

Default values of `add_title_page()` fields. Call arguments override them; empty fields are skipped. The values are document content, so they stay in Ukrainian.

| Field | Example | Where on the title page |
|---|---|---|
| `ministry`, `university`, `faculty`, `department` | «Кафедра ІПІ» | Top block, centered |
| `report_word` | `"ЗВІТ"` | First line of the middle block |
| `work_type`, `number` | «Лабораторна робота», `"3"` | «ЛАБОРАТОРНА РОБОТА №3» |
| `discipline`, `topic`, `variant` | | «з дисципліни «…»», «на тему: «…»», «Варіант …» |
| `student_label`, `group_label`, `group`, `student` | «Виконав:», «студент групи» | Signature block on the right |
| `teacher_label`, `teacher` | «Перевірив:» | Signature block on the right |
| `signature_indent_cm` | `9.0` | Indent of the signature block from the left margin |
| `city`, `year` | «Київ», `"2026"` | Bottom line «Київ — 2026» |
