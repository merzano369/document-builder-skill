# Document Builder

Скіл для ШІ-агентів (Claude Code, Antigravity та інших), який створює документи Word і PDF. Агент пише лише зміст простими викликами `doc.add_*`, а шрифти, відступи, нумерацію й підписи бере з одного файлу `styles_config.json`.

Підходить для документів будь-якого обсягу: від однієї сторінки до сотні з лишком.

## Можливості

- Титульна сторінка, заголовки, абзаци, списки, зміст.
- Автоматична нумерація рисунків, таблиць і лістингів, посилання на них за ключем (`doc.ref("fig", "schema")`).
- Лістинги прямо з файлів коду: потрібний фрагмент між маркерами, без коментарів.
- Зображення в реальному розмірі, які не виходять за поля сторінки.
- Формули LaTeX, що стають нативними формулами Word.
- Inline Markdown у тексті: `**жирний**`, `*курсив*`, `` `код` ``, `$x^2$`.
- Заповнення готового шаблону `.docx`: вставка в потрібний розділ, заміна тексту, заповнення таблиць.
- Конвертація в PDF і перевірка верстки: висячі підписи, пропуски в нумерації, забуті заглушки.

## Встановлення

1. Встановіть Python 3.9 або новіший.

2. Скопіюйте папку `document-builder/` у каталог скілів вашого агента:

   | Агент | Куди копіювати |
   |---|---|
   | Claude Code | `~/.claude/skills/document-builder/` |
   | Antigravity | `~/.gemini/config/skills/document-builder/` |

   Якщо скіл лежить деінде, задайте шлях до нього у змінній середовища `DOCUMENT_BUILDER_PATH`.

3. Для PDF потрібен Microsoft Word або LibreOffice.

Python-пакети встановлювати вручну не обов'язково. Перед першою збіркою в сесії агент сам перевіряє `python-docx`, `lxml`, `latex2mathml`, `matplotlib` і `Pillow` і, якщо чогось бракує, встановлює їх із `requirements.txt` (агент може попросити дозволу на виконання команди). Ця перевірка обов'язкова: без `lxml`, `latex2mathml`, `matplotlib` чи `Pillow` рушій не падає, а мовчки псує формули й розміри зображень.

Якщо хочете встановити все заздалегідь:

```
python -m pip install -r document-builder/requirements.txt
python -m pip install pymupdf   # для перевірки верстки PDF (tools/check_report.py)
```

Якщо `pip` відмовляє через права або «externally-managed-environment», додайте `--user`. Word і LibreOffice через pip не встановлюються: без них агент здасть лише DOCX.

## Як користуватися

Попросіть агента, наприклад: «Зроби документ з описом API, код лежить у `src/`, схеми в `images/`». Агент сам прочитає `SKILL.md` і збере документ.

Приблизно так виглядає код, який він напише:

```python
from DocumentStyles import DocumentStyles

doc = DocumentStyles()
doc.add_heading("Опис модуля")
doc.add_paragraph(f"Основну функцію наведено в лістингу {doc.ref('lst', 'main')}.")
doc.add_listing_from_file("src/main.py", caption="Основна функція", key="main")
doc.add_image("schema.png", caption="Схема модуля")
doc.save("Документ.docx", pdf=True)
```

Щоб змінити оформлення під свої вимоги, скопіюйте `styles_config.json` у папку проєкту і відредагуйте. Усі ключі описані в `document-builder/reference/config.md`, заповнення шаблону викладача — у `document-builder/reference/template.md`.

Інструменти можна запускати й вручну:

```
python document-builder/tools/export_pdf.py Звіт.docx          # DOCX -> PDF після ручних правок у Word
python document-builder/tools/check_report.py Звіт.pdf         # перевірка верстки готового PDF
```

## Структура репозиторію

Поточна версія скіла: 2.2.1.

| Папка | Для кого |
|---|---|
| `document-builder/` | Готовий скіл, саме його копіюють агенту |
| `document-builder-dev/` | Повна версія для тих, хто хоче змінювати скіл: те саме плюс документація архітектури, тести й інструкції для агента-розробника |

Вміст `document-builder/`:

| Шлях | Призначення |
|---|---|
| `SKILL.md` | Інструкція для агента: підключення, API `doc.add_*`, залежності, правила оформлення |
| `DocumentStyles.py` | Рушій: будує `.docx` через python-docx |
| `styles_config.json` | Усе оформлення: шрифти, відступи, підписи, нумерація, поля титулки |
| `DocumentBuilder.py` | Збирач: викликає розділи з `chapters/` і зберігає DOCX/PDF (`python DocumentBuilder.py [назва.docx] [--pdf]`) |
| `chapters/` | Розділи документа, кожен із функцією `render(doc)`; `example_lab.py` — еталонний звіт лабораторної, `sample.sql` — його лістинг із коментарями, `chapter1.py` і `chapter2.py` — шаблони |
| `code_comments.py` | Видалення коментарів із коду для лістингів |
| `report_pdf.py` | Конвертація DOCX → PDF: Word, далі LibreOffice, далі docx2pdf |
| `MML2OMML.XSL` | Перетворення MathML у формули Word |
| `reference/config.md` | Довідник ключів `styles_config.json` |
| `reference/template.md` | Заповнення готового шаблону `.docx` |
| `tools/export_pdf.py` | Окрема конвертація DOCX → PDF |
| `tools/check_report.py` | Перевірка PDF: висячі підписи, нумерація, заглушки, напівпорожні сторінки |
| `images/sample.png` | Зображення для прикладів |
| `requirements.txt` | Python-залежності |

Додатково в `document-builder-dev/`:

| Шлях | Призначення |
|---|---|
| `CLAUDE.md` | Правила для агента, що змінює скіл |
| `docs/ARCHITECTURE.md` | Карта файлів, механізми, інваріанти, підводні камені python-docx і Word |
| `tests/test_build.py` | Тести збірки, зокрема синхронізації `DEFAULT_CONFIG`, `styles_config.json` і `reference/config.md` |
