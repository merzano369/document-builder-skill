---
name: document-builder
description: Модульний генератор академічних та технічних DOCX документів будь-якого обсягу (від 1 до 100+ сторінок) з повним налаштуванням стилів, автоматичними фото, таблицями, вбудованим форматуванням Markdown, нативними формулами LaTeX/MathML та нумерацією.
---

# Document Builder Skill

Скіл для швидкого та простого створення структурованих академічних і технічних `.docx` документів (від коротких звітів на кілька сторінок до великих робіт на 100+ сторінок).

---

## 1. Архітектура скіла

```
Document-builder/
├── SKILL.md             # Головні інструкції скіла для агента та розробника
├── styles_config.json   # Відокремлена конфігурація стилів (шрифти, відступи, поля, таблиці)
├── DocumentStyles.py    # Рушій OpenXML (автоматично читає styles_config.json)
├── DocumentBuilder.py   # Головний збирач (оркестратор) документа
├── MML2OMML.XSL         # Стандартний XSL-транслятор MathML у нативні формули Word oMath
├── requirements.txt     # Залежності: python-docx, latex2mathml, lxml, matplotlib
├── chapters/            # Папка розділів для великих документів (100+ стор.)
│   ├── chapter1.py      # def render(doc): ...
│   └── chapter2.py      # def render(doc): ...
└── images/              # Папка для фото / схем / графіків
    └── (ваші зображення)
```

---

## 2. Головні принципи роботи для ШІ-агента

1. **Жодного низькорівневого коду вручну**: агент не повинен писати сирий `oxml`, рахувати `twips` чи вручну фарбувати комірки. Усі дії виконуються через лаконічні виклики `doc.add_*`.
2. **Відокремлена конфігурація стилів (`styles_config.json`)**:
   - Налаштування зручно розбиті по категоріях: `fonts`, `alignment`, `paragraph_format`, `page_margins`, `page_numbering`, `tables`, `images`.
   - У секції `"alignment"` можна незалежно задавати вирівнювання:
     - `paragraph_align`: вирівнювання звичайних абзаців (за замовчуванням `"justify"`).
     - `bullet_align`: вирівнювання маркованих списків (`"justify"`, `"left"`, `"center"`, `"right"`).
     - `numbered_align`: вирівнювання нумерованих списків (`"justify"`, `"left"`, `"center"`, `"right"`).
     - `heading_align`, `heading_h1_align`, `heading_h2_align`, `heading_h3_align`: вирівнювання заголовків.
   - Агенту або користувачу не потрібно змінювати Python-код для налаштування оформлення!
3. **Вбудоване форматування Markdown (Inline Formatting)**:
   Всередині `add_paragraph()`, `add_bullet()`, `add_numbered()` та комірок таблиць підтримується нативний базовий синтаксис Markdown:
   - `**жирний**` -> напівжирний текст
   - `*курсив*` -> курсив
   - `***жирний курсив***` -> жирний курсив
   - `` `код` `` -> моноширинний Consolas (із вбудованим захистом `w:noProof` проти хвилястих ліній перевірки орфографії)
   - `$inline_formula$` -> вбудована нативна формула Word (наприклад `$x_i$` чи `$\sqrt{a^2 + b^2}$`)
   
   *Приклад:* `doc.add_paragraph("Піраміда **Hlow** зберігає елементи, а $x_i$ надходить на вхід.")`
4. **Підтримка нативних формул LaTeX / MathML (OpenXML oMath)**:
   - `doc.add_formula(r"f(x) = \frac{a \cdot x^2 + b}{\sqrt{x}} + \sum_{i=1}^n x_i", number="1.1")`
   - Конвертує вираз у повноцінний нативний редагований об'єкт формули Word (Cambria Math, вертикальні дроби, квадратні корені, межі суми/інтегралу).
   - Якщо вказано `number="1.1"` — формула автоматично центрується, а номер у дужках `(1.1)` розміщується по правому краю.
   - Має автоматичний fallback: OMML -> Matplotlib PNG -> Текстовий вираз.
5. **Для невеликих документів (до 10 сторінок)**:
   - Можна писати всі виклики `doc.add_*` прямо всередині `DocumentBuilder.py`.
6. **Для великих документів (10+ або 100+ сторінок)**:
   - Розбивайте роботу на окремі файли у папці `chapters/chapterX.py`.
   - У кожному файлі розділу пишіть функцію `def render(doc):`.
   - `DocumentBuilder.py` послідовно викликає `render_ch1(doc)`, `render_ch2(doc)` і зберігає єдиний файл.
   - Це запобігає вичерпанню контекстного вікна LLM та дозволяє зневаджувати розділи окремо.
7. **Робота із зображеннями**:
   - Покладіть файл у папку `images/`.
   - Викличте `doc.add_image("name.png", caption="Підпис")`. Метод сам відцентрує зображення, масштабує його під ширину полів і додасть центрований підпис.
8. **Захист від блокування файлів у Windows (PermissionError: [Errno 13])**:
   - Якщо попередній файл `.docx` відкрито у Microsoft Word, Windows блокує його перезапис. Метод `doc.save(path)` перехоплює блокування і безпечно зберігає файл як `_new.docx` (або `_new_1.docx` тощо), повертаючи шлях без аварійного падіння програми.

---

## 3. Cheat Sheet компонентів (`DocumentStyles`)

| Метод | Опис | Приклад виклику |
|---|---|---|
| `DocumentStyles()` | Ініціалізація документа зі стилями з `styles_config.json` | `doc = DocumentStyles()` |
| `add_heading(text, level, page_break_before, align)` | Заголовок H1–H3 (`keep_with_next=True`, авторозрив сторінки для H1 крім першого) | `doc.add_heading("1. ВСТУП", level=1)` |
| `add_paragraph(text, ...)` | Абзац (ДСТУ: 14pt, 1.25 см відступ, по ширині + автопарсинг Markdown та `$формул$`) | `doc.add_paragraph("Піраміда **Hlow**, а $x_i$...")` |
| `add_bullet(text, bullet_char, parse_markdown, align)` | Нативний маркований пункт Word (`List Bullet`, вирівнювання за `bullet_align` або параметром `align`) | `doc.add_bullet("**Складність:** $O(n \\log n)$")` |
| `add_numbered(text, parse_markdown, align)` | Нативний нумерований пункт Word (`List Number`, вирівнювання за `numbered_align` або параметром `align`) | `doc.add_numbered("Крок 1: обчислення матриці")` |
| `add_image(image_name, caption, width_cm, caption_position, caption_align)` | Фото з папки images/ з центруванням, автопідгонкою та підписом зверху/знизу | `doc.add_image("plot.png", "Рисунок 1: Графік")` |
| `add_code_block(code, caption)` | Код (Consolas, збереження переносу рядків, `w:noProof`) | `doc.add_code_block(code, "Лістинг 1.1")` |
| `add_formula(formula, number)` | Нативна Word oMath формула (LaTeX / MathML, центрована, номер `(1.1)` праворуч) | `doc.add_formula(r"\frac{a}{b} + \sqrt{c}", number="1.1")` |
| `add_toc(title="ЗМІСТ")` | Динамічне поле Word TOC автоматичного змісту | `doc.add_toc()` |
| `add_page_break()` | Розрив сторінки | `doc.add_page_break()` |
| `save(path)` | Безпечне збереження DOCX (із захистом від блокування Word у Windows) | `doc.save("document.docx")` |

---

## 4. Швидкий приклад написання розділу (`chapters/chapter1.py`)

```python
"""
Розділ документу
"""

def render(doc):
    doc.add_heading("1. МЕТА ТА ПОСТАНОВКА ЗАДАЧІ", level=1)
    
    # Використання зручного Markdown та inline LaTeX у add_paragraph
    doc.add_paragraph(
        "У цьому розділі досліджується піраміда **Hlow** та вхідні змінні $x_i$. "
        "Для докладного аналізу використано *теоретичні оцінки*."
    )
    
    doc.add_bullet("Алгоритм **QuickSort**: часова складність $O(n \\log n)$;")
    doc.add_bullet("Алгоритм **MergeSort**: стабільне сортування.")
    
    # Вставка фото
    doc.add_image("diagram.png", caption="Рисунок 1.1: Схема експерименту")
    
    # Вставка нативної LaTeX формули з номером (1.1)
    doc.add_paragraph("Основне розрахункове рівняння експерименту:")
    doc.add_formula(r"f(x) = \frac{a \cdot x^2 + b}{\sqrt{x}} + \sum_{i=1}^n x_i", number="1.1")
    
    # Вставка таблиці
    headers = ["Алгоритм", "Кращий випадок", "Гірший випадок"]
    rows = [
        ["`QuickSort`", "**O(n log n)**", "O(n^2)"],
        ["`MergeSort`", "**O(n log n)**", "**O(n log n)**"],
    ]
    doc.add_table(headers, rows, caption="Таблиця 1.1: Складність алгоритмів")
    
    doc.add_page_break()
```

---

## 5. Швидкий приклад головного збирача (`DocumentBuilder.py`)

```python
from DocumentStyles import DocumentStyles
from chapters.chapter1 import render as render_ch1
from chapters.chapter2 import render as render_ch2

def main():
    doc = DocumentStyles()
    
    # Автозміст
    doc.add_toc()
    doc.add_page_break()
    
    # Послідовний виклик розділів
    render_ch1(doc)
    render_ch2(doc)
    
    doc.save("Report_100_Pages.docx")

if __name__ == "__main__":
    main()
```
