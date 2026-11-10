"""
chapters/chapter1.py
Generic template of modular chapter 1.
Each chapter file implements a single function: render(doc).
"""

def render(doc):
    # 1. Level-1 chapter heading (16pt, Bold)
    doc.add_heading("1. НАЗВА ПЕРШОГО РОЗДІЛУ", level=1)
    
    # 2. Body text (14pt, 1.25 cm first-line indent, justified, inline Markdown)
    doc.add_paragraph(
        "Тут розміщується вступний текст розділу. Функція add_paragraph підтримує "
        "зручне вбудоване форматування Markdown: можна виділяти слова **жирним шрифтом**, "
        "*курсивом*, комбінувати ***жирний курсив*** або вставляти інлайн-код на зразок `x_i` чи `heap_sort()`. "
        "Всі стандарти (шрифт, міжрядковий інтервал, абзацний відступ 1.25 см) застосовуються автоматично."
    )
    
    # 3. Level-2 subsection (14pt, Bold)
    doc.add_heading("1.1. Підрозділ теоретичних відомостей", level=2)
    doc.add_paragraph(
        "Опис основних теоретичних положень або вхідних вимог. Нижче наведено перелік ключових аспектів:"
    )
    
    # 4. Bulleted list (supports Markdown too)
    doc.add_bullet("Перший пункт: **параметр A** визначає швидкість збіжності;")
    doc.add_bullet("Другий пункт: змінна `tolerance` встановлює точність обчислень;")
    doc.add_bullet("Третій пункт: *підсумкова теза* теоретичного огляду.")
    
    # 5. Subsection with an image
    doc.add_heading("1.2. Графічні матеріали та схеми", level=2)
    doc.add_paragraph(
        "Для вставки фото чи схеми достатньо вказати назву файлу з папки images/. "
        "Рушій автоматично відцентрує зображення, масштабує його під ширину аркуша та додасть центрований підпис:"
    )
    
    # Automatic insertion of a picture from the images/ folder
    doc.add_image("sample.png", caption="Приклад графічної схеми або ілюстрації")
    
    # 6. Math formulas (LaTeX / MathML as native Word oMath)
    doc.add_paragraph(
        "Рушій підтримує як вбудовані формули в тексті на зразок $\\sqrt{x^2 + y^2}$ або $\\sum_{i=1}^n x_i$, "
        "так і виділені розрахункові формули з автоматичною нумерацією праворуч:"
    )
    doc.add_formula(r"f(x) = \frac{a \cdot x^2 + b}{\sqrt{x}} + \sum_{i=1}^n x_i", number="1.1")
