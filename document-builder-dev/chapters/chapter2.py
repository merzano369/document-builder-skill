"""
chapters/chapter2.py
Generic template of modular chapter 2.
Each chapter file implements a single function: render(doc).
"""

def render(doc):
    # 1. Level-1 chapter heading
    doc.add_heading("2. ПРАКТИЧНА ЧАСТИНА ТА РЕЗУЛЬТАТИ", level=1)
    
    doc.add_paragraph(
        "У цьому розділі наводяться практичні результати, зведені табличні дані та лістинги програмного коду. "
        "Усі елементи автоматично форматуються відповідно до стилів, визначених у styles_config.json."
    )
    
    # 2. Subsection with a table
    doc.add_heading("2.1. Зведена таблиця показників", level=2)
    doc.add_paragraph("Приклад структурованої таблиці з рамками, відступами та захистом від розриву рядків:")
    
    # Academic / technical table
    headers = ["Параметр / Метрика", "Базове значення", "Оптимізоване", "Приріст (%)"]
    rows = [
        ["Швидкодія (операцій/сек)", "1,250", "4,800", "+284%"],
        ["Використання пам'яті (MB)", "128", "45", "-64.8%"],
        ["Час затримки (мс)", "42.5", "11.2", "-73.6%"],
        ["Точність класифікації", "91.2%", "98.7%", "+7.5%"]
    ]
    doc.add_table(headers, rows, caption="Зведені результати порівняльного аналізу")
    
    # 3. Subsection with a code listing
    doc.add_heading("2.2. Лістинг програмного коду", level=2)
    doc.add_paragraph(
        "Блоки коду додаються зі збереженням структури рядків, моноширинним шрифтом "
        "та вимкненням перевірки орфографії Word (щоб документ на 100+ сторінок не зависав):"
    )
    
    sample_code = (
        "# Приклад універсальної функції обробки\n"
        "def process_pipeline(data: list) -> dict:\n"
        "    results = {}\n"
        "    for item in data:\n"
        "        metric = compute_metric(item)\n"
        "        results[item.id] = metric\n"
        "    return results"
    )
    doc.add_code_block(sample_code, caption="Приклад реалізації модуля обробки")
    
    # 4. Closing paragraph
    doc.add_paragraph(
        "Підсумкові висновки за результатами експериментів. "
        "Отримані показники свідчать про високу ефективність запропонованого підходу."
    )
