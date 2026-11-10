"""
chapters/example_lab.py
Reference structure of a lab report: title page -> goal -> task -> steps
(listing, figure, a sentence about the result) -> conclusion.
Copy the structure, replacing the text. The engine numbers captions itself.
"""

import os

HERE = os.path.dirname(os.path.abspath(__file__))


def render(doc):
    # Fields missing here are taken from the "title" section of styles_config.json
    doc.add_title_page(
        university="Національний технічний університет України «Київський політехнічний інститут імені Ігоря Сікорського»",
        faculty="Факультет інформатики та обчислювальної техніки",
        department="Кафедра інформатики та програмної інженерії",
        number="1",
        discipline="Бази даних",
        topic="Створення та наповнення реляційної бази даних",
        variant="10",
        group="ІП-00",
        student="Прізвище Ім'я По батькові",
        teacher="посада Прізвище І. П.",
    )

    doc.add_heading("Мета роботи", level=1, page_break_before=False)
    doc.add_paragraph("Навчитися створювати таблиці з обмеженнями цілісності та заповнювати їх тестовими даними.")

    doc.add_heading("Постановка задачі", level=1, page_break_before=False)
    doc.add_paragraph("Для предметної області свого варіанта потрібно:")
    doc.add_numbered("створити таблиці з первинними та зовнішніми ключами;")
    doc.add_numbered("заповнити кожну таблицю щонайменше десятьма записами;")
    doc.add_numbered("перевірити роботу обмежень некоректними вставками.")

    doc.add_heading("Хід роботи", level=1, page_break_before=False)
    doc.add_heading("1. Створення таблиць", level=2)
    # A number can be taken before the object itself is inserted
    doc.add_paragraph(
        f"Скрипт створення таблиць наведено в лістингу {doc.ref('lst', 'ddl')}, "
        f"отриману схему показано на рисунку {doc.ref('fig', 'schema')}."
    )
    doc.add_listing_from_file(
        os.path.join(HERE, "sample.sql"),
        start="CREATE TABLE", end="-- [END DDL]",
        caption="Створення таблиць", key="ddl", keep_with_next=True,
    )
    doc.add_image("sample.png", caption="Схема бази даних", key="schema")
    doc.add_paragraph("Обидві таблиці створено, зовнішній ключ `orders.client_id` посилається на `clients.id`.")

    doc.add_heading("2. Перевірка обмежень", level=2)
    doc.add_paragraph(f"Результати спроб некоректних вставок зведено в таблиці {doc.ref('tab', 'checks')}.")
    doc.add_table(
        ["Операція", "Очікування", "Результат"],
        [
            ["INSERT з порожнім `name`", "відмова", "помилка NOT NULL"],
            ["INSERT з неіснуючим `client_id`", "відмова", "помилка FOREIGN KEY"],
        ],
        caption="Перевірка обмежень цілісності", key="checks",
        col_widths_cm=[7.5, 3.5, 6.0], col_align=["left", "center", "left"],
    )
    doc.add_placeholder("скриншот повідомлення про помилку FOREIGN KEY")

    doc.add_heading("Висновок", level=1, page_break_before=False)
    doc.add_paragraph(
        "Я створив дві пов'язані таблиці й переконався, що СУБД відхиляє записи з порожнім іменем "
        "і з посиланням на неіснуючого клієнта. Каскадне видалення я не вмикав: для замовлень "
        "безпечніше отримати помилку, ніж мовчки втратити історію."
    )
