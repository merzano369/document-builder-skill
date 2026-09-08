"""
DocumentBuilder.py
Головний збирач (оркестратор) документа.
Ініціалізує рушій DocumentStyles, послідовно викликає розділи із chapters/ та зберігає готовий DOCX.
"""

from __future__ import annotations
import os
import sys

# Додаємо поточну директорію для коректного імпорту модулів
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

from DocumentStyles import DocumentStyles

# Імпорт розділів документа
from chapters.chapter1 import render as render_chapter1
from chapters.chapter2 import render as render_chapter2


def build_full_document(output_filename: str = "document.docx"):
    """
    Головна функція збирання повного документа.
    """
    print("🚀 Початок формування документа...")
    doc = DocumentStyles()

    # 1. Автоматичний зміст (TOC) на початку документа (за потреби можна вимкнути)
    print("  -> Додавання автоматичного змісту (TOC)...")
    doc.add_toc(title="ЗМІСТ")
    doc.add_page_break()

    # 2. Послідовний виклик розділів (для великих документів на 10-100+ сторінок)
    # Щоб додати новий розділ, створіть chapters/chapter3.py з функцією render(doc) і додайте його сюди
    print("  -> Збирання Розділу 1...")
    render_chapter1(doc)

    print("  -> Збирання Розділу 2...")
    render_chapter2(doc)

    # Збереження готового документа
    output_path = os.path.join(current_dir, output_filename)
    saved_path = doc.save(output_path)
    print(f"🎉 Документ сформовано: {saved_path}")
    return saved_path


if __name__ == "__main__":
    out_name = sys.argv[1] if len(sys.argv) > 1 else "document.docx"
    build_full_document(out_name)
