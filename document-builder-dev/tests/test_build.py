"""
Engine smoke test. Run after every change to DocumentStyles.py:
    python tests/test_build.py          (or pytest tests/)
Builds a document with all methods in a temporary folder, opens it and checks captions,
numbering, lists, listings and comment removal.
"""

import os
import re
import sys
import tempfile

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT)

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

from DocumentStyles import DocumentStyles, strip_code_comments  # noqa: E402
from chapters.example_lab import render as render_lab  # noqa: E402
from chapters.chapter1 import render as render_ch1  # noqa: E402
from chapters.chapter2 import render as render_ch2  # noqa: E402

CONFIG = os.path.join(ROOT, "styles_config.json")


def _texts(path):
    return [p.text for p in Document(path).paragraphs]


def test_build():
    out = os.path.join(tempfile.mkdtemp(prefix="docbuilder_"), "test.docx")
    doc = DocumentStyles(config_path=CONFIG)
    render_lab(doc)
    doc.add_heading("Додаток", level=1)
    render_ch1(doc)
    render_ch2(doc)
    doc.add_toc()
    doc.add_line("Київ — 2026", bold=True)
    doc.add_paragraph(r"Екранування: \*не курсив\* і \$5\$, але *курсив* і $x_i$.")
    doc.add_code_block("x = 1\ny = 2", caption="Лістинг 99 — ручний номер")  # a manual number is left as is
    doc.add_image("no_such_file.png", caption="Відсутній рисунок")
    saved = doc.save(out)

    texts = _texts(saved)
    caps = [t for t in texts if re.match(r"(Рисунок|Таблиця|Лістинг) \d+ — ", t)]
    assert [t for t in caps if t.startswith("Рисунок")] == [
        "Рисунок 1 — Схема бази даних",
        "Рисунок 2 — Приклад графічної схеми або ілюстрації",
        "Рисунок 3 — Відсутній рисунок",
    ], caps
    assert sum(t.startswith("Таблиця") for t in caps) == 2, caps
    assert sum(t.startswith("Лістинг") for t in caps) == 3, caps
    assert "Лістинг 99 — ручний номер" in texts

    # reference before the object is inserted
    assert any("лістингу 1, отриману схему показано на рисунку 1" in t for t in texts)
    assert any("таблиці 1." in t for t in texts)
    # listing caption above the code
    i = texts.index("Лістинг 1 — Створення таблиць")
    assert texts[i + 1].startswith("CREATE TABLE clients")
    # comments removed, the line with '--' inside quotes kept
    listing = "\n".join(texts[i + 1:i + 14])
    assert "ідентифікатор" not in listing and "Багаторядковий" not in listing
    assert "it''s -- not a comment" in listing
    # escapes
    assert any("*не курсив* і $5$" in t for t in texts)
    # placeholders and missing files
    assert doc.missing == ["no_such_file.png"]
    assert len(doc.todo) == 1
    assert not doc.warnings, doc.warnings

    # numbered list: a separate w:num with startOverride
    d = Document(saved)
    num_ids = {p._p.pPr.numPr.numId.val for p in d.paragraphs
               if p.style.name == "List Number" and p._p.pPr is not None and p._p.pPr.numPr is not None}
    assert len(num_ids) == 1, num_ids
    # title page in its own section with vertical alignment
    assert len(d.sections) == 2
    assert d.sections[0]._sectPr.find(qn("w:vAlign")) is not None
    # A4 format, not Letter from the default python-docx template
    assert all(abs(s.page_width.cm - 21.0) < 0.01 and abs(s.page_height.cm - 29.7) < 0.01 for s in d.sections)
    print("test_build: OK", saved)


def test_numbered_restart():
    doc = DocumentStyles(config_path=CONFIG, verbose=False)
    doc.add_numbered("a")
    doc.add_numbered("b")
    doc.add_paragraph("розділювач")
    doc.add_numbered("c")             # automatic restart
    doc.add_numbered("d", restart=True)
    ids = [p._p.pPr.numPr.numId.val for p in doc.doc.paragraphs if p.style.name == "List Number"]
    assert ids[0] == ids[1] and ids[1] != ids[2] and ids[2] != ids[3], ids

    doc.config["numbered_mode"] = "manual"
    doc.add_paragraph("x")
    doc.add_numbered("перший")
    doc.add_numbered("другий")
    assert doc.doc.paragraphs[-1].text == "2)	другий"
    print("test_numbered_restart: OK")


def test_strip_comments():
    py = 'x = "#не коментар"  # коментар\n# лише коментар\ny = 2\n'
    assert strip_code_comments(py, ".py") == 'x = "#не коментар"\ny = 2'
    cs = 'var s = "// ні"; // так\n/* блок */int a = 1;'
    assert strip_code_comments(cs, ".cs") == 'var s = "// ні";\nint a = 1;'
    print("test_strip_comments: OK")


def test_chapter_numbering_and_no_math():
    doc = DocumentStyles(config_path=CONFIG, verbose=False)
    doc.config["caption_numbering"] = "chapter"
    doc.config["inline_math"] = False
    doc.add_heading("1 Розділ")
    doc.add_code_block("a", caption="перший")
    doc.add_heading("2 Розділ")
    doc.add_code_block("b", caption="другий")
    doc.add_paragraph("Ціна $5 і $10$ без формул")
    texts = [p.text for p in doc.doc.paragraphs]
    assert "Лістинг 1.1 — перший" in texts and "Лістинг 2.1 — другий" in texts, texts
    assert "Ціна $5 і $10$ без формул" in texts
    print("test_chapter_numbering_and_no_math: OK")


def _make_template(path):
    """Template with 'hostile' conditions: no List Bullet/List Number, a custom code style, placeholders,
    highlighting, footer text split into several runs."""
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_COLOR_INDEX
    from docx.shared import Cm, Pt
    d = Document()
    for name in ("List Bullet", "List Number"):
        el = d.styles[name].element
        el.getparent().remove(el)
    st = d.styles.add_style("Код исх", WD_STYLE_TYPE.PARAGRAPH)
    st.font.name, st.font.size = "Courier New", Pt(9)
    d.sections[0].left_margin = Cm(2.0)
    f = d.sections[0].footer.paragraphs[0]
    for chunk in ("Студент: {П", "І", "Б}"):
        f.add_run(chunk)
    d.add_paragraph("Звіт виконав {ПІБ}")
    d.add_heading("Хід роботи", 1)
    d.add_paragraph("XXX опишіть хід роботи")
    d.add_paragraph("ще текст шаблону")
    d.add_heading("Результати", 1)
    t = d.add_table(rows=2, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = "Параметр", "Значення"
    t.rows[1].cells[0].paragraphs[0].add_run("впишіть").font.highlight_color = WD_COLOR_INDEX.YELLOW
    d.add_heading("Висновки", 1)
    d.add_paragraph("XXX")
    d.save(path)


def test_template():
    tmp = tempfile.mkdtemp(prefix="docbuilder_tpl_")
    tpl = os.path.join(tmp, "template.docx")
    _make_template(tpl)
    doc = DocumentStyles(config_path=CONFIG, doc_path=tpl, verbose=False)
    assert doc.keep_template_styles
    assert abs(doc.text_width_cm - (21.0 - 2.0 - 2.54)) < 0.2, doc.text_width_cm  # template margins
    before = doc.check(verbose=False)
    assert any("XXX" in p for p in before) and any("highlighted" in p for p in before), before

    assert doc.replace_text("{ПІБ}", "Іваненко І. І.") == 2  # paragraph + footer with split text
    doc.config["code_style"] = "Код исх"
    with doc.insert_at(doc.clear_section("Хід роботи")):
        doc.add_paragraph("Я виконав такі кроки:")
        doc.add_bullet("перший")
        doc.add_numbered("крок один")
        doc.add_numbered("крок два")
        doc.add_code_block("SELECT 1;", caption="Запит")
    doc.fill_table(0, [["Час", "12 мс"], ["Пам'ять", "3 МБ"]])
    with doc.insert_at(doc.clear_section("Висновки")):
        doc.add_paragraph("Висновок.")
    out = doc.save(os.path.join(tmp, "filled.docx"))

    d = Document(out)
    texts = [p.text for p in d.paragraphs]
    i = texts.index("Хід роботи")
    assert texts[i + 1:i + 7] == ["Я виконав такі кроки:", "–\tперший", "1)\tкрок один", "2)\tкрок два",
                                  "Лістинг 1 — Запит", "SELECT 1;"], texts[i:i + 8]
    assert texts[i + 7] == "Результати"
    assert d.paragraphs[i + 6].style.name == "Код исх"
    assert texts[-1] == "Висновок." and "XXX" not in " ".join(texts)
    assert d.sections[0].footer.paragraphs[0].text == "Студент: Іваненко І. І."
    assert "Іваненко" in texts[0]
    assert [c.text for c in d.tables[0].rows[2].cells] == ["Пам'ять", "3 МБ"]
    reopened = DocumentStyles(config_path=CONFIG, doc_path=out, verbose=False)
    assert reopened.check(verbose=False) == [], reopened.check(verbose=False)
    print("test_template: OK", out)


def test_config_sync():
    """Keys do not diverge between DEFAULT_CONFIG, styles_config.json and reference/config.md."""
    import json
    import DocumentStyles as DS
    with open(CONFIG, encoding="utf-8") as f:
        cfg = json.load(f)
    json_keys = {k for section, vals in cfg.items() if section != "title" for k in vals}
    code_keys = set(DS.DEFAULT_CONFIG)
    assert json_keys == code_keys, (json_keys ^ code_keys)
    with open(os.path.join(ROOT, "reference", "config.md"), encoding="utf-8") as f:
        doc_text = f.read()
    undocumented = [k for k in code_keys | set(cfg["title"]) if f"`{k}`" not in doc_text]
    assert not undocumented, undocumented
    print("test_config_sync: OK")


def test_config_live():
    """Changing doc.config after creating the document affects subsequent elements (no settings copies)."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc = DocumentStyles(config_path=CONFIG, verbose=False)
    doc.config["paragraph_align"] = "left"
    doc.config["font_size_body"] = 12
    p = doc.add_paragraph("текст")
    assert p.alignment == WD_ALIGN_PARAGRAPH.LEFT
    assert p.runs[0].font.size.pt == 12 and doc.font_size_body == 12.0
    print("test_config_live: OK")


def test_prev_element():
    """Decisions about breaks and numbering come from the document, not from what was called last."""
    from docx.oxml.ns import qn
    doc = DocumentStyles(config_path=CONFIG, verbose=False)
    doc.add_paragraph("вступ")
    doc.add_page_break()
    doc.add_page_break()                       # a second consecutive break is not added
    body = doc.doc.element.body
    breaks = [el for el in body if el.find(".//" + qn("w:br")) is not None]
    assert len(breaks) == 1, len(breaks)
    doc.add_heading("A")                       # first H1: no break
    doc.add_heading("B")                       # second H1: from a new page
    assert doc.doc.paragraphs[-1].paragraph_format.page_break_before

    # insert_at: the first item inside the block looks at the anchor's neighbour, not at the end of the document
    doc.add_numbered("один")
    doc.add_heading("Кінець", level=2)
    doc.add_paragraph("хвіст документа")
    with doc.insert_at("Кінець"):
        doc.add_numbered("два")                # the anchor is preceded by «один» -> the list continues
    with doc.insert_at("хвіст документа"):
        doc.add_numbered("новий")              # the anchor is preceded by a heading -> a new list
    ps = [p for p in doc.doc.paragraphs if p.style.name == "List Number"]
    ids = [p._p.pPr.numPr.numId.val for p in ps]
    assert [p.text for p in ps] == ["один", "два", "новий"], [p.text for p in ps]
    assert ids[0] == ids[1] != ids[2], ids

    # after the title page H1 gets no extra break
    doc2 = DocumentStyles(config_path=CONFIG, verbose=False)
    doc2.add_title_page(student="X")
    doc2.add_heading("Перший")
    doc2.add_heading("Другий")
    h = [p for p in doc2.doc.paragraphs if p.text in ("Перший", "Другий")]
    assert not h[0].paragraph_format.page_break_before and h[1].paragraph_format.page_break_before
    print("test_prev_element: OK")


def test_load_by_path():
    """DocumentStyles.py can be loaded by file path without the skill folder in sys.path."""
    import subprocess
    code = (
        "import importlib.util, sys\n"
        f"sys.path = [p for p in sys.path if p and __import__('os').path.abspath(p) != {ROOT!r}]\n"
        f"spec = importlib.util.spec_from_file_location('DS', {os.path.join(ROOT, 'DocumentStyles.py')!r})\n"
        "m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m)\n"
        "print(m.__version__, callable(m.convert_to_pdf), callable(m.strip_code_comments))\n"
    )
    r = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True, cwd=tempfile.gettempdir())
    assert r.returncode == 0 and "True True" in r.stdout, r.stderr
    print("test_load_by_path: OK")


if __name__ == "__main__":
    test_load_by_path()
    test_strip_comments()
    test_numbered_restart()
    test_chapter_numbering_and_no_math()
    test_config_sync()
    test_config_live()
    test_prev_element()
    test_build()
    test_template()
