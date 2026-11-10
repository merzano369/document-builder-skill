"""
DocumentStyles.py
Word OpenXML engine for report generation.
All formatting comes from styles_config.json; chapter code contains only doc.add_* calls.
"""

from __future__ import annotations

import contextlib
import copy
import json
import os
import re
import sys
import tempfile
import textwrap
import uuid
from typing import Any, Dict, List, Optional, Tuple

__version__ = "2.2.1"

try:
    import lxml.etree as etree
    HAS_LXML = True
except ImportError:
    etree = None
    HAS_LXML = False

try:
    import latex2mathml.converter
    HAS_LATEX2MATHML = True
except ImportError:
    latex2mathml = None
    HAS_LATEX2MATHML = False

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    HAS_MATPLOTLIB = True
except ImportError:
    plt = None
    HAS_MATPLOTLIB = False

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    Image = None
    HAS_PIL = False

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor
from docx.text.paragraph import Paragraph

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))

# Separate skill modules. The skill folder is appended to the END of sys.path: imports work even when
# DocumentStyles.py was loaded by file path, and the user's own modules are not shadowed.
# Names are re-exported so old code (DS.convert_to_pdf, DS.strip_code_comments) keeps working.
if SKILL_DIR not in sys.path:
    sys.path.append(SKILL_DIR)
from code_comments import strip_code_comments  # noqa: E402,F401
from report_pdf import convert_to_pdf  # noqa: E402,F401

# Default values. Any key can be overridden in styles_config.json
# (in any section: sections exist only for readability).
DEFAULT_CONFIG: Dict[str, Any] = {
    # fonts
    "font_name": "Times New Roman",
    "font_size_body": 14.0,
    "font_size_h1": 16.0,
    "font_size_h2": 14.0,
    "font_size_h3": 14.0,
    # alignment
    "paragraph_align": "justify",
    "bullet_align": "justify",
    "numbered_align": "justify",
    "heading_align": "left",
    "heading_h1_align": "left",
    "heading_h2_align": "left",
    "heading_h3_align": "left",
    # paragraph_format
    "line_spacing": 1.15,
    "paragraph_indent_cm": 1.25,
    "h1_page_break_before": True,
    "numbered_mode": "native",          # "native" (Word list) | "manual" ("1)" prefix)
    # page_margins
    "margin_top_cm": 2.0,
    "margin_bottom_cm": 2.0,
    "margin_left_cm": 3.0,
    "margin_right_cm": 1.0,
    "page_width_cm": 21.0,              # A4; the default python-docx template is Letter
    "page_height_cm": 29.7,
    # page_numbering
    "page_numbering": True,
    "page_number_align": "center",
    "page_number_size": 12.0,
    "page_number_diff_first": True,
    # captions
    "caption_autonumber": True,
    "caption_numbering": "global",      # "global" (Рисунок 5) | "chapter" (Рисунок 2.3)
    "caption_separator": " — ",
    "figure_label": "Рисунок",
    "table_label": "Таблиця",
    "listing_label": "Лістинг",
    "image_caption_italic": False,
    "table_caption_italic": False,
    "listing_caption_italic": False,
    # tables
    "table_caption_position": "above",
    "table_caption_align": "center",
    "table_cell_align": "center",
    "table_font_size": 14.0,
    "table_line_spacing": 1.0,
    "table_space_after_pt": 6,
    "table_border_color": "000000",
    "table_border_sz": "8",
    # images
    "images_dir": "images",
    "image_caption_position": "below",
    "image_caption_align": "center",
    "image_default_size": "real",       # "real" (pixels at 96 dpi) | "width" (full text width)
    "image_max_width_cm": None,         # None = text width
    "image_max_height_cm": 22.0,
    "strict_images": False,
    # code
    "code_style": None,                 # template style name for listings (e.g. "Код исх")
    "code_font_name": "Consolas",
    "code_font_size": 10.0,
    "code_hanging_indent_cm": 0.8,
    "code_keep_together_lines": 15,
    "listing_caption_position": "above",
    "listing_caption_align": "left",
    # markdown
    "inline_math": True,
    # check: regular expressions for leftover template placeholder text (doc.check())
    "template_placeholders": [r"\bTODO\b", r"\b[XХ]{3,}\b"],
    "toc_update_on_open": True,
}

ALIGN_MAP = {
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}

KIND_ALIASES = {
    "fig": "fig", "figure": "fig", "image": "fig", "img": "fig",
    "tab": "tab", "table": "tab",
    "lst": "lst", "listing": "lst", "code": "lst",
}

_MD_PARTS = {
    "code": r"(?P<code>`[^`\n]+`)",
    "math": r"(?P<math>\$(?!\s)[^$\n]+?(?<!\s)\$)",
    "rest": (
        r"(?P<bold_italic>\*\*\*(?!\s)[^*\n]+?(?<!\s)\*\*\*)"
        r"|(?P<bold>\*\*(?!\s)[^*\n]+?(?<!\s)\*\*)"
        r"|(?P<italic>\*(?!\s)[^*\n]+?(?<!\s)\*)"
    ),
}
MD_RE_MATH = re.compile("|".join([_MD_PARTS["code"], _MD_PARTS["math"], _MD_PARTS["rest"]]))
MD_RE_NOMATH = re.compile("|".join([_MD_PARTS["code"], _MD_PARTS["rest"]]))
# Escapes \* \$ \` \\ : replaced with Private Use Area characters and restored after parsing
_ESCAPES = {"\\\\": "", "\\*": "", "\\$": "", "\\`": ""}
_UNESCAPE = {"": "\\", "": "*", "": "$", "": "`"}


def clean_hex(hex_str: str) -> str:
    return str(hex_str).lstrip("#").upper()


def hex_to_rgb(hex_str: str) -> RGBColor:
    c = clean_hex(hex_str)
    if len(c) != 6:
        c = "000000"
    return RGBColor(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))


def _is_page_break(el) -> bool:
    """Empty paragraph containing only a page break."""
    return (el is not None and el.tag == qn("w:p") and not "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
            and el.find(".//" + qn("w:br")) is not None and el.find(".//" + qn("w:drawing")) is None)


def _is_break(el) -> bool:
    """A new page already starts after el: a page break or a section break (end of the title page)."""
    return _is_page_break(el) or (el is not None and el.tag == qn("w:p")
                                  and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None)


def _unescape(text: str) -> str:
    for k, v in _UNESCAPE.items():
        text = text.replace(k, v)
    return text


# ------------------------------------------------------------------------------
# Config lookup
# ------------------------------------------------------------------------------

def find_config(config_path: Optional[str] = None) -> Optional[str]:
    """Explicit path -> ./#TOOLS/ -> ./ -> main script folder -> skill default config."""
    if config_path:
        return config_path
    main_dir = os.path.dirname(os.path.abspath(sys.argv[0])) if sys.argv and sys.argv[0] else os.getcwd()
    candidates = [
        os.path.join(os.getcwd(), "#TOOLS", "styles_config.json"),
        os.path.join(os.getcwd(), "styles_config.json"),
        os.path.join(main_dir, "styles_config.json"),
        os.path.join(SKILL_DIR, "styles_config.json"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


class DocumentStyles:
    """Document formatting engine. Public API: add_* methods, ref(), save()."""

    # Shortcut attributes for config keys: doc.font_size_body == float(doc.config["font_size_body"]).
    # The value is read from self.config every time, so changing doc.config after creating the document
    # affects all subsequent elements. Attributes hold no copies of settings.
    _CONFIG_ATTRS = {
        **{k: float for k in ("font_size_body", "font_size_h1", "font_size_h2", "font_size_h3", "line_spacing",
                              "paragraph_indent_cm", "margin_left_cm", "margin_right_cm", "margin_top_cm",
                              "margin_bottom_cm")},
        **{k: str for k in ("font_name", "paragraph_align", "bullet_align", "numbered_align", "heading_align",
                            "heading_h1_align", "heading_h2_align", "heading_h3_align", "table_border_color",
                            "table_border_sz", "images_dir")},
        **{k: (lambda v: str(v).lower()) for k in ("table_caption_position", "table_caption_align",
                                                   "image_caption_position", "image_caption_align")},
        "h1_page_break_before": bool,
    }

    def __getattr__(self, name):
        cast = DocumentStyles._CONFIG_ATTRS.get(name)
        if cast is None or "config" not in self.__dict__:
            raise AttributeError(name)
        return cast(self.config[name])

    def __init__(self, config_path: Optional[str] = None, doc_path: Optional[str] = None, verbose: bool = True,
                 keep_template_styles: Optional[bool] = None):
        """
        doc_path: .docx template that content is added to.
        keep_template_styles: True — styles, margins, headers/footers and page numbering come from the template,
        the config affects only new elements. Defaults to True when doc_path is given.
        """
        self.verbose = verbose
        if keep_template_styles is None:
            keep_template_styles = bool(doc_path)
        self.keep_template_styles = keep_template_styles
        self.config_path = find_config(config_path)
        self.config = self._load_config(self.config_path)
        c = self.config
        self._log(f"[document-builder {__version__}] config: {self.config_path or 'built-in defaults'}")

        # Document state
        self._h1_count = 0
        self._counters = {"fig": 0, "tab": 0, "lst": 0}
        self._refs: Dict[str, Dict[str, str]] = {"fig": {}, "tab": {}, "lst": {}}
        self._last_num = {"fig": 0, "tab": 0, "lst": 0}
        self._list_num_id: Optional[int] = None
        self._manual_num = 0
        self.missing: List[str] = []
        self.todo: List[str] = []
        self.warnings: List[str] = []
        self.pdf_path: Optional[str] = None

        self.doc = Document(doc_path) if doc_path else Document()
        if doc_path:
            self._log(f"[document-builder] template: {doc_path}"
                      + (" (template styles kept)" if keep_template_styles else " (styles from config)"))
        if not keep_template_styles:
            self._setup_page_margins()
            self._setup_default_styles()
            if c["page_numbering"]:
                self._setup_page_numbering()
        self._insert_ctx = None  # (new elements before the block, anchor, after) inside insert_at()

    # --------------------------------------------------------------------------
    # Config and service methods
    # --------------------------------------------------------------------------

    def _log(self, msg: str):
        if self.verbose:
            print(msg)

    def _load_config(self, path: Optional[str]) -> Dict[str, Any]:
        cfg = dict(DEFAULT_CONFIG)
        cfg["title"] = {}
        if not path:
            return cfg
        if not os.path.isfile(path):
            raise FileNotFoundError(f"styles_config.json not found: {path}")
        try:
            with open(path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
        except Exception as e:
            print(f"[WARNING] Could not read {path} ({e}), using defaults.")
            return cfg
        unknown = []
        for k, v in loaded.items():
            if k == "title" and isinstance(v, dict):
                cfg["title"] = dict(v)
            elif isinstance(v, dict):
                for kk, vv in v.items():
                    if kk not in DEFAULT_CONFIG:
                        unknown.append(kk)
                    cfg[kk] = vv
            else:
                if k not in DEFAULT_CONFIG:
                    unknown.append(k)
                cfg[k] = v
        if unknown:
            print(f"[WARNING] Unknown keys in {os.path.basename(path)}: {', '.join(unknown)}")
        return cfg

    @property
    def text_width_cm(self) -> float:
        s = self.doc.sections[-1]
        return s.page_width.cm - s.left_margin.cm - s.right_margin.cm

    @property
    def text_height_cm(self) -> float:
        s = self.doc.sections[-1]
        return s.page_height.cm - s.top_margin.cm - s.bottom_margin.cm

    def _setup_page_margins(self):
        for section in self.doc.sections:
            section.page_width = Cm(float(self.config["page_width_cm"]))
            section.page_height = Cm(float(self.config["page_height_cm"]))
            section.top_margin = Cm(self.margin_top_cm)
            section.bottom_margin = Cm(self.margin_bottom_cm)
            section.left_margin = Cm(self.margin_left_cm)
            section.right_margin = Cm(self.margin_right_cm)

    def _setup_default_styles(self):
        styles = self.doc.styles
        normal = styles["Normal"]
        normal.font.name = self.font_name
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), self.font_name)
        normal.font.size = Pt(self.font_size_body)
        normal.font.color.rgb = hex_to_rgb("000000")
        pf = normal.paragraph_format
        pf.line_spacing = self.line_spacing
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.alignment = ALIGN_MAP.get(self.paragraph_align.lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)

        for name, align in (("List Bullet", self.bullet_align), ("List Number", self.numbered_align)):
            if name in styles:
                st = styles[name]
                st.font.name = self.font_name
                st.font.size = Pt(self.font_size_body)
                st.font.color.rgb = hex_to_rgb("000000")
                st.paragraph_format.alignment = ALIGN_MAP.get(align.lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)

        sizes = {1: self.font_size_h1, 2: self.font_size_h2, 3: self.font_size_h3}
        for lvl in range(1, 4):
            name = f"Heading {lvl}"
            if name in styles:
                st = styles[name]
                st.font.name = self.font_name
                st.font.size = Pt(sizes[lvl])
                st.font.bold = True
                st.font.italic = False
                st.font.color.rgb = hex_to_rgb("000000")
                st.paragraph_format.line_spacing = self.line_spacing
                st.paragraph_format.keep_with_next = True

    def _setup_page_numbering(self):
        section = self.doc.sections[0]
        if self.config["page_number_diff_first"]:
            section.different_first_page_header_footer = True
        p = section.footer.paragraphs[0]
        p.text = ""
        align = self.config["page_number_align"]
        p.alignment = ALIGN_MAP.get(str(align).lower(), WD_ALIGN_PARAGRAPH.CENTER)
        r = p.add_run()
        self._apply_run_font(r, size_pt=float(self.config["page_number_size"]))
        self._insert_field_code(r, "PAGE")

    def _apply_run_font(self, run, font_name=None, size_pt=None, bold=False, italic=False, color_hex="000000"):
        font = run.font
        font.name = font_name or self.font_name
        font.size = Pt(size_pt or self.font_size_body)
        font.bold = bold
        font.italic = italic
        font.color.rgb = hex_to_rgb(color_hex)

    def _no_proof(self, run):
        run._r.get_or_add_rPr().append(parse_xml(f'<w:noProof {nsdecls("w")}/>'))

    def _insert_field_code(self, run, field_code: str):
        run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w")))
        run._r.append(parse_xml(
            r'<w:instrText %s xml:space="preserve"> %s </w:instrText>' % (nsdecls("w"), field_code.strip())
        ))
        run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w")))
        run._r.append(parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w")))

    def _new_paragraph(self, align="left", space_before=0, space_after=0, line_spacing=None,
                       first_line_indent_cm=None, keep_with_next=False, style=None):
        p = self.doc.add_paragraph(style=style) if style else self.doc.add_paragraph()
        p.alignment = ALIGN_MAP.get(str(align).lower(), WD_ALIGN_PARAGRAPH.LEFT)
        pf = p.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        pf.line_spacing = line_spacing if line_spacing is not None else self.line_spacing
        pf.first_line_indent = Cm(first_line_indent_cm) if first_line_indent_cm is not None else Cm(0)
        if keep_with_next:
            pf.keep_with_next = True
        return p

    # --------------------------------------------------------------------------
    # Formulas
    # --------------------------------------------------------------------------

    def _init_math_engine(self):
        if getattr(self, "_math_engine_initialized", False):
            return
        self._math_engine_initialized = True
        self._xslt = None
        if not HAS_LXML:
            return
        candidates = [os.path.join(SKILL_DIR, "MML2OMML.XSL"), "MML2OMML.XSL"]
        for root in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if root:
                for sub in ("root\\Office16", "Office16", "Office15"):
                    candidates.append(os.path.join(root, "Microsoft Office", sub, "MML2OMML.XSL"))
        for cand in candidates:
            if os.path.exists(cand):
                try:
                    self._xslt = etree.XSLT(etree.parse(cand))
                    break
                except Exception:
                    pass

    @staticmethod
    def _strip_dollars(formula: str) -> str:
        clean = formula.strip()
        if clean.startswith("$$") and clean.endswith("$$"):
            return clean[2:-2].strip()
        if clean.startswith("$") and clean.endswith("$"):
            return clean[1:-1].strip()
        return clean

    def convert_to_omml(self, formula: str) -> Optional[str]:
        self._init_math_engine()
        if not self._xslt:
            return None
        clean = self._strip_dollars(formula)
        try:
            if "<math" in clean:
                mml = clean
            else:
                if not HAS_LATEX2MATHML:
                    return None
                mml = latex2mathml.converter.convert(clean)
            omml = self._xslt(etree.fromstring(mml))
            return etree.tostring(omml, encoding="utf-8").decode("utf-8")
        except Exception:
            return None

    def _render_latex_to_temp_png(self, formula: str) -> Optional[str]:
        if not HAS_MATPLOTLIB:
            return None
        try:
            path = os.path.join(tempfile.gettempdir(), f"formula_{uuid.uuid4().hex[:8]}.png")
            fig = plt.figure()
            fig.text(0.5, 0.5, f"${self._strip_dollars(formula)}$", fontsize=16, ha="center", va="center")
            fig.patch.set_visible(False)
            plt.axis("off")
            fig.savefig(path, bbox_inches="tight", pad_inches=0.05, dpi=300, transparent=True)
            plt.close(fig)
            return path
        except Exception:
            return None

    # --------------------------------------------------------------------------
    # Inline Markdown
    # --------------------------------------------------------------------------

    def _add_formatted_text(self, paragraph, text: str, base_bold=False, base_italic=False,
                            parse_markdown=True, size_pt=None):
        """**bold**, *italic*, ***bold italic***, `code`, $formula$ (if inline_math), escapes \\* \\$ \\`."""
        if not text:
            return
        size = size_pt or self.font_size_body

        def plain(chunk, bold=base_bold, italic=base_italic):
            r = paragraph.add_run(_unescape(chunk))
            self._apply_run_font(r, size_pt=size, bold=bold, italic=italic)

        if not parse_markdown or not re.search(r"[`*$\\]", text):
            plain(text)
            return

        for k, v in _ESCAPES.items():
            text = text.replace(k, v)
        regex = MD_RE_MATH if self.config["inline_math"] else MD_RE_NOMATH

        last = 0
        for m in regex.finditer(text):
            if m.start() > last:
                plain(text[last:m.start()])
            group, val = m.lastgroup, m.group(m.lastgroup)
            if group == "code":
                r = paragraph.add_run(_unescape(val[1:-1]))
                self._apply_run_font(r, font_name=self.config["code_font_name"], size_pt=max(10.0, size - 1.0))
                self._no_proof(r)
            elif group == "math":
                expr = _unescape(val[1:-1])
                omml = self.convert_to_omml(expr)
                if omml:
                    paragraph._p.append(parse_xml(omml))
                else:
                    plain(expr, italic=True)
            elif group == "bold_italic":
                plain(val[3:-3], bold=True, italic=True)
            elif group == "bold":
                plain(val[2:-2], bold=True)
            elif group == "italic":
                plain(val[1:-1], italic=True)
            last = m.end()
        if last < len(text):
            plain(text[last:])

    # --------------------------------------------------------------------------
    # Caption numbering and references
    # --------------------------------------------------------------------------

    def _number_for(self, kind: str, key: Optional[str]) -> str:
        """Returns the number for a new object: the one reserved for key or the next free one."""
        if key and key in self._refs[kind]:
            return self._refs[kind][key]
        self._counters[kind] += 1
        n = self._counters[kind]
        num = f"{self._h1_count}.{n}" if self.config["caption_numbering"] == "chapter" else str(n)
        if key:
            self._refs[kind][key] = num
        return num

    def ref(self, kind: str, key: str) -> str:
        """
        Figure/table/listing number by key: doc.ref("fig", "schema") -> "5".
        Can be called before the object is inserted: the number is reserved in advance.
        """
        kind = KIND_ALIASES[kind]
        return self._number_for(kind, key)

    def _make_caption(self, kind: str, caption: Optional[str], key: Optional[str]) -> Optional[str]:
        if not caption:
            return caption
        if not self.config["caption_autonumber"]:
            return caption
        label = self.config[{"fig": "figure_label", "tab": "table_label", "lst": "listing_label"}[kind]]
        if re.match(rf"\s*{re.escape(label)}\s*\d", caption):
            return caption  # number already written manually
        num = self._number_for(kind, key)
        # order check: numbers must increase
        flat = int(num.split(".")[-1])
        prev = self._last_num[kind]
        if flat < prev and self.config["caption_numbering"] == "global":
            self.warnings.append(f"{label} {num} inserted after number {prev}: check the order of doc.ref()")
        self._last_num[kind] = max(prev, flat)
        return f"{label} {num}{self.config['caption_separator']}{caption}"

    def _caption_paragraph(self, text: str, position: str, align: str, italic: bool):
        above = position == "above"
        p = self._new_paragraph(
            align=align,
            space_before=6 if above else 3,
            space_after=3 if above else 8,
            keep_with_next=above,
        )
        self._add_formatted_text(p, text, base_italic=italic)
        return p

    # Backward compatibility
    def _create_caption_paragraph(self, caption, position="below", align="center", italic=True):
        return self._caption_paragraph(caption, position.lower(), align, italic)

    # --------------------------------------------------------------------------
    # TEXT AND STRUCTURE
    # --------------------------------------------------------------------------

    def add_heading(self, text: str, level: int = 1, page_break_before: Optional[bool] = None,
                    align: Optional[str] = None):
        """Heading H1-H3 (keep_with_next). H1 starts on a new page automatically, except the first one."""
        should_break = page_break_before
        if should_break is None:
            should_break = level == 1 and self.h1_page_break_before and self._h1_count > 0
        if level == 1:
            self._h1_count += 1
            if self.config["caption_numbering"] == "chapter":
                self._counters = {"fig": 0, "tab": 0, "lst": 0}
                self._last_num = {"fig": 0, "tab": 0, "lst": 0}
        sizes = {1: self.font_size_h1, 2: self.font_size_h2, 3: self.font_size_h3}
        style = f"Heading {level}"
        if self._has_style(style):
            p = self.doc.add_paragraph(style=style)
        else:  # the template has no heading style: a plain paragraph with an outline level for the TOC
            p = self.doc.add_paragraph()
            p._p.get_or_add_pPr().append(parse_xml(f'<w:outlineLvl {nsdecls("w")} w:val="{level - 1}"/>'))
        p.paragraph_format.first_line_indent = Cm(0)
        aligns = {1: self.heading_h1_align, 2: self.heading_h2_align, 3: self.heading_h3_align}
        p.alignment = ALIGN_MAP.get(str(align or aligns.get(level, self.heading_align)).lower(),
                                    WD_ALIGN_PARAGRAPH.LEFT)
        p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
        p.paragraph_format.space_after = Pt(10 if level == 1 else 6)
        p.paragraph_format.keep_with_next = True
        # A paragraph property rather than a separate paragraph with a break: no empty pages appear
        if should_break and not _is_break(self._prev_element(p._p)):
            p.paragraph_format.page_break_before = True
        r = p.add_run(text)
        self._apply_run_font(r, size_pt=sizes.get(level, self.font_size_h2), bold=True)
        return p

    def add_paragraph(self, text: str, bold=False, italic=False, align: Optional[str] = None,
                      indent: bool = True, space_after_pt: float = 0, parse_markdown: bool = True,
                      keep_with_next: bool = False):
        """Paragraph with first-line indent, alignment from the config and inline Markdown."""
        p = self._new_paragraph(
            align=align or self.paragraph_align,
            space_after=space_after_pt,
            first_line_indent_cm=self.paragraph_indent_cm if indent else 0,
            keep_with_next=keep_with_next,
        )
        self._add_formatted_text(p, text, base_bold=bold, base_italic=italic, parse_markdown=parse_markdown)
        return p

    def add_line(self, text: str = "", bold=False, italic=False, size: Optional[float] = None,
                 align: str = "center", space_before: float = 0, space_after: float = 0,
                 parse_markdown: bool = True, keep_with_next: bool = False):
        """Line without first-line indent: title page, signature block, 'Київ — 2026'."""
        p = self._new_paragraph(align=align, space_before=space_before, space_after=space_after,
                                keep_with_next=keep_with_next)
        self._add_formatted_text(p, text, base_bold=bold, base_italic=italic,
                                 parse_markdown=parse_markdown, size_pt=size)
        return p

    def add_formatted_paragraph(self, runs: List[Tuple[Any, ...]], align: Optional[str] = None,
                                indent: bool = True, space_after_pt: float = 0):
        """runs = [("text", bold, italic, font_name, size_pt, color_hex), ...]"""
        p = self._new_paragraph(
            align=align or self.paragraph_align,
            space_after=space_after_pt,
            first_line_indent_cm=self.paragraph_indent_cm if indent else 0,
        )
        for item in runs:
            r = p.add_run(str(item[0]) if item else "")
            self._apply_run_font(
                r,
                font_name=item[3] if len(item) > 3 and item[3] else None,
                size_pt=float(item[4]) if len(item) > 4 and item[4] else None,
                bold=bool(item[1]) if len(item) > 1 else False,
                italic=bool(item[2]) if len(item) > 2 else False,
                color_hex=item[5] if len(item) > 5 and item[5] else "000000",
            )
        return p

    def add_bullet(self, text: str, bullet_char: str = "•", parse_markdown: bool = True,
                   align: Optional[str] = None):
        """Bulleted list item (Word List Bullet). If the template lacks the style, a paragraph with «–»."""
        if self._has_style("List Bullet"):
            p = self._new_paragraph(align=align or self.bullet_align, space_after=2, style="List Bullet")
            p.paragraph_format.first_line_indent = None
            p.paragraph_format.left_indent = Cm(self.paragraph_indent_cm)
            self._add_formatted_text(p, text, parse_markdown=parse_markdown)
        else:
            p = self._hanging_item("–", text, align or self.bullet_align, parse_markdown)
        return p

    def _new_list_num_id(self) -> Optional[int]:
        """Creates a new w:num with startOverride=1 based on the abstract list of the List Number style."""
        try:
            numbering = self.doc.part.numbering_part.element
            style_numpr = self.doc.styles["List Number"].element.pPr.numPr
            base_id = int(style_numpr.numId.val)
            abstract_id = None
            for num in numbering.findall(qn("w:num")):
                if int(num.get(qn("w:numId"))) == base_id:
                    abstract_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
                    break
            if abstract_id is None:
                return None
            new_id = max(int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))) + 1
            numbering.append(parse_xml(
                f'<w:num {nsdecls("w")} w:numId="{new_id}">'
                f'<w:abstractNumId w:val="{abstract_id}"/>'
                f'<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride>'
                f'</w:num>'
            ))
            return new_id
        except Exception:
            return None

    def add_numbered(self, text: str, parse_markdown: bool = True, align: Optional[str] = None,
                     restart: Optional[bool] = None):
        """
        Numbered list item. restart=None: a new list starts from 1 automatically
        if the previous document element was not a numbered list item.
        restart=True/False sets this explicitly.
        """
        if restart is None:
            restart = not self._is_numbered_item(self._prev_element())
        manual = str(self.config["numbered_mode"]).lower() == "manual" or not self._has_style("List Number")

        if manual:
            self._manual_num = 1 if restart else self._manual_num + 1
            p = self._hanging_item(f"{self._manual_num})", text, align or self.numbered_align, parse_markdown)
        else:
            if restart or self._list_num_id is None:
                self._list_num_id = self._new_list_num_id()
            p = self._new_paragraph(align=align or self.numbered_align, space_after=2, style="List Number")
            p.paragraph_format.first_line_indent = None
            p.paragraph_format.left_indent = Cm(self.paragraph_indent_cm)
            if self._list_num_id is not None:
                numPr = p._p.get_or_add_pPr().get_or_add_numPr()  # respects the schema element order
                numPr.get_or_add_ilvl().val = 0
                numPr.get_or_add_numId().val = self._list_num_id
            self._add_formatted_text(p, text, parse_markdown=parse_markdown)
        return p

    def _has_style(self, name: str) -> bool:
        try:
            self.doc.styles[name]
            return True
        except KeyError:
            return False

    def _hanging_item(self, marker: str, text: str, align: str, parse_markdown: bool):
        """List item without a Word style: marker, tab, hanging indent."""
        p = self._new_paragraph(align=align, space_after=2)
        p.paragraph_format.left_indent = Cm(self.paragraph_indent_cm + 0.6)
        p.paragraph_format.first_line_indent = Cm(-0.6)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(self.paragraph_indent_cm + 0.6))
        r = p.add_run(marker + "\t")
        self._apply_run_font(r)
        self._add_formatted_text(p, text, parse_markdown=parse_markdown)
        return p

    def add_placeholder(self, text: str):
        """Red placeholder '[ЗАГЛУШКА: ...]'. Recorded in doc.todo; save() warns about it."""
        label = text if text.startswith("[") else f"[ЗАГЛУШКА: {text}]"
        p = self._new_paragraph(align="center", space_before=6, space_after=6)
        r = p.add_run(label)
        self._apply_run_font(r, bold=True, color_hex="C00000")
        self.todo.append(label)
        return p

    # --------------------------------------------------------------------------
    # TITLE PAGE
    # --------------------------------------------------------------------------

    def add_title_page(self, fields: Optional[Dict[str, str]] = None, **kwargs):
        """
        Title page from fields (arguments override the "title" section of the config).
        Fields: ministry, university, faculty, department, work_type, number, discipline, topic,
        variant, student_label, group_label, group, student, teacher_label, teacher, city, year.
        Empty fields are skipped. Call it first: the title page occupies its own section,
        whose content is spread evenly over the page height.
        """
        f = dict(self.config.get("title", {}))
        f.update(fields or {})
        f.update(kwargs)
        g = lambda k, d="": str(f.get(k, d) or "").strip()

        def block(lines, align, left_indent_cm=0.0):
            lines = [l for l in lines if l and l[0]]
            if not lines:
                return
            p = self._new_paragraph(align=align, line_spacing=self.line_spacing)
            if left_indent_cm:
                p.paragraph_format.left_indent = Cm(left_indent_cm)
            for i, (txt, bold, size) in enumerate(lines):
                if i:
                    p.add_run().add_break()
                self._add_formatted_text(p, txt, base_bold=bold, size_pt=size, parse_markdown=False)

        size = self.font_size_body
        block([(g("ministry"), False, size), (g("university"), False, size),
               (g("faculty"), False, size), (g("department"), False, size)], "center")

        number = g("number")
        work = g("work_type", "Лабораторна робота").upper()
        heading = f"{work} №{number}" if number else work
        block([
            (g("report_word", "ЗВІТ"), True, size + 2),
            (heading, True, size),
            (f"з дисципліни «{g('discipline')}»" if g("discipline") else "", False, size),
            (f"на тему: «{g('topic')}»" if g("topic") else "", False, size),
            (f"Варіант {g('variant')}" if g("variant") else "", False, size),
        ], "center")

        student = g("student")
        teacher = g("teacher")
        sig = []
        if student:
            sig.append((g("student_label", "Виконав:"), False, size))
            if g("group"):
                sig.append((f"{g('group_label', 'студент групи')} {g('group')}", False, size))
            sig.append((student, False, size))
        if teacher:
            if sig:
                sig.append((" ", False, size))
            sig.append((g("teacher_label", "Перевірив:"), False, size))
            sig.append((teacher, False, size))
        block(sig, "left", left_indent_cm=float(f.get("signature_indent_cm", 9.0)))

        city_year = " — ".join(x for x in (g("city"), g("year")) if x)
        block([(city_year, False, size)], "center")

        # Separate section: vertical alignment "both" spreads 4 blocks over the page height
        new_sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        title_sec = self.doc.sections[-2]
        title_sec.different_first_page_header_footer = True  # no page number on the title page
        title_sec._sectPr.find(qn("w:titlePg")).addprevious(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="both"/>'))
        new_sec.different_first_page_header_footer = False
        return title_sec

    # --------------------------------------------------------------------------
    # IMAGES
    # --------------------------------------------------------------------------

    def _find_image(self, name: str) -> Optional[str]:
        if os.path.isfile(name):
            return name
        dirs = [os.getcwd()]
        if self.config_path:
            cfg_dir = os.path.dirname(os.path.abspath(self.config_path))
            dirs += [cfg_dir, os.path.dirname(cfg_dir)]
        if sys.argv and sys.argv[0]:
            main_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
            dirs += [main_dir, os.path.dirname(main_dir)]
        dirs.append(SKILL_DIR)
        for d in dirs:
            for cand in (os.path.join(d, self.images_dir, name), os.path.join(d, name)):
                if os.path.isfile(cand):
                    return cand
        return None

    def _image_size_cm(self, path: str, width_cm, max_width_cm, max_height_cm) -> Tuple[float, Optional[float]]:
        max_w = min(max_width_cm or self.text_width_cm, self.text_width_cm)
        max_h = max_height_cm or self.text_height_cm
        if not HAS_PIL:
            return (min(width_cm, max_w) if width_cm else max_w), None
        with Image.open(path) as im:
            w_px, h_px = im.size
        ratio = h_px / w_px
        if width_cm:
            w = width_cm
        elif str(self.config["image_default_size"]).lower() == "width":
            w = max_w
        else:
            w = w_px * 2.54 / 96
        w = min(w, max_w)
        if w * ratio > max_h:
            w = max_h / ratio
        return w, w * ratio

    def add_image(self, image_name_or_path: str, caption: Optional[str] = None, width_cm: Optional[float] = None,
                  caption_position: Optional[str] = None, caption_align: Optional[str] = None,
                  max_width_cm: Optional[float] = None, max_height_cm: Optional[float] = None,
                  key: Optional[str] = None):
        """
        Image from images/ (or by path). Defaults to real size (96 dpi), but no wider than
        the text and no taller than image_max_height_cm. The caption is numbered automatically (caption_autonumber).
        """
        pos = (caption_position or self.image_caption_position).lower()
        aln = (caption_align or self.image_caption_align).lower()
        italic = bool(self.config["image_caption_italic"])
        caption_text = self._make_caption("fig", caption, key)
        path = self._find_image(image_name_or_path)

        if caption_text and pos == "above":
            self._caption_paragraph(caption_text, "above", aln, italic)

        if not path:
            if self.config["strict_images"]:
                raise FileNotFoundError(f"Image not found: {image_name_or_path}")
            self.missing.append(image_name_or_path)
            p_img = self._new_paragraph(align="center", space_before=6, space_after=2,
                                        keep_with_next=(pos == "below"))
            r = p_img.add_run(f"[ЗОБРАЖЕННЯ ВІДСУТНЄ: {image_name_or_path}]")
            self._apply_run_font(r, bold=True, color_hex="C00000")
        else:
            w, _ = self._image_size_cm(path, width_cm, max_width_cm or self.config["image_max_width_cm"],
                                       max_height_cm or self.config["image_max_height_cm"])
            p_img = self._new_paragraph(align="center", space_before=6, space_after=2 if pos == "below" else 6,
                                        line_spacing=1.0, keep_with_next=(pos == "below"))
            p_img.add_run().add_picture(path, width=Cm(w))

        if caption_text and pos == "below":
            self._caption_paragraph(caption_text, "below", aln, italic)
        return p_img

    # --------------------------------------------------------------------------
    # TABLES
    # --------------------------------------------------------------------------

    def add_table(self, headers: List[str], rows: List[List[Any]], caption: Optional[str] = None,
                  col_widths_cm: Optional[List[float]] = None, caption_position: Optional[str] = None,
                  caption_align: Optional[str] = None, col_align: Optional[List[str]] = None,
                  font_size: Optional[float] = None, key: Optional[str] = None):
        """
        Table with borders, repeated header row and cantSplit. The header is always centered,
        data alignment: col_align ["left", "center", "right"] or the table_cell_align key.
        """
        pos = (caption_position or self.table_caption_position).lower()
        aln = (caption_align or self.table_caption_align).lower()
        italic = bool(self.config["table_caption_italic"])
        size = float(font_size or self.config["table_font_size"])
        spacing = float(self.config["table_line_spacing"])
        default_align = str(self.config["table_cell_align"])
        caption_text = self._make_caption("tab", caption, key)

        if caption_text and pos == "above":
            self._caption_paragraph(caption_text, "above", aln, italic)

        n_cols = len(headers)
        table = self.doc.add_table(rows=len(rows) + 1, cols=n_cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        if col_widths_cm and len(col_widths_cm) == n_cols:
            widths = [Cm(w) for w in col_widths_cm]
        else:
            widths = [Cm(self.text_width_cm / n_cols)] * n_cols
        aligns = list(col_align) if col_align and len(col_align) == n_cols else [default_align] * n_cols

        def fill(cell, text, c_idx, bold, align):
            cell.width = widths[c_idx]
            self._set_cell_borders(cell, self.table_border_color, self.table_border_sz)
            self._set_cell_margins(cell)
            p = cell.paragraphs[0]
            p.alignment = ALIGN_MAP.get(str(align).lower(), WD_ALIGN_PARAGRAPH.CENTER)
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.first_line_indent = Cm(0)
            pf.line_spacing = spacing
            self._add_formatted_text(p, "" if text is None else str(text), base_bold=bold, size_pt=size)

        header_row = table.rows[0]
        self._set_repeat_header(header_row)
        self._set_cant_split(header_row)
        for c_idx, text in enumerate(headers):
            fill(header_row.cells[c_idx], text, c_idx, True, "center")
        for r_idx, row_data in enumerate(rows):
            row = table.rows[r_idx + 1]
            self._set_cant_split(row)
            for c_idx in range(n_cols):
                val = row_data[c_idx] if c_idx < len(row_data) else ""
                fill(row.cells[c_idx], val, c_idx, False, aligns[c_idx])

        if caption_text and pos == "below":
            self._caption_paragraph(caption_text, "below", aln, italic)
        elif float(self.config["table_space_after_pt"]) > 0:
            sp = self._new_paragraph(space_after=float(self.config["table_space_after_pt"]))
            sp.paragraph_format.line_spacing = Pt(1)
        return table

    # --------------------------------------------------------------------------
    # LISTINGS
    # --------------------------------------------------------------------------

    def add_code_block(self, code: str, caption: Optional[str] = None, key: Optional[str] = None,
                       keep_together: Optional[bool] = None, keep_with_next: bool = False,
                       caption_position: Optional[str] = None):
        """
        Listing: monospace font code_font_size, single spacing, hanging indent for
        wrapped lines, noProof. A short listing (up to code_keep_together_lines lines)
        is not split across pages. keep_with_next=True keeps it together with the next block.
        """
        c = self.config
        pos = (caption_position or c["listing_caption_position"]).lower()
        aln = str(c["listing_caption_align"]).lower()
        italic = bool(c["listing_caption_italic"])
        caption_text = self._make_caption("lst", caption, key)
        lines = code.rstrip("\n").split("\n")
        if keep_together is None:
            keep_together = len(lines) <= int(c["code_keep_together_lines"])
        hang = float(c["code_hanging_indent_cm"])
        # A template style (e.g. "Код исх") takes priority: then it sets the font and indents
        code_style = c.get("code_style")
        use_style = bool(code_style) and self._has_style(code_style)
        if code_style and not use_style:
            self.warnings.append(f"listing style '{code_style}' not found, using code_font_name")

        if caption_text and pos == "above":
            self._caption_paragraph(caption_text, "above", aln, italic)

        last = len(lines) - 1
        for i, line in enumerate(lines):
            is_last = i == last
            if use_style:
                # the template style sets font and indents, but each code line is a separate paragraph,
                # so the spacing between them is zeroed, otherwise the listing "falls apart"
                p = self.doc.add_paragraph(style=code_style)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(6 if (is_last and pos != "below") else 0)
            else:
                p = self._new_paragraph(align="left", line_spacing=1.0,
                                        space_after=6 if (is_last and pos != "below") else 0)
                p.paragraph_format.left_indent = Cm(hang)
                p.paragraph_format.first_line_indent = Cm(-hang)
            pf = p.paragraph_format
            pf.widow_control = True
            if not is_last and (keep_together or i < 2 or pos == "below" and i >= last - 2):
                pf.keep_with_next = True
            if is_last and (keep_with_next or pos == "below"):
                pf.keep_with_next = True
            r = p.add_run(line.replace("\t", "    ") if line.strip() else " ")
            if not use_style:
                self._apply_run_font(r, font_name=c["code_font_name"], size_pt=float(c["code_font_size"]))
            self._no_proof(r)

        if caption_text and pos == "below":
            cap = self._caption_paragraph(caption_text, "below", aln, italic)
            if keep_with_next:
                cap.paragraph_format.keep_with_next = True
        return p

    def add_listing_from_file(self, path: str, start: Optional[str] = None, end: Optional[str] = None,
                              strip_comments: bool = True, caption: Optional[str] = None,
                              lang: Optional[str] = None, **kwargs):
        """
        Listing from a file. start/end: fragment from the first line containing start up to the line with end
        (the end line is excluded). Comments are removed with quotes taken into account
        (SQL, Python, C-like languages, shell). Other arguments are passed to add_code_block.
        """
        with open(path, "r", encoding="utf-8-sig") as f:
            text = f.read()
        lines = text.split("\n")
        i0, i1 = 0, len(lines)
        if start is not None:
            i0 = next((i for i, l in enumerate(lines) if start in l), None)
            if i0 is None:
                raise ValueError(f"Marker start={start!r} not found in {path}")
        if end is not None:
            i1 = next((i for i, l in enumerate(lines) if i > i0 and end in l), None)
            if i1 is None:
                raise ValueError(f"Marker end={end!r} not found in {path}")
        fragment = "\n".join(lines[i0:i1])
        if strip_comments:
            fragment = strip_code_comments(fragment, lang or os.path.splitext(path)[1] or ".txt")
        fragment = textwrap.dedent(fragment).strip("\n")
        return self.add_code_block(fragment, caption=caption, **kwargs)

    # --------------------------------------------------------------------------
    # FORMULAS, TOC, BREAKS
    # --------------------------------------------------------------------------

    def add_formula(self, formula: str, number: Optional[str] = None,
                    space_before_pt: float = 6, space_after_pt: float = 8):
        """LaTeX/MathML formula as a native Word object. number='1.1' -> number (1.1) on the right."""
        omml = self.convert_to_omml(formula)
        png = None if omml else self._render_latex_to_temp_png(formula)

        def put(p):
            if omml:
                p._p.append(parse_xml(omml))
            elif png:
                p.add_run().add_picture(png)
            else:
                r = p.add_run(self._strip_dollars(formula))
                self._apply_run_font(r, italic=True)

        try:
            if number:
                width = self.text_width_cm
                table = self.doc.add_table(rows=1, cols=2)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                table.autofit = False
                c0, c1 = table.rows[0].cells
                c0.width, c1.width = Cm(width - 2.5), Cm(2.5)
                p0 = c0.paragraphs[0]
                p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p1 = c1.paragraphs[0]
                p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                for p in (p0, p1):
                    p.paragraph_format.space_before = Pt(space_before_pt)
                    p.paragraph_format.space_after = Pt(space_after_pt)
                    p.paragraph_format.first_line_indent = Cm(0)
                put(p0)
                r1 = p1.add_run(f"({number})")
                self._apply_run_font(r1)
                result = table
            else:
                p = self._new_paragraph(align="center", space_before=space_before_pt, space_after=space_after_pt)
                put(p)
                result = p
        finally:
            if png and os.path.exists(png):
                try:
                    os.remove(png)
                except OSError:
                    pass
        return result

    def add_toc(self, title: str = "ЗМІСТ"):
        """Automatic table of contents field. Word fills it during save(pdf=True) or manually (F9)."""
        if title:
            p = self._new_paragraph(align="center", space_after=12, keep_with_next=True)
            r = p.add_run(title)
            self._apply_run_font(r, size_pt=self.font_size_h1, bold=True)
        # After updating, Word leaves this paragraph empty under the TOC entries.
        # Zero height keeps it from spilling onto a separate empty page.
        p = self._new_paragraph(space_before=0, space_after=0)
        p.paragraph_format.line_spacing = Pt(1)
        self._insert_field_code(p.add_run(), r'TOC \o "1-3" \h \z \u')
        if self.config["toc_update_on_open"]:
            self._set_update_fields()
        return p

    def add_page_break(self):
        """Page break. Consecutive breaks and a break before H1 are not duplicated."""
        if _is_break(self._prev_element()):
            return
        self.doc.add_page_break()

    def _drop_trailing_break(self):
        """Removes a page break at the very end of the document (otherwise the last page is empty)."""
        body = self.doc.element.body
        last = body[-2] if len(body) > 1 and body[-1].tag == qn("w:sectPr") else body[-1]
        if _is_page_break(last):
            body.remove(last)

    def _prev_element(self, el=None):
        """
        Document element before el (or before where the next element will go if el=None).
        Inside insert_at() new elements temporarily sit at the end of the document, so for the first
        of them the anchor's neighbour counts as previous, not the end of the document.
        """
        body = self.doc.element.body
        if el is None:
            children = [c for c in body if c.tag != qn("w:sectPr")]
            prev = children[-1] if children else None
        else:
            prev = el.getprevious()
        if self._insert_ctx is not None:
            before, anchor, after = self._insert_ctx
            if prev is None or prev in before:
                return anchor if after else anchor.getprevious()
        return prev

    def _is_numbered_item(self, el) -> bool:
        """Whether el is an item of the current numbered list (native or «1)»)."""
        if el is None or el.tag != qn("w:p"):
            return False
        num = el.find(f"{qn('w:pPr')}/{qn('w:numPr')}/{qn('w:numId')}")
        if num is not None and self._list_num_id is not None and int(num.get(qn("w:val"))) == self._list_num_id:
            return True
        # Paragraph.text accounts for <w:tab/>, into which python-docx turns the "\t" of the «1)\t» marker
        return bool(re.match(r"\d+\)\t", Paragraph(el, self.doc._body).text))

    # --------------------------------------------------------------------------
    # WORKING WITH A TEMPLATE
    # --------------------------------------------------------------------------

    def _body_paragraphs(self):
        return [Paragraph(el, self.doc._body) for el in self.doc.element.body if el.tag == qn("w:p")]

    def find_paragraph(self, text: str, style_prefix: Optional[str] = None, exact: bool = False):
        """
        First body paragraph containing text (exact=True — matches entirely,
        ignoring surrounding whitespace). style_prefix="Heading" limits the search to headings.
        Raises LookupError if nothing is found.
        """
        for p in self._body_paragraphs():
            t = p.text.strip()
            if (t == text.strip()) if exact else (text in t):
                if style_prefix is None or (p.style is not None and p.style.name.startswith(style_prefix)):
                    return p
        raise LookupError(f"Paragraph not found: {text!r}" + (f" (style {style_prefix}*)" if style_prefix else ""))

    def _as_element(self, anchor):
        if isinstance(anchor, str):
            anchor = self.find_paragraph(anchor)
        for attr in ("_p", "_tbl"):
            if hasattr(anchor, attr):
                return getattr(anchor, attr)
        return anchor

    def _heading_level(self, el) -> Optional[int]:
        try:
            p = Paragraph(el, self.doc._body)
            name = p.style.name if p.style is not None else ""
        except Exception:
            name = ""
        m = re.match(r"Heading (\d)", name)
        if m:
            return int(m.group(1))
        lvl = el.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}")
        return int(lvl.get(qn("w:val"))) + 1 if lvl is not None else None

    def clear_section(self, start, end=None):
        """
        Deletes content between heading start and end (both remain). start/end — text or a paragraph.
        If end is not given, deletes up to the next heading of the same or higher level.
        Section breaks (sectPr) are not deleted. Returns the boundary paragraph (or None at the end of
        the document), so you can write right away: with doc.insert_at(doc.clear_section("Хід роботи")): ...
        """
        start_el = self._as_element(start)
        end_el = self._as_element(end) if end is not None else None
        level = self._heading_level(start_el) or 1
        el = start_el.getnext()
        while el is not None and el is not end_el and el.tag != qn("w:sectPr"):
            if end_el is None and el.tag == qn("w:p"):
                lvl = self._heading_level(el)
                if lvl is not None and lvl <= level:
                    end_el = el
                    break
            nxt = el.getnext()
            if el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None:
                el.getparent().remove(el)
            el = nxt
        return Paragraph(end_el, self.doc._body) if end_el is not None else None

    @contextlib.contextmanager
    def insert_at(self, anchor, after: bool = False):
        """
        All add_* calls inside the block insert elements before anchor (after=True — after it)
        instead of at the end of the document. anchor — paragraph text, paragraph, table or None (end of document).
            with doc.insert_at("Висновки"):
                doc.add_paragraph("...")
        Do not call the title page (add_title_page) inside the block.
        """
        body = self.doc.element.body
        before = set(body)  # the elements themselves, not id(): lxml proxies must stay alive
        anchor_el = self._as_element(anchor) if anchor is not None else None
        if anchor_el is not None:
            self._insert_ctx = (before, anchor_el, after)
        try:
            yield
        finally:
            self._insert_ctx = None
            if anchor_el is not None:
                cursor = anchor_el
                new = [e for e in body if e not in before and e.tag != qn("w:sectPr")]
                for e in new:
                    if after:
                        cursor.addnext(e)
                        cursor = e
                    else:
                        cursor.addprevious(e)

    def _text_roots(self):
        roots = [self.doc.element.body]
        for s in self.doc.sections:
            for hf in (s.header, s.footer, s.first_page_header, s.first_page_footer,
                       s.even_page_header, s.even_page_footer):
                try:
                    if not hf.is_linked_to_previous:
                        roots.append(hf._element)
                except Exception:
                    pass
        return roots

    @staticmethod
    def _nearest_p(el):
        el = el.getparent()
        while el is not None and el.tag != qn("w:p"):
            el = el.getparent()
        return el

    def replace_text(self, old: str, new: str) -> int:
        """
        Replaces text in the whole document: body, tables, text boxes, headers/footers,
        including when Word split old into several runs. Formatting is taken from the first
        run of the match. Returns the number of replacements.
        """
        if not old:
            return 0
        count = 0
        done = set()
        for root in self._text_roots():
            if id(root) in done:
                continue
            done.add(id(root))
            for p in root.iter(qn("w:p")):
                ts = [t for t in p.iter(qn("w:t")) if self._nearest_p(t) is p]
                while ts:
                    full = "".join(t.text or "" for t in ts)
                    a = full.find(old)
                    if a < 0:
                        break
                    b, pos, first = a + len(old), 0, True
                    for t in ts:
                        txt = t.text or ""
                        s0, s1 = pos, pos + len(txt)
                        pos = s1
                        if s1 <= a or s0 >= b:
                            continue
                        lo, hi = max(a, s0) - s0, min(b, s1) - s0
                        t.text = txt[:lo] + (new if first else "") + txt[hi:]
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        first = False
                    count += 1
                    if old in new:
                        break  # guard against endless replacement
        return count

    def fill_table(self, table, rows: List[List[Any]], start_row: int = 1, clear_highlight: bool = True,
                   font_size: Optional[float] = None, col_align: Optional[List[str]] = None):
        """
        Fills an existing template table (a Table object or a table index in the document, from 0),
        starting from row start_row. If rows are missing, copies the last one (with borders and width).
        Cell text is formatted as in add_table; highlighting is removed.
        """
        if isinstance(table, int):
            table = self.doc.tables[table]
        size = float(font_size or self.config["table_font_size"])
        spacing = float(self.config["table_line_spacing"])
        default_align = str(self.config["table_cell_align"])
        while len(table.rows) < start_row + len(rows):
            table._tbl.append(copy.deepcopy(table.rows[-1]._tr))
        for r_idx, data in enumerate(rows):
            row = table.rows[start_row + r_idx]
            for c_idx, cell in enumerate(row.cells):
                val = data[c_idx] if c_idx < len(data) else ""
                ps = cell.paragraphs
                for extra in ps[1:]:
                    extra._p.getparent().remove(extra._p)
                p = ps[0]
                for child in list(p._p):
                    if child.tag != qn("w:pPr"):
                        p._p.remove(child)
                align = col_align[c_idx] if col_align and c_idx < len(col_align) else default_align
                p.alignment = ALIGN_MAP.get(str(align).lower(), WD_ALIGN_PARAGRAPH.CENTER)
                pf = p.paragraph_format
                pf.first_line_indent = Cm(0)
                pf.line_spacing = spacing
                pf.space_before = Pt(0)
                pf.space_after = Pt(0)
                self._add_formatted_text(p, "" if val is None else str(val), size_pt=size)
        if clear_highlight:
            for el in list(table._tbl.iter(qn("w:highlight"))):
                el.getparent().remove(el)
        return table

    def _set_update_fields(self):
        """Asks Word to update fields (TOC) when the document is opened."""
        settings = self.doc.settings.element
        if settings.find(qn("w:updateFields")) is not None:
            return
        el = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
        later = ("hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars", "rsids", "mathPr",
                 "attachedSchema", "themeFontLang", "clrSchemeMapping", "doNotIncludeSubdocsInStats",
                 "doNotAutoCompressPictures", "forceUpgrade", "captions", "readModeInkLockDown",
                 "smartTagType", "schemaLibrary", "shapeDefaults", "doNotEmbedSmartTags",
                 "decimalSymbol", "listSeparator")
        for child in settings:
            if isinstance(child.tag, str) and child.tag.split("}")[-1] in later:
                child.addprevious(el)
                return
        settings.append(el)

    # --------------------------------------------------------------------------
    # CHECKS
    # --------------------------------------------------------------------------

    def check(self, verbose: bool = True) -> List[str]:
        """
        Document check (save() calls it itself): leftover template placeholder text
        (template_placeholders) and highlighted text, consecutive empty paragraphs, a table caption
        without a table after it, tables wider than the page margins. Returns a list of problems.
        """
        problems: List[str] = []
        body = self.doc.element.body
        patterns = [re.compile(x) for x in self.config.get("template_placeholders") or []]
        tab_re = re.compile(rf"^\s*{re.escape(self.config['table_label'])}\s+\d")
        math_tag = "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"
        empty_run = 0
        for el in body:
            if el.tag == qn("w:p"):
                text = "".join(t.text or "" for t in el.iter(qn("w:t"))).strip()
                special = any(el.find(".//" + qn(x)) is not None
                              for x in ("w:br", "w:drawing", "w:sectPr", "w:fldChar", "w:pict"))
                special = special or el.find(".//" + math_tag) is not None
                if not text and not special:
                    empty_run += 1
                    if empty_run == 2:
                        problems.append(f"consecutive empty paragraphs after: {self._prev_text(el)[:50]!r}")
                else:
                    empty_run = 0
                for rx in patterns:
                    if text and rx.search(text):
                        problems.append(f"placeholder text ({rx.pattern}): {text[:60]!r}")
                        break
                if tab_re.match(text) and self.table_caption_position == "above":
                    nxt = el.getnext()
                    if nxt is None or nxt.tag != qn("w:tbl"):
                        problems.append(f"caption without a table after it: {text[:60]!r}")
            else:
                empty_run = 0
            if el.tag == qn("w:tbl"):
                width = self._table_width_cm(el)
                if width and width > self.text_width_cm + 0.1:
                    first = "".join(t.text or "" for t in el.iter(qn("w:t")))[:40]
                    problems.append(f"table wider than the text area ({width:.1f} > {self.text_width_cm:.1f} cm): {first!r}")
        highlighted = sum(1 for _ in body.iter(qn("w:highlight")))
        if highlighted:
            problems.append(f"highlighted text (probably template placeholders): {highlighted} fragm.")
        if verbose:
            for p in problems:
                print(f"[WARNING] {p}")
        return problems

    @staticmethod
    def _prev_text(el) -> str:
        prev = el.getprevious()
        while prev is not None:
            t = "".join(x.text or "" for x in prev.iter(qn("w:t"))).strip()
            if t:
                return t
            prev = prev.getprevious()
        return ""

    @staticmethod
    def _table_width_cm(tbl) -> Optional[float]:
        best = 0
        for tr in tbl.iter(qn("w:tr")):
            total = 0
            for tc in tr.findall(qn("w:tc")):
                w = tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
                if w is None or w.get(qn("w:type")) not in (None, "dxa"):
                    total = 0
                    break
                total += int(w.get(qn("w:w"), "0"))
            best = max(best, total)
        if not best:
            best = sum(int(g.get(qn("w:w"), "0")) for g in tbl.iter(qn("w:gridCol")))
        return best / 567.0 if best else None

    # --------------------------------------------------------------------------
    # SAVING AND PDF
    # --------------------------------------------------------------------------

    def _save_docx(self, file_path: str) -> str:
        try:
            self.doc.save(file_path)
            return file_path
        except PermissionError:
            name = os.path.basename(file_path)
            print(f"[WARNING] File '{name}' is locked (probably open in Word). Saving under another name.")
            base, ext = os.path.splitext(file_path)
            for i in range(0, 100):
                alt = f"{base}_new{ext}" if i == 0 else f"{base}_new_{i}{ext}"
                try:
                    self.doc.save(alt)
                    print(f"[WARNING] Close '{name}' in Word so the next build overwrites the main file.")
                    return alt
                except PermissionError:
                    continue
            alt = os.path.join(tempfile.gettempdir(), f"{os.path.splitext(name)[0]}_{uuid.uuid4().hex[:6]}{ext}")
            self.doc.save(alt)
            return alt

    def save(self, file_path: str, pdf: bool = False) -> str:
        """
        Saves the DOCX (if the file is open in Word, writes to *_new.docx). pdf=True additionally
        converts to PDF with field update (TOC, page numbers): Word COM -> soffice -> docx2pdf.
        Prints a summary: caption counts, missing images, placeholders, warnings.
        Returns the DOCX path; the PDF path is stored in doc.pdf_path.
        """
        d = os.path.dirname(os.path.abspath(file_path))
        os.makedirs(d, exist_ok=True)
        self._drop_trailing_break()
        saved = self._save_docx(file_path)
        print(f"[OK] DOCX: {saved}")
        print(f"     figures: {self._counters['fig']}, tables: {self._counters['tab']}, "
              f"listings: {self._counters['lst']}"
              + (" (per-chapter numbering, counter of the last chapter)" if self.config['caption_numbering'] == 'chapter' else ""))
        for w in self.warnings + self.check(verbose=False):
            print(f"[WARNING] {w}")
        if self.missing:
            print(f"[WARNING] Missing images ({len(self.missing)}): {', '.join(self.missing)}")
        if self.todo:
            print(f"[WARNING] Unfilled placeholders ({len(self.todo)}): " + "; ".join(self.todo))
        if pdf:
            target = os.path.splitext(saved)[0] + ".pdf"
            self.pdf_path = convert_to_pdf(saved, target)
        return saved


    # --------------------------------------------------------------------------
    # Internal OpenXML helper methods
    # --------------------------------------------------------------------------

    def _set_cell_borders(self, cell, border_color: str = "000000", sz: str = "8"):
        c = clean_hex(border_color)
        cell._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            + "".join(f'<w:{s} w:val="single" w:sz="{sz}" w:space="0" w:color="{c}"/>'
                      for s in ("top", "left", "bottom", "right"))
            + "</w:tcBorders>"
        ))

    def _set_cell_margins(self, cell, top_twips=100, bottom_twips=100, left_twips=140, right_twips=140):
        cell._tc.get_or_add_tcPr().append(parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'<w:top w:w="{top_twips}" w:type="dxa"/><w:bottom w:w="{bottom_twips}" w:type="dxa"/>'
            f'<w:left w:w="{left_twips}" w:type="dxa"/><w:right w:w="{right_twips}" w:type="dxa"/>'
            f"</w:tcMar>"
        ))

    def _set_repeat_header(self, row):
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

    def _set_cant_split(self, row):
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))


