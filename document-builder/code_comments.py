"""
code_comments.py
Removing comments from source code with quotes taken into account (for add_listing_from_file).
Pure functions without python-docx dependencies: they can be tested and edited separately.
"""

from __future__ import annotations

import io
import re
import tokenize
from typing import Dict, List

# ------------------------------------------------------------------------------
# Removing comments from source code (for add_listing_from_file)
# ------------------------------------------------------------------------------

_LANGS = {
    # extension: (line comment, (block start, end), quotes, escape by doubling)
    ".sql": ("--", ("/*", "*/"), "'\"", True),
    ".c": ("//", ("/*", "*/"), "'\"", False),
    ".h": ("//", ("/*", "*/"), "'\"", False),
    ".cpp": ("//", ("/*", "*/"), "'\"", False),
    ".hpp": ("//", ("/*", "*/"), "'\"", False),
    ".cs": ("//", ("/*", "*/"), "'\"", False),
    ".java": ("//", ("/*", "*/"), "'\"", False),
    ".js": ("//", ("/*", "*/"), "'\"`", False),
    ".ts": ("//", ("/*", "*/"), "'\"`", False),
    ".go": ("//", ("/*", "*/"), "'\"`", False),
    ".rs": ("//", ("/*", "*/"), "\"", False),
    ".kt": ("//", ("/*", "*/"), "'\"", False),
    ".php": ("//", ("/*", "*/"), "'\"", False),
    ".css": (None, ("/*", "*/"), "'\"", False),
    ".sh": ("#", None, "'\"", False),
    ".ps1": ("#", ("<#", "#>"), "'\"", False),
    ".r": ("#", None, "'\"", False),
}


def _strip_python_comments(code: str) -> str:
    """Removes # comments via tokenize, so '#' inside strings is untouched."""
    lines = code.split("\n")
    cuts: Dict[int, int] = {}
    try:
        for tok in tokenize.generate_tokens(io.StringIO(code).readline):
            if tok.type == tokenize.COMMENT:
                cuts[tok.start[0] - 1] = tok.start[1]
    except (tokenize.TokenError, IndentationError, SyntaxError):
        return code
    out = []
    for i, line in enumerate(lines):
        if i in cuts:
            stripped = line[: cuts[i]].rstrip()
            if not stripped:
                continue  # the line consisted only of a comment
            out.append(stripped)
        else:
            out.append(line)
    return "\n".join(out)


def _strip_generic_comments(code: str, line_c, block_c, quotes: str, doubled: bool) -> str:
    out: List[str] = []
    i, n = 0, len(code)
    quote = None
    while i < n:
        c = code[i]
        if quote:
            out.append(c)
            if not doubled and c == "\\" and i + 1 < n:
                out.append(code[i + 1])
                i += 2
                continue
            if c == quote:
                if doubled and i + 1 < n and code[i + 1] == quote:
                    out.append(quote)
                    i += 2
                    continue
                quote = None
            i += 1
            continue
        if c in quotes:
            quote = c
            out.append(c)
            i += 1
            continue
        if line_c and code.startswith(line_c, i):
            while i < n and code[i] != "\n":
                i += 1
            continue
        if block_c and code.startswith(block_c[0], i):
            end = code.find(block_c[1], i + len(block_c[0]))
            i = n if end == -1 else end + len(block_c[1])
            continue
        out.append(c)
        i += 1
    # drop lines that became empty because of a comment, but keep
    # single empty lines that were in the original
    result = []
    orig_blank_prev = False
    for line in "".join(out).split("\n"):
        line = line.rstrip()
        if not line:
            if orig_blank_prev:
                continue
            orig_blank_prev = True
        else:
            orig_blank_prev = False
        result.append(line)
    return "\n".join(result)


def strip_code_comments(code: str, lang: str) -> str:
    """Removes comments from code in language lang ('.sql', '.py', '.cs' ...), respecting quotes."""
    lang = lang.lower() if lang.startswith(".") else "." + lang.lower()
    if lang == ".py":
        text = _strip_python_comments(code)
    elif lang in _LANGS:
        text = _strip_generic_comments(code, *_LANGS[lang])
    else:
        return code
    text = "\n".join(l.rstrip() for l in text.split("\n"))
    return re.sub(r"\n{3,}", "\n\n", text).strip("\n")
