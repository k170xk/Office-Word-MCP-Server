"""
Deployment-wide defaults (overridable via environment).

TABLE_HEADER_FILL / TABLE_HEADER_TEXT control brand styling for ``add_table`` and for
defaults on ``highlight_table_header`` when callers omit colors.
"""

import os


def _normalize_docx_hex6(raw: str) -> str:
    s = raw.strip().lstrip("#").upper()
    if len(s) == 6 and all(c in "0123456789ABCDEF" for c in s):
        return s
    return "77C343"


def _normalize_docx_text_hex(raw: str) -> str:
    s = raw.strip().lstrip("#").upper()
    if len(s) == 6 and all(c in "0123456789ABCDEF" for c in s):
        return s
    return "000000"


TABLE_HEADER_FILL_DEFAULT = _normalize_docx_hex6(
    os.getenv("TABLE_HEADER_FILL", "77c343")
)
TABLE_HEADER_TEXT_DEFAULT = _normalize_docx_text_hex(
    os.getenv("TABLE_HEADER_TEXT", "000000")
)
