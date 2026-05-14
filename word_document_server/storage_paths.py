"""Helpers for tenant-style paths under persistent storage."""

from __future__ import annotations

import re
from typing import Optional

from word_document_server.utils.file_utils import ensure_docx_extension

# Folder / prefix segments (workspaces): safe for directory names under storage root
_SAFE_SEG = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_-]{0,127}$")

# Final path component before .docx (letters, numbers, common punctuation, spaces allowed)
_SAFE_BASENAME = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_. -]{0,240}$", re.IGNORECASE)


def validate_workspace_segment(name: Optional[str]) -> Optional[str]:
    """
    Validates a single top-level workspace / tenant folder under the disk root.
    Raises ValueError if invalid.
    """
    if name is None or name == "":
        return None
    n = str(name).strip()
    if not _SAFE_SEG.match(n):
        raise ValueError(
            "Invalid workspace ID: use 1–128 chars starting with alphanumeric, "
            "then letters, digits, hyphen, underscore only."
        )
    return n


def normalize_storage_document_key(user_path: str) -> str:
    """
    Map a caller-provided logical document key to a safe POSIX-style relative path
    stored under DISK_PATH (e.g. ``acme-corp/board-report.docx``).

    Raises ValueError on traversal, unsupported characters, or empty paths.
    """
    if user_path is None or not isinstance(user_path, str) or not user_path.strip():
        raise ValueError("filename is required")

    raw = user_path.replace("\\", "/").strip().lstrip("/")
    segments = [p for p in raw.split("/") if p not in ("", ".")]
    if not segments:
        raise ValueError("invalid filename")
    if ".." in segments:
        raise ValueError("path segments cannot include '..'")
    if len(segments) > 32:
        raise ValueError("path is too deep (max 32 segments)")

    for seg in segments[:-1]:
        if not _SAFE_SEG.match(seg):
            raise ValueError(
                f"invalid folder segment '{seg}': "
                "use only letters, digits, hyphen, underscore; max 128 chars"
            )

    base = segments[-1]
    if "/" in base or "\\" in base:
        raise ValueError("invalid filename")

    final_name = ensure_docx_extension(base)
    stem = final_name[:-5] if final_name.lower().endswith(".docx") else final_name
    if not _SAFE_BASENAME.match(stem):
        raise ValueError(
            "invalid document name: use .docx, start with alphanumeric, "
            "and only common filename characters (_ . - spaces)"
        )

    return "/".join([*segments[:-1], final_name])
