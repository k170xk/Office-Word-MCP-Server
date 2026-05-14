import pytest

from word_document_server.storage_paths import (
    apply_workspace_document_prefix,
    normalize_storage_document_key,
    validate_workspace_segment,
)


def test_workspace_segment_ok():
    assert validate_workspace_segment("user_abc-1") == "user_abc-1"


def test_workspace_segment_blank():
    assert validate_workspace_segment(None) is None
    assert validate_workspace_segment("") is None


def test_normalize_plain():
    assert normalize_storage_document_key("report") == "report.docx"
    assert normalize_storage_document_key("report.docx") == "report.docx"


def test_normalize_nested():
    assert normalize_storage_document_key("tenant_1/note") == "tenant_1/note.docx"


def test_workspace_bad():
    with pytest.raises(ValueError):
        validate_workspace_segment("evil/evil")


def test_normalize_rejects_traversal():
    with pytest.raises(ValueError):
        normalize_storage_document_key("../x.docx")


def test_prefix_single_segment_only():
    assert apply_workspace_document_prefix("u1", "a.docx") == "u1/a.docx"


def test_prefix_multisegment_unchanged():
    assert apply_workspace_document_prefix("u1", "other/doc.docx") == "other/doc.docx"


def test_prefix_workspace_none_returns_raw():
    assert apply_workspace_document_prefix(None, "a.docx") == "a.docx"


def test_prefix_invalid_workspace_raises():
    with pytest.raises(ValueError):
        apply_workspace_document_prefix("no spaces!", "x.docx")
