import pytest

from word_document_server.storage_paths import (
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
