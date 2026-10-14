import pytest

from word_mcp import operations
from word_mcp.errors import ConfirmationRequired, InvalidArgument


def test_format_mode_is_validated():
    assert operations.check_format_mode("style_only") == "style_only"
    with pytest.raises(InvalidArgument):
        operations.check_format_mode("guess")


def test_small_delete_needs_no_confirm():
    operations.gate_delete("short text", confirm=False)


def test_large_delete_needs_confirm():
    with pytest.raises(ConfirmationRequired) as exc:
        operations.gate_delete("x" * 600, confirm=False)
    assert exc.value.details["characters_threshold"] == 500


def test_multi_paragraph_delete_needs_confirm():
    with pytest.raises(ConfirmationRequired):
        operations.gate_delete("one\r\rtwo", confirm=False)


def test_confirm_allows_large_delete():
    operations.gate_delete("x" * 600, confirm=True)


def test_large_range_needs_confirm():
    with pytest.raises(ConfirmationRequired):
        operations.gate_range(2500, confirm=False)


def test_require_confirm_passes_when_confirmed():
    operations.require_confirm(True, "accept all revisions")
