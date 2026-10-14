"""Opt-in tests. Requires Word open with samples/plain.docx.

    uv run pytest -m integration

Tests run in file order and edit the open document. Regenerate a clean sample
first: close it in Word, run `uv run python samples/make_samples.py`, reopen it.
The last test saves the document to a temporary copy and puts it back.
"""

import pytest

from word_mcp import inspect
from word_mcp.errors import WordMcpError
from word_mcp.history import ChangeStack
from word_mcp.operations import document as document_ops
from word_mcp.operations import structure as structure_ops
from word_mcp.operations import text as text_ops
from word_mcp.session import ComWordSession

pytestmark = pytest.mark.integration


@pytest.fixture(scope="module")
def session():
    s = ComWordSession()
    try:
        s.pin("plain.docx")
    except WordMcpError as exc:
        pytest.skip(f"Word or samples/plain.docx is not available: {exc}")
    yield s


def test_attach_and_list(session):
    names = [d.name for d in session.documents()]
    assert "plain.docx" in names


def test_read_and_find(session):
    found = inspect.find_all(session, inspect.TargetSpec(text="Total"))
    assert len(found) >= 2, "the sample has two Total paragraphs"


def test_insert_then_revert(session):
    stack = ChangeStack()
    result = text_ops.edit_text(
        session, stack, inspect.TargetSpec(text="Total riders"), inspect.resolve,
        action="insert", text=" (draft)", position="end",
    )
    assert result["operation_id"] == 1
    document_ops.revert_change(session, stack, result["operation_id"])
    assert "draft" not in inspect.read_range(session, inspect.TargetSpec(text="Total riders"))["text"]


def test_style_only_insert_has_no_direct_formatting(session):
    stack = ChangeStack()
    text_ops.edit_text(
        session, stack, inspect.TargetSpec(text="Total riders"), inspect.resolve,
        action="insert", text=" extra", position="end", format_mode="style_only",
    )
    fmt = inspect.get_formatting(session, inspect.TargetSpec(text="Total riders"))
    assert fmt["deviations_from_style"] == {}


def test_caption_lands_outside_the_table(session):
    stack = ChangeStack()
    structure_ops.edit_table(
        session, stack, inspect.resolve, action="insert",
        spec=inspect.TargetSpec(text="Totals"), rows=2, columns=2,
        values=[["a", "b"], ["c", "d"]],
    )
    caption = structure_ops.edit_table(
        session, stack, inspect.resolve, action="caption", table_index=1,
        title="Integration caption",
    )
    assert "Integration caption" in caption["text"]
    assert "Integration caption" not in session.doc().Tables(1).Range.Text


def test_saved_copy_carries_the_text(tmp_path, session):
    """Read the saved file with python-docx, so the check does not depend on our
    own reader. Must stay last: it saves the document to a copy, then puts it
    back where it was."""
    from docx import Document

    doc = session.doc()
    original = session.pinned().path
    path = tmp_path / "check.docx"
    session.guard(lambda: doc.SaveAs2(str(path)))
    try:
        body = "\n".join(p.text for p in Document(str(path)).paragraphs)
        assert "Total riders" in body
        assert "Integration caption" in body
    finally:
        session.guard(lambda: doc.SaveAs2(original))
        session.unpin()
