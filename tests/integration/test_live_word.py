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


def test_every_instance_is_searched(session):
    """A machine runs several WINWORD processes, and the one holding this document
    used to be hidden behind GetActiveObject returning a windowless instance. The
    sample is open in whichever instance, so finding it proves the search covers
    all of them."""
    assert session.applications(), "at least one Word instance"
    assert any(d.path.endswith("plain.docx") for d in session.documents())


def test_read_and_find(session):
    found = inspect.find_all(session, inspect.TargetSpec(text="Total"))
    assert len(found) >= 2, "the sample has two Total paragraphs"


def test_outline_reports_real_ancestors(session):
    outline = inspect.document_outline(session)
    got = [
        (h["text"], h["level"], h["paragraph_index"], h["heading_path"])
        for h in outline["headings"]
    ]
    # Both headings are level 1, so neither has an ancestor. The old walk
    # reported "Overview" as the parent of "Risks".
    assert got == [("Overview", 1, 0, ""), ("Risks", 1, 3, "")]
    assert len(outline["images"]) == 1
    assert len(outline["sections"]) == 1


def test_index_is_reused_until_the_document_changes(session):
    first = session.index()
    assert session.index() is first, "an unchanged document reuses the index"
    stack = ChangeStack()
    result = text_ops.edit_text(
        session, stack, inspect.TargetSpec(text="Total riders"), inspect.resolve,
        action="insert", text=" and horses", position="end",
    )
    assert session.index() is not first, "an edit forces a rebuild"
    document_ops.revert_change(session, stack, result["operation_id"])


def paragraph_text(session, spec):
    """The whole paragraph holding a match, not just the matched text."""
    match = inspect.resolve(session, spec)
    return str(session.guard(lambda: match.handle.Paragraphs(1).Range.Text))


def test_insert_then_revert(session):
    spec = inspect.TargetSpec(text="Total riders")
    before = paragraph_text(session, spec)
    stack = ChangeStack()
    result = text_ops.edit_text(
        session, stack, inspect.TargetSpec(text="Total riders"), inspect.resolve,
        action="insert", text=" (draft)", position="end",
    )
    assert result["operation_id"] == 1
    document_ops.revert_change(session, stack, result["operation_id"])
    # The whole paragraph, so a revert that eats the surrounding text fails here.
    assert paragraph_text(session, inspect.TargetSpec(text="Total riders")) == before


def test_footer_edit_stays_in_its_story(session):
    """Header and footer offsets are story-relative, so an edit there must land in
    that story and a revert there must not damage the body."""
    body_spec = inspect.TargetSpec(text="Overview")
    body_before = paragraph_text(session, body_spec)
    structure_ops.edit_page_layout(
        session, inspect.resolve, action="header_footer",
        side="footer", which="primary", text="Stage footer",
    )
    footer = lambda: str(session.guard(lambda: session.doc().Sections(1).Footers(1).Range.Text))
    stack = ChangeStack()
    result = text_ops.edit_text(
        session, stack, inspect.TargetSpec(text="Stage footer", container="footer"),
        inspect.resolve, action="insert", text=" draft", position="end",
    )
    assert "Stage footer draft" in footer()
    assert paragraph_text(session, body_spec) == body_before
    document_ops.revert_change(session, stack, result["operation_id"])
    assert footer().strip() == "Stage footer"
    assert paragraph_text(session, body_spec) == body_before


def test_footer_matches_keep_their_own_story(session):
    count = session.guard(lambda: session.doc().Sections(1).Footers(1).Range.Paragraphs.Count)
    matches = inspect.find_all(
        session, inspect.TargetSpec(text="Stage footer", container="footer")
    )
    assert matches, "the footer written by the earlier test still holds text"
    assert 0 <= matches[0].paragraph_index < count
    assert matches[0].container == "footer"


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
