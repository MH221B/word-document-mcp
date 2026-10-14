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


def test_an_edit_patches_the_index_instead_of_rebuilding(session):
    first = session.index()
    stack = ChangeStack()
    # The MCP tools wrap every mutation in begin_change / settle_index; do the
    # same here so the patch path runs rather than a plain fingerprint rebuild.
    session.begin_change()
    result = text_ops.edit_text(
        session, stack, inspect.TargetSpec(text="Total riders"), inspect.resolve,
        action="insert", text=" and horses", position="end",
    )
    session.settle_index()
    assert session.index() is first, "a describable edit patches the index"
    # The patch must agree with a fresh read of the document.
    outline = inspect.document_outline(session)
    assert [h["text"] for h in outline["headings"]] == ["Overview", "Risks"]
    assert session.index().fingerprint == session.guard(lambda: session.doc().Content.End)
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


def test_a_structural_edit_still_rebuilds(session):
    first = session.index()
    stack = ChangeStack()
    session.begin_change()
    structure_ops.edit_table(
        session, stack, inspect.resolve, action="insert",
        spec=inspect.TargetSpec(text="First item"), table_index=1,
        rows=1, columns=1, values=[["x"]],
    )
    session.settle_index()
    assert session.index() is not first, "a table insert cannot be described"


def test_offsets_survive_an_edit_and_a_delete(session):
    """A patched index must report the same paragraph as a fresh one would."""
    first = session.index()
    stack = ChangeStack()
    session.begin_change()
    inserted = text_ops.edit_text(
        session, stack, inspect.TargetSpec(text="Total risks"), inspect.resolve,
        action="insert", text=" (checked)", position="end",
    )
    session.settle_index()
    assert session.index() is first, "the insert patched the index"
    after_insert = inspect.resolve(session, inspect.TargetSpec(text="Total risks"))
    paragraph = str(session.guard(lambda: after_insert.handle.Paragraphs(1).Range.Text))
    assert "(checked)" in paragraph
    assert inspect.resolve(session, inspect.TargetSpec(text="Second item")).paragraph_index > 0
    document_ops.revert_change(session, stack, inserted["operation_id"])


def test_find_text_limit_bounds_the_work_and_the_results(session):
    every = inspect.find_all(session, inspect.TargetSpec(text="Total"))
    assert len(every) >= 2
    first = inspect.find_all(session, inspect.TargetSpec(text="Total"), stop_after=1)
    assert len(first) == 1
    assert first[0].text == every[0].text


def test_inventory_is_cached_until_the_document_changes(session):
    assert session.inventory() is None, "not read yet"
    first = inspect.inspect_document(session)
    assert session.inventory() is first
    assert inspect.inspect_document(session) is first
    assert inspect.inspect_document(session, refresh=True) is not first


def test_caption_above_the_anchor_is_found(session):
    stack = ChangeStack()
    structure_ops.edit_table(
        session, stack, inspect.resolve, action="insert", spec=inspect.TargetSpec(text="Second item"),
        rows=2, columns=2, values=[["a", "b"], ["c", "d"]],
    )
    caption = structure_ops.edit_table(
        session, stack, inspect.resolve, action="caption", table_index=1,
        title="Above caption", position="above",
    )
    assert "Above caption" in caption["text"], "the caption above the table is located"


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
