"""Shared validation for the operation modules."""

from word_mcp.errors import ConfirmationRequired, InvalidArgument, TargetNotFound

DELETE_PARA_THRESHOLD = 1
DELETE_CHAR_THRESHOLD = 500
RANGE_CHAR_THRESHOLD = 2000

FORMAT_MODES = ("inherit", "style_only", "explicit")


def check_format_mode(mode: str) -> str:
    if mode not in FORMAT_MODES:
        raise InvalidArgument(f"Unknown format_mode {mode!r}.", allowed=list(FORMAT_MODES))
    return mode


def check_formatting_keys(formatting: dict | None) -> None:
    """Reject unknown formatting keys before the document is touched."""
    if not formatting:
        return
    from word_mcp.inspect import FONT_PROPS, PARA_PROPS

    known = set(FONT_PROPS) | set(PARA_PROPS)
    unknown = sorted(k for k in formatting if k not in known)
    if unknown:
        raise InvalidArgument(f"Unknown formatting property(ies) {unknown}.", allowed=sorted(known))


def gate_delete(text: str, confirm: bool) -> None:
    """Deletions big enough to hurt need an explicit confirm."""
    paragraphs = text.count("\r")
    if not confirm and (paragraphs > DELETE_PARA_THRESHOLD or len(text) > DELETE_CHAR_THRESHOLD):
        raise ConfirmationRequired(
            "This deletes a large amount of content. Pass confirm=true to proceed.",
            paragraphs=paragraphs,
            characters=len(text),
            paragraphs_threshold=DELETE_PARA_THRESHOLD,
            characters_threshold=DELETE_CHAR_THRESHOLD,
        )


def gate_range(length: int, confirm: bool) -> None:
    if not confirm and length > RANGE_CHAR_THRESHOLD:
        raise ConfirmationRequired(
            "This touches a large range. Pass confirm=true to proceed.",
            characters=length,
            characters_threshold=RANGE_CHAR_THRESHOLD,
        )


def require_confirm(confirm: bool, what: str) -> None:
    if not confirm:
        raise ConfirmationRequired(f"{what} requires confirm=true.")


def caption_paragraph(session, doc, start: int, end: int, position: str):
    """The paragraph InsertCaption just added, from the anchor's own boundary.

    Word inserts the caption as its own paragraph directly before or after the
    anchor, so the caption starts at the anchor's old boundary either way: below,
    at the offset the anchor ended; above, at the offset the anchor began.
    Scanning every paragraph for the offset only worked for `below`.
    """
    point = max(end, 0) if position == "below" else max(start, 0)
    return session.guard(lambda: doc.Range(point, point).Paragraphs(1).Range)


def insert_caption(session, host, label: str, title: str, position: str) -> dict:
    """Insert a caption paragraph next to `host` and report its text."""
    doc = session.doc()
    start = int(session.guard(lambda: host.Start))
    end = int(session.guard(lambda: host.End))
    session.guard(
        lambda: host.InsertCaption(
            Label=label, Title=title, Position=0 if position == "above" else 1
        )
    )
    caption_para = caption_paragraph(session, doc, start, end, position)
    session.guard(lambda: setattr(caption_para, "Style", "Caption"))
    return {
        "caption_added": True,
        "label": label,
        "title": title,
        "position": position,
        "text": str(session.guard(lambda: caption_para.Text)).strip(),
    }


def pick(session, collection, index: int, what: str, key: str):
    """The `index`-th item of a collection, or TargetNotFound."""
    items = session.guard(lambda: list(collection))
    if index < 1 or index > len(items):
        raise TargetNotFound(f"There is no {what} {index}.", **{key: len(items)})
    return items[index - 1]


def note_span(session, rng) -> None:
    """Tell the session which span this operation changed.

    `rng` is the range the change produced, in post-edit offsets: the inserted
    text for an insert, the new text for a replace, the collapsed point for a
    delete. A tool that does not call this drops the index, as it always did.
    """
    session.note_change(
        int(session.guard(lambda: rng.Start)), int(session.guard(lambda: rng.End))
    )
