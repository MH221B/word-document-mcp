"""edit_text and format_text."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument
from word_mcp.history import ChangeStack
from word_mcp.operations import check_format_mode, gate_delete, gate_range
from word_mcp.session import require_writable


def edit_text(
    session,
    stack: ChangeStack,
    spec,
    resolve,
    action: str,
    text: str = "",
    position: str = "end",
    format_mode: str = "inherit",
    formatting: dict | None = None,
    confirm: bool = False,
) -> dict:
    check_format_mode(format_mode)
    doc = session.doc()
    require_writable(doc)
    match = resolve(session, spec)
    rng = match.handle
    before = str(session.guard(lambda: rng.Text))

    if action == "insert":
        gate_range(len(text), confirm)
        return _insert(session, stack, doc, rng, text, position, format_mode, formatting)
    if action == "replace":
        gate_delete(before, confirm)
        return _replace(session, stack, doc, rng, text, format_mode, formatting)
    if action == "delete":
        gate_delete(before, confirm)
        change = stack.record("text", f"delete {len(before)} chars", rng, text_before=before)
        session.guard(lambda: rng.Delete())
        return _result(change, doc)
    raise InvalidArgument(f"Unknown action {action!r}.", allowed=["insert", "replace", "delete"])


def _insert(session, stack, doc, rng, text, position, format_mode, formatting) -> dict:
    style_before = str(session.guard(lambda: rng.Paragraphs(1).Range.Style.NameLocal))
    start = int(session.guard(lambda: rng.Start))
    change = stack.record("text", f"insert at {position}", rng, text_before="")
    if position == "end":
        session.guard(lambda: rng.InsertAfter(text))
    elif position == "start":
        session.guard(lambda: rng.InsertBefore(text))
    elif position == "replace":
        session.guard(lambda: setattr(rng, "Text", text))
    else:
        raise InvalidArgument(f"Unknown position {position!r}.", allowed=["start", "end", "replace"])
    fresh = session.guard(lambda: doc.Range(start, int(rng.End)))
    _apply_mode(session, fresh, style_before, format_mode, formatting)
    return _result(change, doc, style=style_before)


def _replace(session, stack, doc, rng, text, format_mode, formatting) -> dict:
    style_before = str(session.guard(lambda: rng.Paragraphs(1).Range.Style.NameLocal))
    before = str(session.guard(lambda: rng.Text))
    start = int(session.guard(lambda: rng.Start))
    change = stack.record("text", "replace", rng, text_before=before)
    session.guard(lambda: setattr(rng, "Text", text))
    fresh = session.guard(lambda: doc.Range(start, int(rng.End)))
    _apply_mode(session, fresh, style_before, format_mode, formatting)
    return _result(change, doc, style=style_before)


def _apply_mode(session, rng, style_before, format_mode, formatting) -> None:
    """inherit leaves Word's own formatting alone, which is the default."""
    if format_mode == "style_only":
        session.guard(lambda: rng.ClearFormatting())
        session.guard(lambda: setattr(rng, "Style", style_before))
    elif format_mode == "explicit":
        if not formatting:
            raise InvalidArgument("format_mode='explicit' needs a formatting dict.")
        session.guard(lambda: rng.ClearFormatting())
        session.guard(lambda: setattr(rng, "Style", style_before))
        apply_font(session, rng, formatting)


def apply_font(session, rng, formatting: dict) -> None:
    for prop, value in formatting.items():
        if not hasattr(rng.Font, prop) and not hasattr(rng.ParagraphFormat, prop):
            raise InvalidArgument(f"Unknown formatting property {prop!r}.")
        target = rng.Font if hasattr(rng.Font, prop) else rng.ParagraphFormat
        session.guard(lambda t=target, p=prop, v=value: setattr(t, p, v))


def _result(change, doc, **extra) -> dict:
    import json

    payload = {
        "operation_id": change.op_id,
        "changed": change.description,
        "tracked_as_revision": bool(doc.TrackRevisions),
        **extra,
    }
    return json.loads(json.dumps(payload, default=str))


WD_UNDEFINED = 9999999


def format_text(
    session,
    stack: ChangeStack,
    spec,
    resolve,
    action: str,
    formatting: dict | None = None,
    source: dict | None = None,
    confirm: bool = False,
) -> dict:
    doc = session.doc()
    require_writable(doc)
    match = resolve(session, spec)
    rng = match.handle
    length = int(session.guard(lambda: rng.End - rng.Start))
    gate_range(length, confirm)

    if action == "clear":
        before = _capture(session, rng)
        change = stack.record("format", "clear formatting", rng, format_before=before)
        session.guard(lambda: rng.ClearFormatting())
        return _result(change, doc)

    if action == "copy_from":
        if not source:
            raise InvalidArgument("action='copy_from' needs a source target.")
        src = resolve(session, inspect_spec(source))
        before = _capture(session, rng)
        change = stack.record("format", "copy formatting", rng, format_before=before)
        _copy_formatting(session, rng, src.handle)
        return _result(change, doc)

    if action in ("character", "paragraph"):
        if not formatting:
            raise InvalidArgument(f"action={action!r} needs a formatting dict.")
        before = _capture(session, rng)
        change = stack.record("format", action, rng, format_before=before)
        apply_font(session, rng, formatting)
        return _result(change, doc)

    raise InvalidArgument(
        f"Unknown action {action!r}.", allowed=["character", "paragraph", "clear", "copy_from"]
    )


def inspect_spec(target: dict):
    from word_mcp.addressing import TargetSpec

    return TargetSpec(**target)


def _copy_formatting(session, rng, source_handle) -> None:
    """Copy formatting only. FormattedText would copy the text along with it."""
    from word_mcp.inspect import FONT_PROPS, PARA_PROPS

    source_props = _capture(session, source_handle)
    apply_font(session, rng, {k: v for k, v in source_props.items() if k in FONT_PROPS})
    for prop, value in source_props.items():
        if prop in PARA_PROPS:
            session.guard(lambda p=prop, v=value: setattr(rng.ParagraphFormat, p, v))


def _capture(session, rng) -> dict:
    from word_mcp.inspect import FONT_PROPS, PARA_PROPS

    captured = {}
    for prop in FONT_PROPS + PARA_PROPS:
        target = rng.Font if prop in FONT_PROPS else rng.ParagraphFormat
        try:
            value = session.guard(lambda t=target, p=prop: getattr(t, p))
        except Exception:
            continue
        if value == WD_UNDEFINED:
            continue
        captured[prop] = value
    return captured
