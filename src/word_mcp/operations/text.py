"""edit_text and format_text."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument, WordMcpError
from word_mcp.history import ChangeStack
from word_mcp.operations import check_format_mode, check_formatting_keys, gate_delete, gate_range, note_span


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
    tracked: bool | None = None,
) -> dict:
    check_format_mode(format_mode)
    check_formatting_keys(formatting)
    doc = session.doc()
    session.require_writable()
    match = resolve(session, spec)
    rng = match.handle

    if action == "insert":
        gate_range(len(text), confirm)
        return _insert(
            session, stack, doc, rng, text, position, format_mode, formatting, match.container,
            tracked,
        )
    before = str(session.guard(lambda: rng.Text))
    if action == "replace":
        gate_delete(before, confirm)
        return _replace(
            session, stack, doc, rng, text, before, format_mode, formatting, match.container,
            tracked,
        )
    if action == "delete":
        gate_delete(before, confirm)
        session.guard(lambda: rng.Delete())
        change = stack.record(
            "text", f"delete {len(before)} chars", rng, container=match.container, text_before=before
        )
        note_span(session, rng, match.container)
        session.note_result_range(rng, match.container)
        return _result(change, doc, tracked=tracked)
    raise InvalidArgument(f"Unknown action {action!r}.", allowed=["insert", "replace", "delete"])


def _insert(
    session, stack, doc, rng, text, position, format_mode, formatting, container, tracked=None
) -> dict:
    style_before = str(session.guard(lambda: rng.Paragraphs(1).Range.Style.NameLocal))
    if position in ("end", "start"):
        # Insert through a collapsed copy so the recorded range covers the inserted
        # text alone. Inserting through `rng` itself expands it over the original
        # text too, and reverting that would delete text the insert never touched.
        point = session.guard(lambda: rng.Duplicate)
        session.guard(lambda: point.Collapse(0 if position == "end" else 1))
        if position == "end":
            session.guard(lambda: point.InsertAfter(text))
        else:
            session.guard(lambda: point.InsertBefore(text))
        inserted = session.guard(lambda: point.Duplicate)
        note_span(session, inserted, container)
        session.note_result_range(inserted, container)
        change = stack.record(
            "text", f"insert at {position}", inserted, container=container, text_before=""
        )
    elif position == "replace":
        before = str(session.guard(lambda: rng.Text))
        session.guard(lambda: setattr(rng, "Text", text))
        inserted = session.guard(lambda: rng.Duplicate)
        change = stack.record(
            "text", "insert at replace", inserted, container=container, text_before=before
        )
        note_span(session, inserted, container)
        session.note_result_range(inserted, container)
    else:
        raise InvalidArgument(f"Unknown position {position!r}.", allowed=["start", "end", "replace"])
    _apply_mode(session, inserted, style_before, format_mode, formatting)
    return _result(change, doc, style=style_before, tracked=tracked)


def _replace(
    session, stack, doc, rng, text, before, format_mode, formatting, container, tracked=None
) -> dict:
    style_before = str(session.guard(lambda: rng.Paragraphs(1).Range.Style.NameLocal))
    session.guard(lambda: setattr(rng, "Text", text))
    fresh = session.guard(lambda: rng.Duplicate)
    change = stack.record("text", "replace", fresh, container=container, text_before=before)
    note_span(session, fresh, container)
    session.note_result_range(fresh, container)
    _apply_mode(session, fresh, style_before, format_mode, formatting)
    return _result(change, doc, style=style_before, tracked=tracked)


def _apply_mode(session, rng, style_before, format_mode, formatting) -> None:
    """inherit leaves Word's own formatting alone, which is the default."""
    if format_mode == "style_only":
        _clear_formatting(session, rng)
        session.guard(lambda: setattr(rng, "Style", style_before))
    elif format_mode == "explicit":
        if not formatting:
            raise InvalidArgument("format_mode='explicit' needs a formatting dict.")
        _clear_formatting(session, rng)
        session.guard(lambda: setattr(rng, "Style", style_before))
        apply_font(session, rng, formatting)


def _clear_formatting(session, rng) -> None:
    """Clear direct formatting.

    Range.ClearFormatting is not dispatchable on this Word build, so reset the
    font and the paragraph format, which clears the same direct formatting.
    """
    session.guard(lambda: rng.Font.Reset())
    session.guard(lambda: rng.ParagraphFormat.Reset())


def apply_font(session, rng, formatting: dict) -> None:
    from word_mcp.inspect import FONT_PROPS, PARA_PROPS

    font = set(FONT_PROPS)
    known = font | set(PARA_PROPS)
    for prop, value in formatting.items():
        if prop not in known:
            raise InvalidArgument(f"Unknown formatting property {prop!r}.")
        target = rng.Font if prop in font else rng.ParagraphFormat
        session.guard(lambda t=target, p=prop, v=value: setattr(t, p, v))


def _result(change, doc, tracked: bool | None = None, **extra) -> dict:
    import json

    payload = {
        "operation_id": change.op_id,
        "changed": change.description,
        "tracked_as_revision": bool(doc.TrackRevisions) if tracked is None else tracked,
        **extra,
    }
    return json.loads(json.dumps(payload, default=str))


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
    session.require_writable()
    match = resolve(session, spec)
    rng = match.handle
    note_span(session, rng, match.container)
    session.note_result_range(rng, match.container)
    length = int(session.guard(lambda: rng.End - rng.Start))
    gate_range(length, confirm)

    if action == "clear":
        before = _capture(session, rng)
        change = stack.record("format", "clear formatting", rng, container=match.container, format_before=before)
        _clear_formatting(session, rng)
        return _result(change, doc)

    if action == "copy_from":
        if not source:
            raise InvalidArgument("action='copy_from' needs a source target.")
        src = resolve(session, inspect_spec(source))
        before = _capture(session, rng)
        change = stack.record("format", "copy formatting", rng, container=match.container, format_before=before)
        _copy_formatting(session, rng, src.handle)
        return _result(change, doc)

    if action in ("character", "paragraph"):
        if not formatting:
            raise InvalidArgument(f"action={action!r} needs a formatting dict.")
        before = _capture(session, rng)
        change = stack.record("format", action, rng, container=match.container, format_before=before)
        apply_font(session, rng, formatting)
        return _result(change, doc)

    raise InvalidArgument(
        f"Unknown action {action!r}.", allowed=["character", "paragraph", "clear", "copy_from"]
    )


def inspect_spec(target: dict):
    from word_mcp.addressing import TargetSpec

    return TargetSpec.from_dict(target)


def _copy_formatting(session, rng, source_handle) -> None:
    """Copy formatting only. FormattedText would copy the text along with it."""
    from word_mcp.inspect import FONT_PROPS, PARA_PROPS

    source_props = _capture(session, source_handle)
    apply_font(session, rng, {k: v for k, v in source_props.items() if k in FONT_PROPS})
    for prop, value in source_props.items():
        if prop in PARA_PROPS:
            session.guard(lambda p=prop, v=value: setattr(rng.ParagraphFormat, p, v))


def _capture(session, rng) -> dict:
    from word_mcp.inspect import FONT_PROPS, PARA_PROPS, _snapshot

    return {
        **_snapshot(rng.Font, FONT_PROPS, session),
        **_snapshot(rng.ParagraphFormat, PARA_PROPS, session),
    }


def apply_edits(session, stack, edits: list[dict], resolve, confirm: bool = False) -> dict:
    """Apply several text edits in one call. Stops at the first failure.

    A failure is reported under the standard error key so a caller that keys on
    `error` cannot mistake a half-applied batch for success. Each applied entry
    carries its own result; recover a partial batch with reverse-order
    `revert_change` calls.
    """
    applied: list[dict] = []
    tracked = bool(session.guard(lambda: session.doc().TrackRevisions))
    for i, edit in enumerate(edits):
        try:
            result = edit_text(
                session, stack, inspect_spec(edit.get("target", {})), resolve,
                action=edit.get("action", "replace"),
                text=edit.get("text", ""),
                position=edit.get("position", "end"),
                format_mode=edit.get("format_mode", "inherit"),
                formatting=edit.get("formatting"),
                confirm=edit.get("confirm", confirm),
                tracked=tracked,
            )
        except WordMcpError as exc:
            return {
                "applied": applied,
                "error": {"code": "batch_stopped", "index": i, "cause": exc.to_dict()},
            }
        applied.append({"index": i, **result})
    return {"applied": applied, "error": None}
