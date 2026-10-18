"""manage_style and manage_list."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument, WordMcpError
from word_mcp.history import ChangeStack
from word_mcp.operations import gate_range, note_span

WD_STYLE_PARAGRAPH = 1
WD_STYLE_CHARACTER = 2


def manage_style(
    session,
    stack: ChangeStack,
    resolve,
    action: str,
    spec=None,
    name: str = "",
    based_on: str = "",
    formatting: dict | None = None,
    confirm: bool = False,
) -> dict:
    if action == "list":
        # Reuse the style inventory inspect_document already builds and caches.
        from word_mcp import inspect

        return {
            "styles": [
                {"name": s["name"], "built_in": s["built_in"], "in_use": s["in_use"]}
                for s in inspect.inspect_document(session)["styles"]
            ]
        }

    session.require_writable()
    doc = session.doc()

    if action == "create":
        if not name:
            raise InvalidArgument("action='create' needs a name.")
        style_type = WD_STYLE_CHARACTER if (formatting or {}).get("character_style") else WD_STYLE_PARAGRAPH
        style = session.guard(lambda: doc.Styles.Add(name, style_type))
        _apply_style_formatting(session, style, formatting or {})
        return {"created": name, "type": "character" if style_type == WD_STYLE_CHARACTER else "paragraph"}

    if action == "update":
        style = session.guard(lambda: doc.Styles(name))
        _apply_style_formatting(session, style, formatting or {})
        if based_on:
            session.guard(lambda: setattr(style, "BaseStyle", based_on))
        return {"updated": name, "based_on": based_on or None}

    if action == "delete":
        session.guard(lambda: doc.Styles(name).Delete())
        return {"deleted": name}

    if action == "apply":
        match = resolve(session, spec)
        length = int(session.guard(lambda: match.handle.End - match.handle.Start))
        gate_range(length, confirm)
        before = str(session.guard(lambda: match.handle.Style.NameLocal))
        change = stack.record("format", f"apply style {name}", match.handle, container=match.container, format_before={"Style": before})
        note_span(session, match.handle, match.container)
        session.guard(lambda: setattr(match.handle, "Style", name))
        return {"operation_id": change.op_id, "applied": name, "previous": before}

    raise InvalidArgument(
        f"Unknown action {action!r}.", allowed=["list", "create", "update", "delete", "apply"]
    )


def _apply_style_formatting(session, style, formatting: dict) -> None:
    from word_mcp.inspect import FONT_PROPS, PARA_PROPS

    font = set(FONT_PROPS)
    known = font | set(PARA_PROPS)
    for prop, value in formatting.items():
        if prop == "character_style":
            continue
        if prop not in known:
            raise InvalidArgument(f"Unknown style property {prop!r}.")
        target = style.Font if prop in font else style.ParagraphFormat
        session.guard(lambda t=target, p=prop, v=value: setattr(t, p, v))


def manage_list(
    session,
    stack: ChangeStack,
    resolve,
    action: str,
    spec=None,
    list_format: str = "bullet",
    level: int = 1,
    start_at: int = 1,
    confirm: bool = False,
) -> dict:
    doc = session.doc()
    session.require_writable()
    match = resolve(session, spec)
    rng = match.handle
    note_span(session, rng, match.container)
    # ListGalleries is an Application property, not a Document one.
    galleries = session.guard(lambda: session.app().ListGalleries)

    if action == "apply":
        if list_format not in ("bullet", "number"):
            raise InvalidArgument(
                f"Unknown list_format {list_format!r}.", allowed=["bullet", "number"]
            )
        template = session.guard(lambda: galleries(2).ListTemplates(1))  # wdBulletGallery
        if list_format == "number":
            template = session.guard(lambda: galleries(3).ListTemplates(1))  # wdNumberGallery
        session.guard(lambda: rng.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=template, ContinuePreviousList=False, ApplyTo=0, DefaultListBehavior=2
        ))
        return {"applied": list_format, "to": match.describe()}

    if action == "continue":
        session.guard(lambda: rng.ListFormat.ApplyListTemplate(
            ListTemplate=session.guard(lambda: galleries(1).ListTemplates(1))
        ))
        return {"continued": True, "to": match.describe()}

    if action == "restart":
        # ListValue is read-only, so the start number comes from the template's
        # level before it is applied, and the value Word actually used is read back.
        template = session.guard(lambda: galleries(3).ListTemplates(1))  # wdNumberGallery
        levels = session.guard(lambda: template.ListLevels)
        level_count = int(session.guard(lambda: levels.Count))
        if level < 1 or level > level_count:
            raise InvalidArgument(f"Level {level} is out of range.", levels=level_count)
        session.guard(lambda: setattr(levels(level), "StartAt", start_at))
        session.guard(lambda: rng.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=template,
            ContinuePreviousList=False,
            ApplyTo=0,
            DefaultListBehavior=2,
        ))
        actual = int(session.guard(lambda: rng.ListFormat.ListValue))
        return {
            "restarted_at": actual,
            "requested": start_at,
            "level": level,
            "to": match.describe(),
            "note": "StartAt is set on the shared number gallery template, which persists for the session.",
        }

    if action == "set_level":
        # A style-only bullet (for example the built-in List Bullet style) carries no
        # concrete list template, so Word refuses a direct ListLevelNumber change.
        # Attach a gallery template with the level, then set the level.
        try:
            session.guard(lambda: setattr(rng.ListFormat, "ListLevelNumber", level))
        except WordMcpError:
            template = session.guard(lambda: galleries(2).ListTemplates(1))  # wdBulletGallery
            session.guard(lambda: rng.ListFormat.ApplyListTemplateWithLevel(
                ListTemplate=template, ContinuePreviousList=False, ApplyTo=0, DefaultListBehavior=2
            ))
            session.guard(lambda: setattr(rng.ListFormat, "ListLevelNumber", level))
        return {"level": level, "to": match.describe()}

    raise InvalidArgument(
        f"Unknown action {action!r}.", allowed=["apply", "continue", "restart", "set_level"]
    )
