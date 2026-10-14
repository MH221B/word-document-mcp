"""manage_style and manage_list."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument, WordMcpError
from word_mcp.history import ChangeStack
from word_mcp.operations import gate_range
from word_mcp.session import require_writable

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
    doc = session.doc()
    require_writable(doc)

    if action == "list":
        return {
            "styles": [
                {
                    "name": str(session.guard(lambda s=style: s.NameLocal)),
                    "built_in": bool(session.guard(lambda s=style: s.BuiltIn)),
                    "in_use": bool(session.guard(lambda s=style: s.InUse)),
                }
                for style in session.guard(lambda: list(doc.Styles))
            ]
        }

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
        change = stack.record("format", f"apply style {name}", match.handle, format_before={"Style": before})
        session.guard(lambda: setattr(match.handle, "Style", name))
        return {"operation_id": change.op_id, "applied": name, "previous": before}

    raise InvalidArgument(
        f"Unknown action {action!r}.", allowed=["list", "create", "update", "delete", "apply"]
    )


def _apply_style_formatting(session, style, formatting: dict) -> None:
    for prop, value in formatting.items():
        if prop == "character_style":
            continue
        if hasattr(style.Font, prop):
            session.guard(lambda p=prop, v=value: setattr(style.Font, p, v))
        elif hasattr(style.ParagraphFormat, prop):
            session.guard(lambda p=prop, v=value: setattr(style.ParagraphFormat, p, v))
        else:
            raise InvalidArgument(f"Unknown style property {prop!r}.")


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
    require_writable(doc)
    match = resolve(session, spec)
    rng = match.handle
    # ListGalleries is an Application property, not a Document one.
    galleries = session.guard(lambda: session.app().ListGalleries)

    if action == "apply":
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
        session.guard(lambda: rng.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=session.guard(lambda: galleries(3).ListTemplates(1)),
            ContinuePreviousList=False,
            DefaultListBehavior=2,
        ))
        restarted = session.guard(lambda: rng.ListFormat.List.ListParagraphs(1).Range.ListFormat)
        session.guard(lambda: setattr(restarted, "ListValue", start_at))
        return {"restarted_at": start_at, "to": match.describe()}

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
