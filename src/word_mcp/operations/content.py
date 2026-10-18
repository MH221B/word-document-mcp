"""edit_media, edit_links, edit_content_controls."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument, TargetNotFound
from word_mcp.operations import insert_caption, pick

WD_WRAP_INLINE = 0
WD_WRAP_SQUARE = 1
WD_WRAP_TIGHT = 2
WD_WRAP_THROUGH = 3
WD_WRAP_TOP_BOTTOM = 4


def find_inline_shape(session, index: int, doc):
    return pick(session, doc.InlineShapes, index, "image", "images")


def edit_media(
    session,
    resolve,
    action: str,
    spec=None,
    image_index: int = 1,
    path: str = "",
    width: float | None = None,
    height: float | None = None,
    wrap: str = "inline",
    label: str = "Figure",
    title: str = "",
    position: str = "below",
) -> dict:
    doc = session.doc()
    session.require_writable()

    if action == "insert":
        if not path:
            raise InvalidArgument("action='insert' needs a path.")
        match = resolve(session, spec)
        shape = session.guard(
            lambda: doc.InlineShapes.AddPicture(
                FileName=path, LinkToFile=False, SaveWithDocument=True, Range=match.handle
            )
        )
        _apply_size(session, shape, width, height)
        return {
            "inserted_image": True,
            "at": match.describe(),
            "width": float(session.guard(lambda: shape.Width)),
            "height": float(session.guard(lambda: shape.Height)),
        }

    shape = find_inline_shape(session, image_index, doc)

    if action == "size":
        _apply_size(session, shape, width, height)
        return {
            "image": image_index,
            "width": float(session.guard(lambda: shape.Width)),
            "height": float(session.guard(lambda: shape.Height)),
        }

    if action == "wrap":
        if wrap == "inline":
            return {"image": image_index, "wrap": "inline"}
        mapping = {
            "square": WD_WRAP_SQUARE,
            "tight": WD_WRAP_TIGHT,
            "through": WD_WRAP_THROUGH,
            "top_bottom": WD_WRAP_TOP_BOTTOM,
        }
        if wrap not in mapping:
            raise InvalidArgument(f"Unknown wrap {wrap!r}.", allowed=["inline", *mapping])
        floating = session.guard(lambda: shape.ConvertToShape())
        session.guard(lambda: setattr(floating.WrapFormat, "Type", mapping[wrap]))
        return {"image": image_index, "wrap": wrap}

    if action == "caption":
        if position not in ("above", "below"):
            raise InvalidArgument(f"Unknown position {position!r}.", allowed=["above", "below"])
        return _caption_image(session, shape, label, title, position)

    raise InvalidArgument(f"Unknown action {action!r}.", allowed=["insert", "size", "wrap", "caption"])


def _apply_size(session, shape, width, height) -> None:
    if width:
        session.guard(lambda: setattr(shape, "Width", width))
    if height:
        session.guard(lambda: setattr(shape, "Height", height))


def _caption_image(session, shape, label: str, title: str, position: str) -> dict:
    """Insert a caption paragraph next to the image's paragraph."""
    para = session.guard(lambda: shape.Range.Paragraphs(1).Range)
    return insert_caption(session, para, label, title, position)


def edit_links(session, resolve, action: str, spec=None, **kwargs) -> dict:
    doc = session.doc()

    if action == "hyperlink":
        session.require_writable()
        address = kwargs.get("address", "")
        if not address:
            raise InvalidArgument("action='hyperlink' needs an address.")
        match = resolve(session, spec)
        session.guard(
            lambda: doc.Hyperlinks.Add(
                Anchor=match.handle,
                Address=address,
                SubAddress=kwargs.get("sub_address", ""),
                ScreenTip=kwargs.get("screen_tip", ""),
                TextToDisplay=kwargs.get("display_text") or match.text,
            )
        )
        return {"hyperlink": address, "at": match.describe()}

    if action == "bookmark_add":
        session.require_writable()
        name = kwargs.get("name", "")
        if not name:
            raise InvalidArgument("action='bookmark_add' needs a name.")
        match = resolve(session, spec)
        session.guard(lambda: doc.Bookmarks.Add(Name=name, Range=match.handle))
        return {"bookmark": name, "at": match.describe()}

    if action == "bookmark_find":
        name = kwargs.get("name", "")
        exists = session.guard(lambda: bool(doc.Bookmarks.Exists(name)))
        if not exists:
            raise TargetNotFound(f"No bookmark named {name!r}.")
        rng = session.guard(lambda: doc.Bookmarks(name).Range)
        return {"bookmark": name, "text": str(session.guard(lambda: rng.Text))}

    if action in ("footnote", "endnote"):
        session.require_writable()
        text = kwargs.get("text", "")
        match = resolve(session, spec)
        if action == "footnote":
            session.guard(lambda: match.handle.Footnotes.Add(match.handle, text))
        else:
            session.guard(lambda: match.handle.Endnotes.Add(match.handle, text))
        return {action: text, "at": match.describe()}

    raise InvalidArgument(
        f"Unknown action {action!r}.",
        allowed=["hyperlink", "bookmark_add", "bookmark_find", "footnote", "endnote"],
    )


WD_CONTENT_CONTROL_RICH_TEXT = 0
WD_CONTENT_CONTROL_TEXT = 1
WD_CONTENT_CONTROL_DROPDOWN = 3
WD_CONTENT_CONTROL_DATE = 6
WD_CONTENT_CONTROL_CHECKBOX = 8

CONTROL_TYPES = {
    "rich_text": WD_CONTENT_CONTROL_RICH_TEXT,
    "text": WD_CONTENT_CONTROL_TEXT,
    "dropdown": WD_CONTENT_CONTROL_DROPDOWN,
    "date": WD_CONTENT_CONTROL_DATE,
    "checkbox": WD_CONTENT_CONTROL_CHECKBOX,
}


def edit_content_controls(
    session,
    resolve,
    action: str,
    spec=None,
    control_index: int = 1,
    control_type: str = "rich_text",
    tag: str = "",
    title: str = "",
    value: str = "",
    items: list | None = None,
) -> dict:
    doc = session.doc()

    if action == "insert":
        session.require_writable()
        if control_type not in CONTROL_TYPES:
            raise InvalidArgument(f"Unknown control type {control_type!r}.", allowed=sorted(CONTROL_TYPES))
        match = resolve(session, spec)
        control = session.guard(lambda: doc.ContentControls.Add(CONTROL_TYPES[control_type], match.handle))
        if tag:
            session.guard(lambda: setattr(control, "Tag", tag))
        if title:
            session.guard(lambda: setattr(control, "Title", title))
        if value:
            session.guard(lambda: setattr(control.Range, "Text", value))
        if items and control_type == "dropdown":
            control = session.guard(lambda: control.DropdownListEntries)
            for item in items:
                session.guard(lambda i=item: control.Add(Text=i, Value=i))
        return {"inserted": control_type, "tag": tag, "title": title, "at": match.describe()}

    if action == "list":
        controls = session.guard(lambda: list(doc.ContentControls))
        return {
            "content_controls": [
                {
                    "index": i,
                    "type": int(session.guard(lambda c=c: c.Type)),
                    "tag": str(session.guard(lambda c=c: c.Tag)),
                    "title": str(session.guard(lambda c=c: c.Title)),
                    "text": str(session.guard(lambda c=c: c.Range.Text))[:120],
                }
                for i, c in enumerate(controls, start=1)
            ]
        }

    if action == "set_value":
        session.require_writable()
        control = pick(session, doc.ContentControls, control_index, "content control", "controls")
        if session.guard(lambda: control.Type) == WD_CONTENT_CONTROL_CHECKBOX:
            session.guard(lambda: setattr(control, "Checked", value in ("true", "True", "1", "yes")))
        else:
            session.guard(lambda: setattr(control.Range, "Text", value))
        return {"content_control": control_index, "value": value}

    raise InvalidArgument(f"Unknown action {action!r}.", allowed=["insert", "set_value", "list"])
