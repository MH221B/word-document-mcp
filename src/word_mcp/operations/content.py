"""edit_media, edit_links, edit_content_controls."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument, TargetNotFound
from word_mcp.session import require_writable

WD_WRAP_INLINE = 0
WD_WRAP_SQUARE = 1
WD_WRAP_TIGHT = 2
WD_WRAP_THROUGH = 3
WD_WRAP_TOP_BOTTOM = 4


def find_inline_shape(session, index: int, doc):
    shapes = session.guard(lambda: list(doc.InlineShapes))
    if index < 1 or index > len(shapes):
        raise TargetNotFound(f"There is no image {index}.", images=len(shapes))
    return shapes[index - 1]


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
    require_writable(doc)

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
        floating = session.guard(lambda: shape.ConvertToShape())
        mapping = {
            "square": WD_WRAP_SQUARE,
            "tight": WD_WRAP_TIGHT,
            "through": WD_WRAP_THROUGH,
            "top_bottom": WD_WRAP_TOP_BOTTOM,
        }
        if wrap not in mapping:
            raise InvalidArgument(f"Unknown wrap {wrap!r}.", allowed=["inline", *mapping])
        session.guard(lambda: setattr(floating.WrapFormat, "Type", mapping[wrap]))
        return {"image": image_index, "wrap": wrap}

    if action == "caption":
        return _caption_image(session, shape, label, title, position)

    raise InvalidArgument(f"Unknown action {action!r}.", allowed=["insert", "size", "wrap", "caption"])


def _apply_size(session, shape, width, height) -> None:
    if width:
        session.guard(lambda: setattr(shape, "Width", width))
    if height:
        session.guard(lambda: setattr(shape, "Height", height))


def _caption_image(session, shape, label: str, title: str, position: str) -> dict:
    """Insert a caption paragraph next to the image's paragraph."""
    doc = session.doc()
    para = session.guard(lambda: shape.Range.Paragraphs(1).Range)
    # Caption the image's own paragraph. InsertCaption adds the caption as its own
    # paragraph; find that paragraph so it can be styled and its text reported.
    start = int(session.guard(lambda: para.End))
    session.guard(
        lambda: para.InsertCaption(
            Label=label, Title=title, Position=-1 if position == "above" else 1
        )
    )
    caption_para = None
    for candidate in session.guard(lambda: list(doc.Paragraphs)):
        if int(session.guard(lambda c=candidate: c.Range.Start)) == start:
            caption_para = candidate.Range
            break
    if caption_para is not None:
        session.guard(lambda: setattr(caption_para, "Style", "Caption"))
    text = (
        str(session.guard(lambda: caption_para.Text)).strip()
        if caption_para is not None
        else ""
    )
    return {
        "caption_added": True,
        "label": label,
        "title": title,
        "position": position,
        "text": text,
    }


def edit_links(session, resolve, action: str, spec=None, **kwargs) -> dict:
    doc = session.doc()
    require_writable(doc)

    if action == "hyperlink":
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
