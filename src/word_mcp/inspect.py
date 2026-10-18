"""Read-only tools: outline, read, find, formatting, document state, OOXML."""

from __future__ import annotations

from typing import Any

import win32com.client

from word_mcp.addressing import (
    MAX_CANDIDATES,
    Match,
    TargetSpec,
    TokenStore,
    choose,
    normalise,
)
from word_mcp.comfast import STYLE, raw_items
from word_mcp.errors import InvalidArgument
from word_mcp.session import story_range

TOKENS = TokenStore()


def heading_path_of(paragraphs: list, paragraph_index: int) -> str:
    """Walk back to the nearest enclosing headings, nearest last.

    Ancestors are headings with strictly smaller outline levels. A previous
    sibling at the same level is not an ancestor, so the walk starts from the
    described paragraph's own level rather than from body text.
    """
    if paragraph_index <= 0:
        return ""
    cap = int(paragraphs[paragraph_index].OutlineLevel)
    parts: list[str] = []
    for i in range(paragraph_index - 1, -1, -1):
        para = paragraphs[i]
        level = int(para.OutlineLevel)
        if level < cap:
            cap = level
            parts.insert(0, para.Range.Text.strip())
            if level == 1:
                break
    return " > ".join(parts)


def paragraph_context(session, rng: Any, paragraphs: list) -> tuple[int, str, str]:
    """Paragraph index, heading path, and style for a range, within its own story."""
    for i, para in enumerate(paragraphs):
        para_range = session.guard(lambda p=para: p.Range)
        if session.guard(lambda r=para_range: r.Start <= rng.Start < r.End):
            style = str(session.guard(lambda p=para: p.Range.Style.NameLocal))
            return i, heading_path_of(paragraphs, i), style
    return -1, "", ""


def match_from_range(
    session,
    handle: Any,
    index: Any,
    paragraphs: list,
    container: str = "body",
    need_style: bool = True,
) -> Match:
    """Wrap a Word range as a Match and register an anchor token for it.

    Body context comes from the structure index. A header or footer has its own
    offset space, so it is scanned within its own story instead.
    """
    if index is not None:
        offset = session.guard(lambda: handle.Start)
        paragraph_index = index.paragraph_index(offset)
        path = index.heading_path(offset)
    else:
        paragraph_index, path, _ = paragraph_context(session, handle, paragraphs)
    text = str(session.guard(lambda: handle.Text)).replace("\r", "").replace("\x07", "")
    style = str(session.guard(lambda: handle.Style.NameLocal)) if need_style else ""
    match = Match(
        token="",
        text=text,
        paragraph_index=paragraph_index,
        heading_path=path,
        style=style,
        container=container,
        snippet=text[:80],
        handle=handle,
    )
    TOKENS.add(match)
    return match


def anchor(session, handle, container: str = "body") -> Match:
    """Register a token for a range and describe where it is.

    Body context comes from the settled structure index; a header or footer has
    its own offset space, so it is scanned within its own story.
    """
    if container == "body":
        index = session.index()
        paragraphs: list = []
    else:
        index = None
        host = story_range(session.doc(), container)
        paragraphs = session.guard(lambda: list(host.Paragraphs))
    return match_from_range(session, handle, index, paragraphs, container)


def _heading_match(index, heading, handle, token: str = "") -> Match:
    """A heading's result, described from live index data."""
    return Match(
        token=token,
        text=heading.text,
        paragraph_index=index.paragraph_index(heading.start),
        heading_path=heading.path,
        style=heading.style,
        container="body",
        snippet=heading.text[:80],
        handle=handle,
    )


def _mint_anchors(session, doc, index) -> list[tuple[str, Any]]:
    """One token and one live range per heading, minted once per index."""
    out: list[tuple[str, Any]] = []
    for heading in index.headings:
        handle = session.guard(lambda h=heading: doc.Range(h.start, h.end))
        out.append((TOKENS.add(_heading_match(index, heading, handle)), handle))
    return out


def document_outline(session, depth: int = 9, refresh: bool = False) -> dict:
    """Headings, tables, sections, and images, each with an anchor token.

    Structure comes from the session's cached index, so a second call on an
    unchanged document costs almost nothing. `refresh` forces a rebuild, for a
    structural change made by hand in Word.
    """
    if refresh:
        session.invalidate_index()
    doc = session.doc()
    index = session.index()
    kept = [n for n, h in enumerate(index.headings) if h.level <= depth]
    if (
        index.anchors is None
        or index.anchor_depth != depth
        or len(index.anchors) != len(kept)
    ):
        full = _mint_anchors(session, doc, index)
        index.anchors = [full[n] for n in kept]
        index.anchor_depth = depth
    headings = []
    for pos, n in enumerate(kept):
        heading = index.headings[n]
        token, handle = index.anchors[pos]
        headings.append(
            {"level": heading.level, **_heading_match(index, heading, handle, token).describe()}
        )
    payload = {
        "document": session.pinned().__dict__,
        "headings": headings,
        "tables": index.tables,
        "sections": index.sections,
        "images": index.images,
        "structure_source": index.source,
    }
    if index.source == "com" and session.macro_denied():
        payload["structure_hint"] = (
            "Word's VBA path is unavailable; enable 'Trust access to the VBA "
            "project object model' for a faster scan."
        )
    return payload


WD_UNDEFINED = 9999999

FONT_PROPS = [
    "Name", "Size", "Bold", "Italic", "Underline", "Color", "HighlightColorIndex",
    "Superscript", "Subscript", "AllCaps", "SmallCaps", "Spacing", "Scaling",
]

PARA_PROPS = [
    "Alignment", "LeftIndent", "RightIndent", "FirstLineIndent", "SpaceBefore",
    "SpaceAfter", "LineSpacing", "LineSpacingRule", "KeepWithNext", "KeepTogether",
    "WidowControl", "PageBreakBefore",
]


def _snapshot(obj: Any, props: list[str], session) -> dict:
    out = {}
    for prop in props:
        try:
            value = session.guard(lambda p=prop: getattr(obj, p))
        except Exception:
            continue  # a property Word does not expose here is simply absent
        if value == WD_UNDEFINED:
            continue
        out[prop] = value
    return out


def deviations(session, font: dict, paragraph: dict, style) -> dict:
    """Properties where direct formatting differs from the style definition.

    Only properties present in the direct formatting are compared, so only those
    are read from the style.
    """
    inherited = _snapshot(style.Font, [p for p in FONT_PROPS if p in font], session)
    para_style = _snapshot(
        style.ParagraphFormat, [p for p in PARA_PROPS if p in paragraph], session
    )
    return {
        prop: {"direct": value, "style": inherited.get(prop)}
        for prop, value in {**font, **paragraph}.items()
        if prop in {**inherited, **para_style} and value != {**inherited, **para_style}[prop]
    }


def read_range(session, spec, include_formatting: bool = False) -> dict:
    match = resolve(session, spec)
    payload = {
        "target": match.describe(),
        "text": session.guard(lambda: match.handle.Text),
    }
    if include_formatting:
        payload["formatting"] = _formatting(session, match)
    return payload


def get_formatting(session, spec) -> dict:
    return _formatting(session, resolve(session, spec))


def _formatting(session, match) -> dict:
    handle = match.handle
    paragraph = session.guard(lambda: handle.Paragraphs(1))
    style = session.guard(lambda: handle.Style)
    font = _snapshot(handle.Font, FONT_PROPS, session)
    para = _snapshot(handle.ParagraphFormat, PARA_PROPS, session)
    return {
        "target": match.describe(),
        "paragraph_style": str(session.guard(lambda: paragraph.Style.NameLocal)),
        "character_style": str(session.guard(lambda: style.NameLocal)),
        "surrounding_style": str(session.guard(lambda: paragraph.Range.Style.NameLocal)),
        "font": font,
        "paragraph": para,
        "deviations_from_style": deviations(session, font, para, style),
    }


def search_range(session, doc: Any, spec: TargetSpec):
    """The range a search runs against."""
    return story_range(doc, spec.container)


def find_all(session, spec: TargetSpec, stop_after: int | None = None) -> list[Match]:
    """Every match, without the ambiguity failure. Used by find_text.

    `stop_after` bounds the walk. Word redefines the searched range on every
    Execute and describing one match costs ~120 ms, so a caller that wants the
    first few matches must not pay for the rest. The scan itself is ~0.25 s.
    """
    doc = session.doc()
    host = search_range(session, doc, spec)
    # A header or footer has its own offset space, so it keeps its own paragraph
    # list; the body answers from the index instead of enumerating paragraphs.
    index = session.index() if spec.container == "body" else None
    paragraphs = [] if index is not None else session.guard(lambda: list(host.Paragraphs))
    def _setup():
        find = host.Find
        find.Text = spec.text or ""
        find.Forward = True
        find.Wrap = 0  # wdFindStop
        find.MatchCase = spec.match_case
        find.MatchWildcards = spec.wildcards
        if spec.style:
            find.Style = spec.style
        return find

    find = session.guard(_setup)
    found: list[Match] = []
    while session.guard(lambda: bool(find.Execute())):
        # Word redefines the searched range to the match, so Duplicate it before
        # the next Execute moves it on. doc.Range() would map these offsets into
        # the body and lose the match inside a header or footer.
        rng = session.guard(lambda: find.Parent.Duplicate)
        found.append(
            match_from_range(session, rng, index, paragraphs, spec.container, bool(spec.style))
        )
        if stop_after is not None and len(found) >= stop_after:
            break
        if len(found) >= MAX_CANDIDATES:
            break
    return found


def resolve(session, spec: TargetSpec) -> Match:
    """One location, or an exception carrying the candidates."""
    if spec.token:
        return TOKENS.get(spec.token)
    if spec.text or spec.style or spec.heading_path:
        # A heading_path filter runs after the search, so bounding the walk would
        # bound the wrong list. With an explicit occurrence, occurrence+1 matches
        # answer the request, and when the occurrence is out of range the loop
        # ends by itself so the reported count is still exact.
        bound = None
        if not spec.heading_path and spec.occurrence is not None:
            bound = spec.occurrence + 1
        found = find_all(session, spec, stop_after=bound)
        if spec.heading_path:
            narrowed = [m for m in found if spec.heading_path in m.heading_path]
            found = narrowed or found
        return choose(spec.text or spec.style or "", found, spec.occurrence)[0]
    raise InvalidArgument("Give a token, or text, style, or heading_path.")


def inspect_document(session, refresh: bool = False) -> dict:
    """Styles, comments, revisions, bookmarks, content controls, template, protection.

    Read once per document version: the styles scan alone is ~2-3 s. `refresh`
    re-reads it after a change made by hand in Word.
    """
    if refresh:
        session.set_inventory(None)
    cached = session.inventory()
    if cached is not None:
        return cached
    payload = _read_inventory(session)
    session.set_inventory(payload)
    return payload


def _read_inventory(session) -> dict:
    doc = session.doc()
    styles_raw = session.guard(lambda: raw_items(doc.Styles))
    styles = []
    for style in styles_raw:
        def _st(s=style):
            return (
                str(STYLE.get(s, "NameLocal")),
                bool(STYLE.get(s, "BuiltIn")),
                int(STYLE.get(s, "Type")),
                bool(STYLE.get(s, "InUse")),
            )

        name, built_in, typ, in_use = session.guard(_st)
        styles.append({"name": name, "built_in": built_in, "type": typ, "in_use": in_use})
    comments_raw = session.guard(lambda: raw_items(doc.Comments))
    comments = []
    for i, comment in enumerate(comments_raw, start=1):
        def _cm(c=comment):
            c = win32com.client.Dispatch(c)
            return (str(c.Author), str(c.Range.Text), bool(c.Done), str(c.Scope.Text)[:120])

        author, text, done, scope = session.guard(_cm)
        comments.append(
            {"index": i, "author": author, "text": text, "done": done, "scope": scope}
        )
    revs_raw = session.guard(lambda: raw_items(doc.Revisions))
    revisions = []
    for i, rev in enumerate(revs_raw, start=1):
        def _rv(r=rev):
            r = win32com.client.Dispatch(r)
            return (int(r.Type), str(r.Author), str(r.Range.Text)[:120])

        typ, author, text = session.guard(_rv)
        revisions.append({"index": i, "type": typ, "author": author, "text": text})
    bms_raw = session.guard(lambda: raw_items(doc.Bookmarks))
    bookmarks = []
    for bm in bms_raw:
        bookmarks.append(str(session.guard(lambda b=win32com.client.Dispatch(bm): b.Name)))
    ccs_raw = session.guard(lambda: raw_items(doc.ContentControls))
    controls = []
    for i, cc in enumerate(ccs_raw, start=1):
        def _cc(c=cc):
            c = win32com.client.Dispatch(c)
            return (int(c.Type), str(c.Tag), str(c.Title), str(c.Range.Text)[:120])

        typ, tag, title, text = session.guard(_cc)
        controls.append(
            {"index": i, "type": typ, "tag": tag, "title": title, "text": text}
        )
    # doc.UpdateStylesOnOpen is one of the unverified members flagged in the spec.
    # If it raises, drop that key rather than failing the whole tool.
    try:
        template_update = bool(session.guard(lambda: doc.UpdateStylesOnOpen))
    except Exception:
        template_update = None
    return {
        "template": str(session.guard(lambda: doc.AttachedTemplate.FullName)),
        "template_update_on_open": template_update,
        "track_revisions": bool(session.guard(lambda: doc.TrackRevisions)),
        "protection_type": int(session.guard(lambda: doc.ProtectionType)),
        "properties": {
            "title": str(session.guard(lambda: doc.BuiltInDocumentProperties("Title"))),
            "author": str(session.guard(lambda: doc.BuiltInDocumentProperties("Author"))),
        },
        "styles": styles,
        "comments": comments,
        "revisions": revisions,
        "bookmarks": bookmarks,
        "content_controls": controls,
    }


def get_ooxml(session, spec) -> dict:
    match = resolve(session, spec)
    return {"target": match.describe(), "ooxml": str(session.guard(lambda: match.handle.XML))}
