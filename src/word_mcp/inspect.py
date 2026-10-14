"""Read-only tools: outline, read, find, formatting, document state, OOXML."""

from __future__ import annotations

from typing import Any

from word_mcp.addressing import (
    MAX_CANDIDATES,
    Match,
    TargetSpec,
    TokenStore,
    choose,
    normalise,
)
from word_mcp.errors import InvalidArgument
from word_mcp.session import story_range

TOKENS = TokenStore()
WD_OUTLINE_BODY_TEXT = 10


def heading_path_of(paragraphs: list, paragraph_index: int) -> str:
    """Walk back to the nearest enclosing headings, nearest last."""
    parts: list[tuple[int, str]] = []
    cap = WD_OUTLINE_BODY_TEXT
    for i in range(paragraph_index - 1, -1, -1):
        para = paragraphs[i]
        level = int(para.OutlineLevel)
        if level < cap:
            cap = level
            parts.insert(0, (level, para.Range.Text.strip()))
            if level == 1:
                break
    return " > ".join(text for _, text in parts)


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
    paragraphs: list,
    paragraph_index: int = -1,
    container: str = "body",
) -> Match:
    """Wrap a Word range as a Match and register an anchor token for it.

    `paragraphs` is the paragraph list of the range's own story, so headings and
    indices stay correct inside a header or footer.
    """
    if paragraph_index < 0:
        paragraph_index, path, style_name = paragraph_context(session, handle, paragraphs)
    else:
        path = heading_path_of(paragraphs, paragraph_index)
        style_name = str(session.guard(lambda: handle.Style.NameLocal))
    text = str(session.guard(lambda: handle.Text)).replace("\r", "").replace("\x07", "")
    match = Match(
        token="",
        text=text,
        paragraph_index=paragraph_index,
        heading_path=path,
        style=style_name,
        container=container,
        snippet=text[:80],
        handle=handle,
    )
    TOKENS.add(match)
    return match


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
    headings = []
    for heading in index.headings:
        if heading.level > depth:
            continue
        handle = session.guard(lambda h=heading: doc.Range(h.start, h.end))
        match = Match(
            token="",
            text=heading.text,
            paragraph_index=index.paragraph_index(heading.start),
            heading_path=heading.path,
            style=heading.style,
            container="body",
            snippet=heading.text[:80],
            handle=handle,
        )
        TOKENS.add(match)
        headings.append({"level": heading.level, **match.describe()})
    return {
        "document": session.pinned().__dict__,
        "headings": headings,
        "tables": index.tables,
        "sections": index.sections,
        "images": index.images,
    }


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


def deviations(session, handle) -> dict:
    """Properties where direct formatting differs from the style definition."""
    style = session.guard(lambda: handle.Style)
    direct = _snapshot(handle.Font, FONT_PROPS, session)
    inherited = _snapshot(style.Font, FONT_PROPS, session)
    para_direct = _snapshot(handle.ParagraphFormat, PARA_PROPS, session)
    para_style = _snapshot(style.ParagraphFormat, PARA_PROPS, session)
    return {
        prop: {"direct": value, "style": inherited.get(prop)}
        for prop, value in {**direct, **para_direct}.items()
        if prop in {**inherited, **para_style} and value != {**inherited, **para_style}[prop]
    }


def read_range(session, spec, include_formatting: bool = False) -> dict:
    match = resolve(session, spec)
    text = session.guard(lambda: match.handle.Text)
    payload = {"target": match.describe(), "text": text}
    if include_formatting:
        payload["formatting"] = get_formatting(session, spec)
    return payload


def get_formatting(session, spec) -> dict:
    match = resolve(session, spec)
    handle = match.handle
    paragraph = session.guard(lambda: handle.Paragraphs(1))
    payload = {
        "target": match.describe(),
        "paragraph_style": str(session.guard(lambda: paragraph.Style.NameLocal)),
        "character_style": str(session.guard(lambda: handle.Style.NameLocal)),
        "surrounding_style": str(session.guard(lambda: handle.Paragraphs(1).Range.Style.NameLocal)),
        "font": _snapshot(handle.Font, FONT_PROPS, session),
        "paragraph": _snapshot(handle.ParagraphFormat, PARA_PROPS, session),
        "deviations_from_style": deviations(session, handle),
    }
    return payload


def search_range(session, doc: Any, spec: TargetSpec):
    """The range a search runs against."""
    return story_range(doc, spec.container)


def find_all(session, spec: TargetSpec) -> list[Match]:
    """Every match, without the ambiguity failure. Used by find_text."""
    doc = session.doc()
    host = search_range(session, doc, spec)
    # Offsets in a found range are relative to its own story, so the paragraph
    # list has to come from that story too.
    paragraphs = session.guard(lambda: list(host.Paragraphs))
    find = session.guard(lambda: host.Find)
    session.guard(lambda: setattr(find, "Text", spec.text or ""))
    session.guard(lambda: setattr(find, "Forward", True))
    session.guard(lambda: setattr(find, "Wrap", 0))  # wdFindStop
    session.guard(lambda: setattr(find, "MatchCase", spec.match_case))
    session.guard(lambda: setattr(find, "MatchWildcards", spec.wildcards))
    if spec.style:
        session.guard(lambda: setattr(find, "Style", spec.style))
    found: list[Match] = []
    while session.guard(lambda: bool(find.Execute())):
        # Word redefines the searched range to the match, so Duplicate it before
        # the next Execute moves it on. doc.Range() would map these offsets into
        # the body and lose the match inside a header or footer.
        rng = session.guard(lambda: find.Parent.Duplicate)
        found.append(match_from_range(session, rng, paragraphs, container=spec.container))
        if len(found) >= MAX_CANDIDATES:
            break
    return found


def resolve(session, spec: TargetSpec) -> Match:
    """One location, or an exception carrying the candidates."""
    if spec.token:
        return TOKENS.get(spec.token)
    if spec.text or spec.style or spec.heading_path:
        found = find_all(session, spec)
        if spec.heading_path:
            narrowed = [m for m in found if spec.heading_path in m.heading_path]
            found = narrowed or found
        return choose(spec.text or spec.style or "", found, spec.occurrence)[0]
    raise InvalidArgument("Give a token, or text, style, or heading_path.")


def inspect_document(session) -> dict:
    doc = session.doc()
    styles = [
        {
            "name": str(session.guard(lambda s=style: s.NameLocal)),
            "built_in": bool(session.guard(lambda s=style: s.BuiltIn)),
            "type": int(session.guard(lambda s=style: s.Type)),
            "in_use": bool(session.guard(lambda s=style: s.InUse)),
        }
        for style in session.guard(lambda: list(doc.Styles))
    ]
    comments = [
        {
            "index": i,
            "author": str(session.guard(lambda c=comment: c.Author)),
            "text": str(session.guard(lambda c=comment: c.Range.Text)),
            "done": bool(session.guard(lambda c=comment: c.Done)),
            "scope": str(session.guard(lambda c=comment: c.Scope.Text))[:120],
        }
        for i, comment in enumerate(session.guard(lambda: list(doc.Comments)), start=1)
    ]
    revisions = [
        {
            "index": i,
            "type": int(session.guard(lambda r=rev: r.Type)),
            "author": str(session.guard(lambda r=rev: r.Author)),
            "text": str(session.guard(lambda r=rev: r.Range.Text))[:120],
        }
        for i, rev in enumerate(session.guard(lambda: list(doc.Revisions)), start=1)
    ]
    bookmarks = [
        str(session.guard(lambda b=bm: b.Name))
        for bm in session.guard(lambda: list(doc.Bookmarks))
    ]
    controls = [
        {
            "index": i,
            "type": int(session.guard(lambda c=cc: c.Type)),
            "tag": str(session.guard(lambda c=cc: c.Tag)),
            "title": str(session.guard(lambda c=cc: c.Title)),
            "text": str(session.guard(lambda c=cc: c.Range.Text))[:120],
        }
        for i, cc in enumerate(session.guard(lambda: list(doc.ContentControls)), start=1)
    ]
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
