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

TOKENS = TokenStore()
WD_OUTLINE_BODY_TEXT = 10


def heading_path_of(doc: Any, paragraph_index: int) -> str:
    """Walk back to the nearest enclosing headings, nearest last."""
    parts: list[tuple[int, str]] = []
    cap = WD_OUTLINE_BODY_TEXT
    for i in range(paragraph_index - 1, -1, -1):
        para = doc.Paragraphs(i + 1)
        level = int(para.OutlineLevel)
        if level < cap:
            cap = level
            parts.insert(0, (level, para.Range.Text.strip()))
            if level == 1:
                break
    return " > ".join(text for _, text in parts)


def paragraph_context(session, doc: Any, rng: Any) -> tuple[int, str, str]:
    """Paragraph index, heading path, and style for a range."""
    paragraphs = session.guard(lambda: list(doc.Paragraphs))
    for i, para in enumerate(paragraphs):
        para_range = session.guard(lambda p=para: p.Range)
        if session.guard(lambda r=para_range: r.Start <= rng.Start < r.End):
            style = str(session.guard(lambda p=para: p.Range.Style.NameLocal))
            return i, heading_path_of(doc, i), style
    return -1, "", ""


def match_from_range(
    session, doc: Any, handle: Any, paragraph_index: int = -1, container: str = "body"
) -> Match:
    """Wrap a Word range as a Match and register an anchor token for it."""
    if paragraph_index < 0:
        paragraph_index, path, style_name = paragraph_context(session, doc, handle)
    else:
        path = heading_path_of(doc, paragraph_index)
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


def document_outline(session, depth: int = 9) -> dict:
    doc = session.doc()
    paragraphs = session.guard(lambda: list(doc.Paragraphs))
    headings = []
    for i, para in enumerate(paragraphs):
        level = int(session.guard(lambda p=para: p.OutlineLevel))
        if level < WD_OUTLINE_BODY_TEXT and level <= depth:
            match = match_from_range(session, doc, para.Range, i)
            own_path = " > ".join([p for p in [match.heading_path, match.text] if p])
            headings.append({"level": level, "heading_path": own_path, **match.describe()})
    tables = []
    for t_index, table in enumerate(session.guard(lambda: list(doc.Tables)), start=1):
        rows = int(session.guard(lambda t=table: t.Rows.Count))
        cols = int(session.guard(lambda t=table: t.Columns.Count))
        tables.append({"table": t_index, "rows": rows, "columns": cols})
    sections = []
    for s_index, section in enumerate(session.guard(lambda: list(doc.Sections)), start=1):
        sections.append(
            {
                "section": s_index,
                "page_width": float(session.guard(lambda s=section: s.PageSetup.PageWidth)),
                "page_height": float(session.guard(lambda s=section: s.PageSetup.PageHeight)),
                "orientation": int(session.guard(lambda s=section: s.PageSetup.Orientation)),
            }
        )
    images = []
    for i_index, shape in enumerate(session.guard(lambda: list(doc.InlineShapes)), start=1):
        images.append(
            {
                "inline_shape": i_index,
                "type": int(session.guard(lambda s=shape: s.Type)),
                "caption_paragraph": session.guard(
                    lambda s=shape: s.Range.Paragraphs(1).Range.Text.strip()
                ),
            }
        )
    return {
        "document": session.pinned().__dict__,
        "headings": headings,
        "tables": tables,
        "sections": sections,
        "images": images,
    }
