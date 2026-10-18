"""MCP tool registration. Validation and dispatch only, no COM here."""

from __future__ import annotations

import json
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from word_mcp import inspect
from word_mcp.errors import WordMcpError
from word_mcp.history import ChangeStack
from word_mcp.operations import content as content_ops
from word_mcp.operations import document as document_ops
from word_mcp.operations import ooxml as ooxml_ops
from word_mcp.operations import review as review_ops
from word_mcp.operations import structure as structure_ops
from word_mcp.operations import styles as style_ops
from word_mcp.operations import text as text_ops
from word_mcp.session import ComWordSession

mcp = FastMCP("word")
SESSION = ComWordSession()
HISTORY = ChangeStack()


def tool_error(exc: WordMcpError) -> str:
    return json.dumps({"error": exc.to_dict()}, indent=2)


def run(fn: Callable[[], Any]) -> str:
    """Call an operation, turning domain errors into tool output."""
    try:
        return json.dumps(fn(), indent=2, default=str)
    except WordMcpError as exc:
        return tool_error(exc)


def mutate(fn: Callable[[], Any]) -> str:
    """Call a mutating operation, settle the index, then attach `at`.

    An operation that reported the span it changed patches the index in place; a
    failed operation, and every operation that cannot describe itself, drops it.
    The post-edit anchor is built only when the index survived, so a multi-span
    batch never triggers a cold rebuild just to describe a range.
    """
    SESSION.begin_change()
    try:
        out = fn()
    except WordMcpError as exc:
        SESSION.invalidate_index()
        return tool_error(exc)
    SESSION.settle_index()
    stash = SESSION.take_result_range()
    if stash is not None and SESSION.has_index():
        handle, container = stash
        out["at"] = inspect.anchor(SESSION, handle, container).describe()
    return json.dumps(out, indent=2, default=str)


@mcp.tool()
def list_open_documents() -> str:
    """List every document currently open in Word."""
    def go():
        return {"documents": [d.__dict__ for d in SESSION.documents()]}
    return run(go)


@mcp.tool()
def connect_document(path_or_name: str) -> str:
    """Pin one open document for the session. Fails if the name is ambiguous."""
    def go():
        info = SESSION.pin(path_or_name)
        return {"pinned": info.__dict__, "read_only": info.read_only}
    return run(go)


@mcp.tool()
def disconnect_document() -> str:
    """Forget the pinned document."""
    def go():
        SESSION.unpin()
        inspect.TOKENS.clear()
        return {"pinned": None}
    return run(go)


@mcp.tool()
def document_outline(depth: int = 9, refresh: bool = False) -> str:
    """Headings, tables, sections, and images, each with an anchor token."""
    return run(lambda: inspect.document_outline(SESSION, depth, refresh))


@mcp.tool()
def read_range(target: dict, include_formatting: bool = False) -> str:
    """Read text at a target. Target is {"token": "..."} or {"text": "...", ...}."""
    return run(lambda: inspect.read_range(SESSION, inspect.TargetSpec.from_dict(target), include_formatting))


@mcp.tool()
def get_formatting(target: dict) -> str:
    """Formatting for a target, plus where direct formatting deviates from the style."""
    return run(lambda: inspect.get_formatting(SESSION, inspect.TargetSpec.from_dict(target)))


@mcp.tool()
def find_text(
    text: str,
    container: str = "body",
    heading_path: str | None = None,
    style: str | None = None,
    wildcards: bool = False,
    match_case: bool = False,
    limit: int = 20,
) -> str:
    """Find every match with its context. Never guesses; returns candidates."""
    def go():
        spec = inspect.TargetSpec.from_dict({
            "text": text, "container": container, "heading_path": heading_path,
            "style": style, "wildcards": wildcards, "match_case": match_case,
        })
        norm = inspect.normalise(text)
        # The heading_path filter runs after the search, so only bound the walk
        # when there is no filter to apply.
        stop = None if heading_path else limit
        out = []
        for match in inspect.find_all(SESSION, spec, stop_after=stop):
            if heading_path and heading_path not in match.heading_path:
                continue
            out.append({**match.describe(), "matched": norm in inspect.normalise(match.text)})
            if len(out) >= limit:
                break
        return {"searched": text, "container": container, "matches": out}
    return run(go)


@mcp.tool()
def inspect_document(refresh: bool = False) -> str:
    """Styles, comments, revisions, bookmarks, content controls, template, protection."""
    return run(lambda: inspect.inspect_document(SESSION, refresh))


@mcp.tool()
def get_ooxml(target: dict) -> str:
    """WordprocessingML for a target, for detailed inspection or copying."""
    return run(lambda: inspect.get_ooxml(SESSION, inspect.TargetSpec.from_dict(target)))


@mcp.tool()
def edit_text(
    target: dict,
    action: str,
    text: str = "",
    position: str = "end",
    format_mode: str = "inherit",
    formatting: dict | None = None,
    confirm: bool = False,
) -> str:
    """Insert, replace, or delete text. format_mode is inherit, style_only, or explicit."""
    def go():
        return text_ops.edit_text(
            SESSION, HISTORY, inspect.TargetSpec.from_dict(target), inspect.resolve,
            action=action, text=text, position=position, format_mode=format_mode,
            formatting=formatting, confirm=confirm,
        )
    return mutate(go)


@mcp.tool()
def format_text(
    target: dict,
    action: str,
    formatting: dict | None = None,
    source: dict | None = None,
    confirm: bool = False,
) -> str:
    """Character, paragraph, clear, or copy-from formatting for a target."""
    def go():
        return text_ops.format_text(
            SESSION, HISTORY, inspect.TargetSpec.from_dict(target), inspect.resolve,
            action=action, formatting=formatting, source=source, confirm=confirm,
        )
    return mutate(go)


@mcp.tool()
def apply_edits(edits: list[dict], confirm: bool = False) -> str:
    """Apply several text edits in one call. Stops at the first failure."""
    def go():
        return text_ops.apply_edits(SESSION, HISTORY, edits, inspect.resolve, confirm=confirm)
    return mutate(go)


@mcp.tool()
def revert_change(operation_id: int) -> str:
    """Revert the newest change. Reverts text and formatting only."""
    return mutate(lambda: document_ops.revert_change(SESSION, HISTORY, operation_id))


@mcp.tool()
def save_document() -> str:
    """Save the pinned document."""
    return run(lambda: document_ops.save_document(SESSION))


@mcp.tool()
def set_document_properties(properties: dict) -> str:
    """Set built-in (Title, Author, ...) or custom document properties."""
    return mutate(lambda: document_ops.set_document_properties(SESSION, properties))


@mcp.tool()
def manage_style(
    action: str,
    target: dict | None = None,
    name: str = "",
    based_on: str = "",
    formatting: dict | None = None,
    confirm: bool = False,
) -> str:
    """List, create, update, delete, or apply a style."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return style_ops.manage_style(
            SESSION, HISTORY, inspect.resolve, action=action, spec=spec, name=name,
            based_on=based_on, formatting=formatting, confirm=confirm,
        )
    return run(go) if action == "list" else mutate(go)


@mcp.tool()
def manage_list(
    target: dict,
    action: str,
    list_format: str = "bullet",
    level: int = 1,
    start_at: int = 1,
    confirm: bool = False,
) -> str:
    """Apply, continue, restart, or change the level of a list."""
    def go():
        return style_ops.manage_list(
            SESSION, HISTORY, inspect.resolve, action=action,
            spec=inspect.TargetSpec.from_dict(target), list_format=list_format, level=level,
            start_at=start_at, confirm=confirm,
        )
    return mutate(go)


@mcp.tool()
def edit_table(
    action: str,
    target: dict | None = None,
    table_index: int = 1,
    row: int = 0,
    column: int = 0,
    rows: int = 1,
    columns: int = 1,
    values: list | None = None,
    formatting: dict | None = None,
    label: str = "Table",
    title: str = "",
    position: str = "below",
    confirm: bool = False,
) -> str:
    """Insert or edit a table, including captions."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return structure_ops.edit_table(
            SESSION, HISTORY, inspect.resolve, action=action, spec=spec,
            table_index=table_index, row=row, column=column, rows=rows,
            columns=columns, values=values, formatting=formatting, label=label,
            title=title, position=position, confirm=confirm,
        )
    return mutate(go)


@mcp.tool()
def edit_page_layout(
    action: str,
    target: dict | None = None,
    section: int = 1,
    orientation: str = "",
    points: dict | None = None,
    columns: int = 1,
    side: str = "header",
    which: str = "primary",
    text: str = "",
    different_first_page: bool | None = None,
    odd_and_even: bool | None = None,
) -> str:
    """Page setup, section breaks, columns, headers and footers, and layout flags."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return structure_ops.edit_page_layout(
            SESSION, inspect.resolve, action=action, spec=spec, section=section,
            orientation=orientation, points=points, columns=columns, side=side,
            which=which, text=text, different_first_page=different_first_page,
            odd_and_even=odd_and_even,
        )
    return mutate(go)


@mcp.tool()
def edit_fields(
    action: str,
    target: dict | None = None,
    field: str = "page",
    toc_options: dict | None = None,
) -> str:
    """Insert fields, update fields, insert a table of contents or table of figures."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return structure_ops.edit_fields(
            SESSION, inspect.resolve, action=action, spec=spec, field=field,
            toc_options=toc_options,
        )
    return mutate(go)


@mcp.tool()
def edit_media(
    action: str,
    target: dict | None = None,
    image_index: int = 1,
    path: str = "",
    width: float | None = None,
    height: float | None = None,
    wrap: str = "inline",
    label: str = "Figure",
    title: str = "",
    position: str = "below",
) -> str:
    """Insert an image, size it, set wrapping, or add a caption."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return content_ops.edit_media(
            SESSION, inspect.resolve, action=action, spec=spec, image_index=image_index,
            path=path, width=width, height=height, wrap=wrap, label=label, title=title,
            position=position,
        )
    return mutate(go)


@mcp.tool()
def edit_links(
    action: str,
    target: dict | None = None,
    address: str = "",
    sub_address: str = "",
    screen_tip: str = "",
    display_text: str = "",
    name: str = "",
    text: str = "",
) -> str:
    """Hyperlinks, bookmarks, footnotes, and endnotes."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return content_ops.edit_links(
            SESSION, inspect.resolve, action=action, spec=spec, address=address,
            sub_address=sub_address, screen_tip=screen_tip, display_text=display_text,
            name=name, text=text,
        )
    return run(go) if action == "bookmark_find" else mutate(go)


@mcp.tool()
def edit_content_controls(
    action: str,
    target: dict | None = None,
    control_index: int = 1,
    control_type: str = "rich_text",
    tag: str = "",
    title: str = "",
    value: str = "",
    items: list | None = None,
) -> str:
    """Insert, set, or list content controls."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return content_ops.edit_content_controls(
            SESSION, inspect.resolve, action=action, spec=spec,
            control_index=control_index, control_type=control_type, tag=tag,
            title=title, value=value, items=items,
        )
    return run(go) if action == "list" else mutate(go)


@mcp.tool()
def edit_annotations(
    action: str,
    target: dict | None = None,
    text: str = "",
    comment_index: int = 1,
    done: bool = True,
    mode: str = "",
    confirm: bool = False,
) -> str:
    """Comments, track changes mode, and accepting or rejecting revisions."""
    def go():
        spec = inspect.TargetSpec.from_dict(target) if target else None
        return review_ops.edit_annotations(
            SESSION, inspect.resolve, action=action, spec=spec, text=text,
            comment_index=comment_index, done=done, mode=mode, confirm=confirm,
        )
    return mutate(go)


@mcp.tool()
def insert_ooxml(target: dict, ooxml: str, confirm: bool = False) -> str:
    """Insert a WordprocessingML fragment at a target. Use for watermarks and shapes."""
    def go():
        return ooxml_ops.insert_ooxml(
            SESSION, HISTORY, inspect.resolve, inspect.TargetSpec.from_dict(target), ooxml, confirm=confirm
        )
    return mutate(go)
