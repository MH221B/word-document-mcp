"""MCP tool registration. Validation and dispatch only, no COM here."""

from __future__ import annotations

import json
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from word_mcp import inspect
from word_mcp.errors import WordMcpError
from word_mcp.history import ChangeStack
from word_mcp.operations import document as document_ops
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
        return {"pinned": None}
    return run(go)


@mcp.tool()
def document_outline(depth: int = 9) -> str:
    """Headings, tables, sections, and images, each with an anchor token."""
    return run(lambda: inspect.document_outline(SESSION, depth))


@mcp.tool()
def read_range(target: dict, include_formatting: bool = False) -> str:
    """Read text at a target. Target is {"token": "..."} or {"text": "...", ...}."""
    return run(lambda: inspect.read_range(SESSION, inspect.TargetSpec(**target), include_formatting))


@mcp.tool()
def get_formatting(target: dict) -> str:
    """Formatting for a target, plus where direct formatting deviates from the style."""
    return run(lambda: inspect.get_formatting(SESSION, inspect.TargetSpec(**target)))


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
        spec = inspect.TargetSpec(
            text=text, container=container, heading_path=heading_path, style=style,
            wildcards=wildcards, match_case=match_case,
        )
        norm = inspect.normalise(text)
        out = []
        for match in inspect.find_all(SESSION, spec):
            if heading_path and heading_path not in match.heading_path:
                continue
            out.append({**match.describe(), "matched": norm in inspect.normalise(match.text)})
            if len(out) >= limit:
                break
        return {"searched": text, "container": container, "matches": out}
    return run(go)


@mcp.tool()
def inspect_document() -> str:
    """Styles, comments, revisions, bookmarks, content controls, template, protection."""
    return run(lambda: inspect.inspect_document(SESSION))


@mcp.tool()
def get_ooxml(target: dict) -> str:
    """WordprocessingML for a target, for detailed inspection or copying."""
    return run(lambda: inspect.get_ooxml(SESSION, inspect.TargetSpec(**target)))


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
            SESSION, HISTORY, inspect.TargetSpec(**target), inspect.resolve,
            action=action, text=text, position=position, format_mode=format_mode,
            formatting=formatting, confirm=confirm,
        )
    return run(go)


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
            SESSION, HISTORY, inspect.TargetSpec(**target), inspect.resolve,
            action=action, formatting=formatting, source=source, confirm=confirm,
        )
    return run(go)


@mcp.tool()
def revert_change(operation_id: int) -> str:
    """Revert the newest change. Reverts text and formatting only."""
    return run(lambda: document_ops.revert_change(SESSION, HISTORY, operation_id))


@mcp.tool()
def save_document() -> str:
    """Save the pinned document."""
    return run(lambda: document_ops.save_document(SESSION))


@mcp.tool()
def set_document_properties(properties: dict) -> str:
    """Set built-in (Title, Author, ...) or custom document properties."""
    return run(lambda: document_ops.set_document_properties(SESSION, properties))
