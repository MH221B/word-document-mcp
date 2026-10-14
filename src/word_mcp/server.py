"""MCP tool registration. Validation and dispatch only, no COM here."""

from __future__ import annotations

import json
from typing import Any, Callable

from mcp.server.fastmcp import FastMCP

from word_mcp import inspect
from word_mcp.errors import WordMcpError
from word_mcp.session import ComWordSession

mcp = FastMCP("word")
SESSION = ComWordSession()


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
