"""Snapshot stack and revert. Reverts the newest change only."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from word_mcp.errors import InvalidArgument, WordMcpError
from word_mcp.session import story_range

MAX_DEPTH = 20


@dataclass
class Change:
    op_id: int
    kind: str  # "text" or "format"
    description: str
    start: int
    end: int
    container: str = "body"
    text_before: str = ""
    format_before: dict = field(default_factory=dict)
    reverted: bool = False


class ChangeStack:
    def __init__(self, depth: int = MAX_DEPTH):
        self._items: list[Change] = []
        self._depth = depth
        self._next_id = 1

    def record(self, kind: str, description: str, rng: Any, **kwargs) -> Change:
        change = Change(
            op_id=self._next_id,
            kind=kind,
            description=description,
            start=int(rng.Start),
            end=int(rng.End),
            **kwargs,
        )
        self._next_id += 1
        self._items.append(change)
        del self._items[: max(0, len(self._items) - self._depth)]
        return change

    def newest(self) -> Change | None:
        return self._items[-1] if self._items else None

    def by_id(self, op_id: int) -> Change:
        for change in reversed(self._items):
            if change.op_id == op_id:
                return change
        raise InvalidArgument(f"No recorded change with id {op_id}.")

    def pop(self) -> Change:
        if not self._items:
            raise WordMcpError("There is nothing to revert.")
        return self._items.pop()

    def clear(self) -> None:
        self._items.clear()


def revert(session, stack: ChangeStack, op_id: int) -> dict:
    """Revert one change. Only the newest change can be reverted."""
    change = stack.by_id(op_id)
    newest = stack.newest()
    if newest is None or newest.op_id != change.op_id:
        raise InvalidArgument(
            f"Only the newest change can be reverted. Newest is {newest.op_id if newest else None}.",
            requested=op_id,
        )
    doc = session.doc()
    # Offsets are relative to the story the change was made in, so the range is
    # rebuilt from that story rather than from the body.
    rng = session.guard(lambda: story_range(doc, change.container).Duplicate)
    session.guard(lambda: setattr(rng, "Start", change.start))
    session.guard(lambda: setattr(rng, "End", change.end))
    if change.kind == "text":
        session.guard(lambda: setattr(rng, "Text", change.text_before))
    elif change.kind == "format":
        for prop, value in change.format_before.items():
            target = rng.Font if hasattr(rng.Font, prop) else rng.ParagraphFormat
            session.guard(lambda t=target, p=prop, v=value: setattr(t, p, v))
    else:
        raise InvalidArgument(f"Cannot revert a {change.kind} change.")
    stack.pop()
    return {
        "reverted": change.op_id,
        "kind": change.kind,
        "description": change.description,
        "limits": "Text and formatting only. Deleted images, fields, and comments are not restored.",
    }
