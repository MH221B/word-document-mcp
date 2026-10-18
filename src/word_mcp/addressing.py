"""Turn an agent request into concrete document locations, or into candidates."""

from __future__ import annotations

import re
import secrets
from dataclasses import dataclass, field, fields
from typing import Any

from word_mcp.errors import AmbiguousTarget, InvalidArgument, StaleToken, TargetNotFound

MAX_CANDIDATES = 50
CONTAINERS = {"body", "header", "footer", "footnote", "endnote", "comment"}


def split_container(container: str) -> tuple[str, int]:
    """Split "header:2" into ("header", 2). A bare name means section 1.

    The section rides in the container string because that string is already
    threaded through TargetSpec, Match, Change, and revert; a separate field
    would have to be carried by every one of them.
    """
    name, sep, suffix = container.partition(":")
    if not sep:
        return name, 1
    if not suffix.isdigit() or int(suffix) < 1:
        raise InvalidArgument(
            f"Container section must be a positive integer, got {suffix!r}.",
            container=container,
        )
    return name, int(suffix)


@dataclass
class TargetSpec:
    token: str | None = None
    text: str | None = None
    wildcards: bool = False
    match_case: bool = False
    style: str | None = None
    heading_path: str | None = None
    container: str = "body"
    occurrence: int | None = None

    @classmethod
    def from_dict(cls, data: dict) -> "TargetSpec":
        """Build a spec from a tool argument, rejecting unknown keys."""
        if not isinstance(data, dict):
            raise InvalidArgument("target must be an object.")
        allowed = {f.name for f in fields(cls)}
        unknown = sorted(set(data) - allowed)
        if unknown:
            raise InvalidArgument(
                f"Unknown target key(s) {unknown}.", allowed=sorted(allowed)
            )
        spec = cls(**data)
        name, _ = split_container(spec.container)
        if name not in CONTAINERS:
            raise InvalidArgument(
                f"Unknown container {spec.container!r}.", allowed=sorted(CONTAINERS)
            )
        return spec


@dataclass
class Match:
    token: str
    text: str
    paragraph_index: int
    heading_path: str
    style: str
    container: str
    snippet: str
    handle: Any = field(repr=False, default=None)

    def describe(self) -> dict:
        return {
            "token": self.token,
            "text": self.text,
            "paragraph_index": self.paragraph_index,
            "heading_path": self.heading_path,
            "style": self.style,
            "container": self.container,
            "snippet": self.snippet,
        }


class TokenStore:
    """Holds resolved locations for the life of the session.

    Each entry keeps a live COM range, so the store is capped: the oldest are
    evicted, and using an evicted token raises the same StaleToken the caller
    already handles. Tokens are deliberately kept across edits, because Word
    moves the ranges with the text.
    """

    def __init__(self, limit: int = 1000) -> None:
        self._items: dict[str, Match] = {}
        self._limit = limit

    def add(self, match: Match) -> str:
        token = match.token or secrets.token_hex(4)
        match.token = token
        self._items[token] = match
        while len(self._items) > self._limit:
            self._items.pop(next(iter(self._items)))
        return token

    def get(self, token: str) -> Match:
        match = self._items.get(token)
        if match is None:
            raise StaleToken(
                f"Anchor {token} is unknown. Read the document again to get a fresh one."
            )
        self.validate(match)
        return match

    def validate(self, match: Match) -> None:
        handle = match.handle
        if handle is None:
            return
        try:
            if handle.Start > handle.End:
                raise StaleToken(f"Anchor {match.token} no longer covers any content.")
        except StaleToken:
            raise
        except Exception as exc:  # the range's document was closed
            del self._items[match.token]
            raise StaleToken(
                f"Anchor {match.token} is no longer valid; the document changed."
            ) from exc

    def clear(self) -> None:
        self._items.clear()


def normalise(text: str) -> str:
    """Agents and Word disagree about whitespace and quotes. Meet in the middle."""
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = text.replace("\u201c", '"').replace("\u201d", '"')
    return re.sub(r"\s+", " ", text).strip()


def choose(text: str, candidates: list[Match], occurrence: int | None) -> list[Match]:
    """Narrow matches by occurrence, or fail and hand the candidates back."""
    if occurrence is not None:
        if occurrence < 0 or occurrence >= len(candidates):
            raise TargetNotFound(
                f"Occurrence {occurrence} is out of range.",
                found=len(candidates),
            )
        return [candidates[occurrence]]
    if len(candidates) > 1:
        raise AmbiguousTarget(
            f"{len(candidates)} locations match. Narrow with heading_path, "
            "container, occurrence, or use one of the returned tokens.",
            candidates=[c.describe() for c in candidates[:MAX_CANDIDATES]],
        )
    if not candidates:
        raise TargetNotFound(f"Nothing matched {text!r}.")
    return candidates
