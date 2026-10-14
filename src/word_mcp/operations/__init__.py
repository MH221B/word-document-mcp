"""Shared validation for the operation modules."""

from word_mcp.errors import ConfirmationRequired, InvalidArgument

DELETE_PARA_THRESHOLD = 1
DELETE_CHAR_THRESHOLD = 500
RANGE_CHAR_THRESHOLD = 2000

FORMAT_MODES = ("inherit", "style_only", "explicit")


def check_format_mode(mode: str) -> str:
    if mode not in FORMAT_MODES:
        raise InvalidArgument(f"Unknown format_mode {mode!r}.", allowed=list(FORMAT_MODES))
    return mode


def gate_delete(text: str, confirm: bool) -> None:
    """Deletions big enough to hurt need an explicit confirm."""
    paragraphs = text.count("\r")
    if not confirm and (paragraphs > DELETE_PARA_THRESHOLD or len(text) > DELETE_CHAR_THRESHOLD):
        raise ConfirmationRequired(
            "This deletes a large amount of content. Pass confirm=true to proceed.",
            paragraphs=paragraphs,
            characters=len(text),
            paragraphs_threshold=DELETE_PARA_THRESHOLD,
            characters_threshold=DELETE_CHAR_THRESHOLD,
        )


def gate_range(length: int, confirm: bool) -> None:
    if not confirm and length > RANGE_CHAR_THRESHOLD:
        raise ConfirmationRequired(
            "This touches a large range. Pass confirm=true to proceed.",
            characters=length,
            characters_threshold=RANGE_CHAR_THRESHOLD,
        )


def require_confirm(confirm: bool, what: str) -> None:
    if not confirm:
        raise ConfirmationRequired(f"{what} requires confirm=true.")
