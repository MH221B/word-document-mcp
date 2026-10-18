"""Comments and revisions."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument
from word_mcp.operations import pick, require_confirm


def edit_annotations(
    session,
    resolve,
    action: str,
    spec=None,
    text: str = "",
    comment_index: int = 1,
    done: bool = True,
    mode: str = "",
    confirm: bool = False,
) -> dict:
    doc = session.doc()
    session.require_writable()

    if action == "add_comment":
        if not text:
            raise InvalidArgument("action='add_comment' needs text.")
        match = resolve(session, spec)
        comment = session.guard(lambda: doc.Comments.Add(Range=match.handle, Text=text))
        return {
            "comment": text,
            "author": str(session.guard(lambda: comment.Author)),
            "at": match.describe(),
        }

    if action == "reply_comment":
        comment = _comment(session, doc, comment_index)
        if not text:
            raise InvalidArgument("action='reply_comment' needs text.")
        replies = session.guard(lambda: comment.Replies)
        session.guard(lambda: replies.Add(comment.Scope, text))
        return {"reply": text, "to_comment": comment_index}

    if action == "resolve_comment":
        comment = _comment(session, doc, comment_index)
        session.guard(lambda: setattr(comment, "Done", done))
        return {"comment": comment_index, "done": done}

    if action == "delete_comment":
        comment = _comment(session, doc, comment_index)
        session.guard(lambda: comment.Delete())
        return {"deleted_comment": comment_index}

    if action == "track_changes":
        if mode not in ("on", "off", "mine"):
            raise InvalidArgument(f"Unknown mode {mode!r}.", allowed=["on", "off", "mine"])
        if mode == "on":
            session.guard(lambda: setattr(doc, "TrackRevisions", True))
        elif mode == "off":
            session.guard(lambda: setattr(doc, "TrackRevisions", False))
        else:
            session.guard(lambda: setattr(doc, "TrackRevisions", True))
            session.guard(lambda: setattr(doc, "TrackMoves", True))
        return {"track_changes": mode}

    if action == "accept_all" or action == "reject_all":
        require_confirm(confirm, f"{action} revisions")
        count = int(session.guard(lambda: doc.Revisions.Count))
        if action == "accept_all":
            session.guard(lambda: doc.Revisions.AcceptAll())
        else:
            session.guard(lambda: doc.Revisions.RejectAll())
        return {"revisions": count, "action": action}

    raise InvalidArgument(
        f"Unknown action {action!r}.",
        allowed=[
            "add_comment", "reply_comment", "resolve_comment", "delete_comment",
            "track_changes", "accept_all", "reject_all",
        ],
    )


def _comment(session, doc, index: int):
    return pick(session, doc.Comments, index, "comment", "comments")
