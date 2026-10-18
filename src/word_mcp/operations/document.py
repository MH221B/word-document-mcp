"""Document-level tools."""

from __future__ import annotations

from word_mcp.errors import InvalidArgument
from word_mcp.history import ChangeStack, revert


def revert_change(session, stack: ChangeStack, operation_id: int) -> dict:
    """Revert the newest change. Older changes must be reverted in order."""
    return revert(session, stack, operation_id)


def save_document(session) -> dict:
    doc = session.doc()
    session.require_writable()
    session.guard(lambda: doc.Save())
    return {"saved": session.pinned().__dict__}


BUILT_IN_ALLOWED = {"Title", "Author", "Subject", "Category", "Comments", "Keywords", "Company"}


def set_document_properties(session, properties: dict) -> dict:
    doc = session.doc()
    session.require_writable()
    applied = {}
    for name, value in properties.items():
        if name in BUILT_IN_ALLOWED:
            session.guard(lambda n=name, v=value: setattr(doc.BuiltInDocumentProperties(n), "Value", v))
        else:
            session.guard(lambda n=name, v=value: _set_custom(doc, n, v))
        applied[name] = value
    return {"applied": applied}


def _set_custom(doc, name: str, value) -> None:
    custom = doc.CustomDocumentProperties
    for prop in custom:
        if prop.Name == name:
            prop.Value = value
            return
    custom.Add(Name=name, LinkToContent=False, Type=4, Value=str(value))  # msoPropertyTypeString
