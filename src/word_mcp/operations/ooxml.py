"""The OOXML escape hatch, for watermarks and other constructs with no object model.

COM takes and returns flat WordprocessingML, not the pkg:package wrapper the
Office.js API uses.
"""

from __future__ import annotations

from word_mcp.errors import InvalidArgument
from word_mcp.history import ChangeStack
from word_mcp.session import require_writable


def insert_ooxml(
    session,
    stack: ChangeStack,
    resolve,
    spec,
    ooxml: str,
    confirm: bool = False,
) -> dict:
    doc = session.doc()
    require_writable(doc)
    if not ooxml.strip().startswith("<"):
        raise InvalidArgument("ooxml must start with an XML element.")
    if "<w:" not in ooxml and "xmlns:w" not in ooxml:
        raise InvalidArgument(
            "ooxml must be WordprocessingML, for example <w:p>...</w:p> with the w namespace."
        )
    match = resolve(session, spec)
    handle = match.handle
    before = str(session.guard(lambda: handle.Text))
    change = stack.record("text", "insert ooxml", handle, text_before=before)
    session.guard(lambda: handle.InsertXML(ooxml))
    return {
        "operation_id": change.op_id,
        "inserted": len(ooxml),
        "at": match.describe(),
        "note": "Inserted content takes the destination section's properties unless the fragment carries its own sectPr.",
    }
