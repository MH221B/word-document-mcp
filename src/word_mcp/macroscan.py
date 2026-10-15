"""Batched paragraph scan through an ephemeral VBA module.

The scanner runs inside Word and returns every paragraph in one call,
so offsets need no ordinal mapping. It lives in a throwaway document
that is closed without saving: user docs, Normal, and disk stay clean.
Any denial falls back to the COM path via macros_unavailable.
"""

from __future__ import annotations

from typing import Any

VBA_PROBE = "Function ProbePing() As Long\r\n    ProbePing = 1\r\nEnd Function\r\n"

VBA_SCAN = (
    "Function ScanDoc(fullName As String) As Variant\r\n"
    "    Dim d As Document\r\n"
    "    Dim i As Long, n As Long\r\n"
    "    Dim out() As Variant\r\n"
    "    For Each d In Application.Documents\r\n"
    "        If StrComp(d.FullName, fullName, vbTextCompare) = 0 Then Exit For\r\n"
    "    Next d\r\n"
    "    n = d.Paragraphs.Count\r\n"
    "    ReDim out(1 To n, 1 To 5)\r\n"
    "    For i = 1 To n\r\n"
    "        out(i, 1) = d.Paragraphs(i).Range.Start\r\n"
    "        out(i, 2) = d.Paragraphs(i).Range.End\r\n"
    "        out(i, 3) = d.Paragraphs(i).OutlineLevel\r\n"
    "        out(i, 4) = d.Paragraphs(i).Range.Style.NameLocal\r\n"
    "        out(i, 5) = d.Paragraphs(i).Range.Text\r\n"
    "    Next i\r\n"
    "    ScanDoc = out\r\n"
    "End Function\r\n"
)


def _inject(doc: Any, name: str, source: str) -> Any:
    mod = doc.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
    mod.Name = name
    mod.CodeModule.AddFromString(source)
    return mod


def probe(session: Any) -> bool:
    """True when a VBA scan can run here. The session caches the answer."""
    app = session.app()
    tmp = session.guard(lambda: app.Documents.Add())
    try:
        mod = session.guard(lambda: _inject(tmp, "McpProbe", VBA_PROBE))
        try:
            return int(session.guard(lambda: app.Run("McpProbe.ProbePing"))) == 1
        finally:
            session.guard(lambda: tmp.VBProject.VBComponents.Remove(mod))
    finally:
        session.guard(lambda: tmp.Close(SaveChanges=0))


def scan(session: Any, doc: Any) -> list[tuple]:
    """Every paragraph as (start, end, level, style, text), in one VBA call."""
    app = session.app()
    target = str(session.guard(lambda: doc.FullName))
    tmp = session.guard(lambda: app.Documents.Add())
    try:
        mod = session.guard(lambda: _inject(tmp, "McpScan", VBA_SCAN))
        try:
            raw = session.guard(lambda: app.Run("McpScan.ScanDoc", target))
        finally:
            session.guard(lambda: tmp.VBProject.VBComponents.Remove(mod))
    finally:
        session.guard(lambda: tmp.Close(SaveChanges=0))
    return [tuple(r) for r in raw]


def rows_to_shapes(rows: list[tuple]) -> tuple[list[int], list, dict]:
    """Shared (starts, heading_rows, detail) shape, from VBA rows.

    Matches the _read_paragraphs contract so _headings is reused as is.
    Level 10 is body text, mirroring WD_OUTLINE_BODY_TEXT in docindex.
    """
    starts: list[int] = []
    heading_rows: list[tuple[int, int, Any]] = []
    detail: dict[int, tuple[str, str, int]] = {}
    for r in rows:
        start, end, level, style, text = (
            int(r[0]),
            int(r[1]),
            int(r[2]),
            str(r[3]),
            str(r[4]).strip(),
        )
        idx = len(starts)
        starts.append(start)
        if level < 10:
            heading_rows.append((idx, level, None))
            detail[idx] = (text, style, max(start, end - 1))
    return starts, heading_rows, detail
