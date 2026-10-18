"""Batched paragraph scan through an ephemeral VBA module.

The scanner runs inside Word and returns every paragraph in one call,
so offsets need no ordinal mapping. It lives in a throwaway document
that is closed without saving: user docs, Normal, and disk stay clean.
Any denial falls back to the COM path via macros_unavailable.
"""

from __future__ import annotations

from typing import Any

from word_mcp.errors import (
    MacrosUnavailable,
    WordBusy,
    WordDisconnected,
    WordNotRunning,
)


class ScanFailed(Exception):
    """The VBA scan could not produce a usable result; fall back to COM."""

VBA_PROBE = "Function ProbePing() As Long\r\n    ProbePing = 1\r\nEnd Function\r\n"

VBA_SCAN = (
    "Function ScanDoc(fullName As String) As Variant\r\n"
    "    Dim d As Document\r\n"
    "    Dim p As Paragraph\r\n"
    "    Dim i As Long, n As Long, h As Long\r\n"
    "    Dim t As Long, s As Long, g As Long\r\n"
    "    Dim starts As Variant\r\n"
    "    Dim heads As Variant\r\n"
    "    Dim tbls As Variant\r\n"
    "    Dim secs As Variant\r\n"
    "    Dim imgs As Variant\r\n"
    "    For Each d In Application.Documents\r\n"
    "        If StrComp(d.FullName, fullName, vbTextCompare) = 0 Then Exit For\r\n"
    "    Next d\r\n"
    "    n = d.Paragraphs.Count\r\n"
    "    ReDim starts(1 To n)\r\n"
    # heads is over-allocated: only the first h rows are headings, so the loop
    # stays one pass with no ReDim Preserve (which can only resize the last
    # dimension). Body text cells stay Empty and never cross COM.
    "    ReDim heads(1 To n, 1 To 5)\r\n"
    "    h = 0\r\n"
    "    For Each p In d.Paragraphs\r\n"
    "        i = i + 1\r\n"
    "        starts(i) = p.Range.Start\r\n"
    "        If p.OutlineLevel < 10 Then\r\n"
    "            h = h + 1\r\n"
    "            heads(h, 1) = i - 1\r\n"
    "            heads(h, 2) = p.OutlineLevel\r\n"
    "            heads(h, 3) = p.Range.Style.NameLocal\r\n"
    "            heads(h, 4) = p.Range.Text\r\n"
    "            heads(h, 5) = p.Range.End\r\n"
    "        End If\r\n"
    "    Next p\r\n"
    "    t = d.Tables.Count\r\n"
    "    If t > 0 Then\r\n"
    "        ReDim tbls(1 To t, 1 To 2)\r\n"
    "        For i = 1 To t\r\n"
    "            tbls(i, 1) = d.Tables(i).Rows.Count\r\n"
    "            tbls(i, 2) = d.Tables(i).Columns.Count\r\n"
    "        Next i\r\n"
    "    Else\r\n"
    "        tbls = Empty\r\n"
    "    End If\r\n"
    "    s = d.Sections.Count\r\n"
    "    If s > 0 Then\r\n"
    "        ReDim secs(1 To s, 1 To 3)\r\n"
    "        For i = 1 To s\r\n"
    "            secs(i, 1) = d.Sections(i).PageSetup.PageWidth\r\n"
    "            secs(i, 2) = d.Sections(i).PageSetup.PageHeight\r\n"
    "            secs(i, 3) = d.Sections(i).PageSetup.Orientation\r\n"
    "        Next i\r\n"
    "    Else\r\n"
    "        secs = Empty\r\n"
    "    End If\r\n"
    "    g = d.InlineShapes.Count\r\n"
    "    If g > 0 Then\r\n"
    "        ReDim imgs(1 To g, 1 To 2)\r\n"
    "        For i = 1 To g\r\n"
    "            imgs(i, 1) = d.InlineShapes(i).Type\r\n"
    "            imgs(i, 2) = d.InlineShapes(i).Range.Paragraphs(1).Range.Text\r\n"
    "        Next i\r\n"
    "    Else\r\n"
    "        imgs = Empty\r\n"
    "    End If\r\n"
    "    ScanDoc = Array(starts, heads, h, tbls, secs, imgs)\r\n"
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


def scan(session: Any, doc: Any) -> tuple:
    """One VBA call: (starts, heads, h, tables, sections, images); blocks may be None."""
    app = session.app()
    target = str(session.guard(lambda: doc.FullName))
    tmp = session.guard(lambda: app.Documents.Add())
    try:
        mod = session.guard(lambda: _inject(tmp, "McpScan", VBA_SCAN))
        try:
            raw = session.guard(lambda: app.Run("McpScan.ScanDoc", target))
        except MacrosUnavailable:
            raise
        except (WordBusy, WordDisconnected, WordNotRunning):
            raise
        except Exception as exc:
            raise ScanFailed(str(exc)) from exc
        finally:
            session.guard(lambda: tmp.VBProject.VBComponents.Remove(mod))
    finally:
        session.guard(lambda: tmp.Close(SaveChanges=0))
    return tuple(raw)


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


def shapes_from_scan(raw: tuple) -> tuple[list[int], list, dict, list, list, list]:
    """StructureIndex parts from one ScanDoc return, with no COM.

    Heads columns are (para_index, level, style, text, end); only the first
    h rows are headings. Table/section/image blocks mirror _summaries shapes.
    A short or ragged return raises, so build() falls back to COM.
    """
    starts = [int(v) for v in raw[0]]
    h = int(raw[2])
    heading_rows: list[tuple[int, int, Any]] = []
    detail: dict[int, tuple[str, str, int]] = {}
    for r in [tuple(x) for x in raw[1]][:h]:
        idx0, level, style, text, end = (
            int(r[0]),
            int(r[1]),
            str(r[2]),
            str(r[3]).strip(),
            int(r[4]),
        )
        heading_rows.append((idx0, level, None))
        detail[idx0] = (text, style, max(starts[idx0], end - 1))
    tables = [
        {"table": n + 1, "rows": int(r[0]), "columns": int(r[1])}
        for n, r in enumerate([tuple(x) for x in raw[3]] if raw[3] is not None else [])
    ]
    sections = [
        {
            "section": n + 1,
            "page_width": float(r[0]),
            "page_height": float(r[1]),
            "orientation": int(r[2]),
        }
        for n, r in enumerate([tuple(x) for x in raw[4]] if raw[4] is not None else [])
    ]
    images = [
        {"inline_shape": n + 1, "type": int(r[0]), "caption_paragraph": str(r[1]).strip()}
        for n, r in enumerate([tuple(x) for x in raw[5]] if raw[5] is not None else [])
    ]
    return starts, heading_rows, detail, tables, sections, images
