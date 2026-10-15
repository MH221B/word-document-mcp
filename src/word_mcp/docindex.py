"""The cached structure of one document version: paragraph offsets and headings."""

from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
from typing import Any

from word_mcp.comfast import PARAGRAPH, RANGE, STYLE, raw_items

WD_OUTLINE_BODY_TEXT = 10


@dataclass
class Heading:
    level: int
    style: str
    text: str
    start: int
    end: int
    path: str  # ancestors only, nearest last


class StructureIndex:
    """Paragraph offsets and headings, read once per document version.

    `paragraph_index` and `heading_path` answer from the stored offsets, so a
    lookup costs no COM at all.
    """

    def __init__(
        self,
        fingerprint: int,
        paragraph_starts: list[int],
        headings: list[Heading],
        tables: list[dict],
        sections: list[dict],
        images: list[dict],
        paragraph_count: int = 0,
    ) -> None:
        self.fingerprint = fingerprint
        self.paragraph_starts = paragraph_starts
        self.headings = headings
        self.heading_starts = [heading.start for heading in headings]
        self.tables = tables
        self.sections = sections
        self.images = images
        # The paragraph count is the guard that catches an edit which added or
        # removed a paragraph mark, which would invalidate the whole offset shift.
        self.paragraph_count = paragraph_count
        # Minted on the first document_outline: one (token, handle) per heading.
        # Live Word ranges move with the document, so a patched index keeps them.
        self.anchors: list[tuple[str, Any]] | None = None

    @classmethod
    def build(cls, session: Any, fingerprint: int, paragraph_count: int) -> "StructureIndex":
        doc = session.doc()
        # One guard around the whole pass: a busy Word retries the read rather
        # than raising a raw COM error halfway through.
        starts, heading_rows, detail = session.guard(lambda: cls._read_paragraphs(doc))
        tables, sections, images = cls._summaries(session, doc)
        return cls(
            fingerprint,
            starts,
            cls._headings(heading_rows, starts, detail),
            tables,
            sections,
            images,
            paragraph_count,
        )

    @staticmethod
    def _read_paragraphs(
        doc: Any,
    ) -> tuple[list[int], list[tuple[int, int, Any]], dict[int, tuple[str, str, int]]]:
        """One pass: starts, heading rows, plus heading detail by para index."""
        starts: list[int] = []
        heading_rows: list[tuple[int, int, Any]] = []
        detail: dict[int, tuple[str, str, int]] = {}
        for para in raw_items(doc.Paragraphs):
            rng = PARAGRAPH.get(para, "Range")
            start = RANGE.get(rng, "Start")
            level = int(PARAGRAPH.get(para, "OutlineLevel"))
            idx = len(starts)
            starts.append(start)
            if level < WD_OUTLINE_BODY_TEXT:
                heading_rows.append((idx, level, para))
                end = max(start, RANGE.get(rng, "End") - 1)
                text = str(RANGE.get(rng, "Text")).strip()
                style = str(STYLE.get(RANGE.get(rng, "Style"), "NameLocal"))
                detail[idx] = (text, style, end)
        return starts, heading_rows, detail

    @staticmethod
    def _headings(
        rows: list[tuple[int, int, Any]],
        starts: list[int],
        detail: dict[int, tuple[str, str, int]],
    ) -> list[Heading]:
        headings: list[Heading] = []
        stack: list[Heading] = []
        for idx, level, _para in rows:
            while stack and stack[-1].level >= level:
                stack.pop()
            start = starts[idx]
            text, style, end = detail[idx]
            headings.append(
                Heading(
                    level=level,
                    style=style,
                    text=text,
                    start=start,
                    end=end,
                    path=" > ".join(h.text for h in stack),
                )
            )
            stack.append(headings[-1])
        return headings

    @staticmethod
    def _summaries(session: Any, doc: Any) -> tuple[list[dict], list[dict], list[dict]]:
        tables_raw = session.guard(lambda: raw_items(doc.Tables))
        tables = []
        for index, table in enumerate(tables_raw, start=1):
            def _t(t=table):
                return (int(t.Rows.Count), int(t.Columns.Count))

            rows, cols = session.guard(_t)
            tables.append({"table": index, "rows": rows, "columns": cols})
        sections_raw = session.guard(lambda: raw_items(doc.Sections))
        sections = []
        for index, section in enumerate(sections_raw, start=1):
            def _s(s=section):
                ps = s.PageSetup
                return (float(ps.PageWidth), float(ps.PageHeight), int(ps.Orientation))

            w, h, o = session.guard(_s)
            sections.append(
                {"section": index, "page_width": w, "page_height": h, "orientation": o}
            )
        shapes_raw = session.guard(lambda: raw_items(doc.InlineShapes))
        images = []
        for index, shape in enumerate(shapes_raw, start=1):
            def _i(s=shape):
                return (int(s.Type), str(s.Range.Paragraphs(1).Range.Text.strip()))

            typ, cap = session.guard(_i)
            images.append({"inline_shape": index, "type": typ, "caption_paragraph": cap})
        return tables, sections, images

    def paragraph_index(self, offset: int) -> int:
        """The paragraph containing a character offset, or -1 before the first."""
        return bisect_right(self.paragraph_starts, offset) - 1

    def heading_path(self, offset: int) -> str:
        """The enclosing headings for a character offset, nearest last.

        A paragraph that is itself a heading does not name itself, so a heading
        match reports its ancestors only.
        """
        index = bisect_right(self.heading_starts, offset) - 1
        if index < 0:
            return ""
        heading = self.headings[index]
        if offset < heading.end:
            return heading.path
        return " > ".join(part for part in (heading.path, heading.text) if part)

    def patch(
        self, start: int, delta: int, fingerprint: int, paragraph_count: int
    ) -> bool:
        """Shift every stored offset a change moved. False when it does not fit.

        `start` is the first offset the edit touched, so offsets strictly greater
        than it move by `delta` and the paragraph containing it does not. Only
        called once the paragraph count has been confirmed unchanged.
        """
        shifted = [s + delta if s > start else s for s in self.paragraph_starts]
        if any(b < a for a, b in zip(shifted, shifted[1:])):
            return False
        self.paragraph_starts = shifted
        for heading in self.headings:
            if heading.start > start:
                heading.start += delta
            if heading.end > start:
                heading.end += delta
        self.heading_starts = [h.start for h in self.headings]
        self.fingerprint = fingerprint
        self.paragraph_count = paragraph_count
        return True

    def levels_match(self, session: Any, doc: Any, start: int, end: int) -> bool:
        """Whether a span's paragraphs are headings exactly as stored.

        A formatting change can move a paragraph in or out of the outline without
        changing a single character, which the fingerprint cannot see. Anything
        that cannot be read is treated as a mismatch, so the index is dropped.
        """
        stored = {self.paragraph_index(h.start): h.level for h in self.headings}
        first = max(self.paragraph_index(start), 0)
        last = min(self.paragraph_index(end), len(self.paragraph_starts) - 1)
        for i in range(first, last + 1):
            offset = self.paragraph_starts[i]
            try:
                para = session.guard(lambda o=offset: doc.Range(o, o).Paragraphs(1))
                level = int(PARAGRAPH.get(para, "OutlineLevel"))
            except Exception:
                return False
            if level != stored.get(i, WD_OUTLINE_BODY_TEXT):
                return False
        return True

    def refresh_span(self, session: Any, doc: Any, start: int, end: int) -> None:
        """Re-read the headings a change touched, then recompute every path."""
        touched = [h for h in self.headings if h.start <= end and h.end >= start]
        if not touched:
            return
        for heading in touched:
            rng = session.guard(lambda h=heading: doc.Range(h.start, h.end))
            heading.text = RANGE.get(rng, "Text").strip()
            heading.style = STYLE.get(RANGE.get(rng, "Style"), "NameLocal")
        self.repath()
        self.anchors = None  # the minted ranges belong to the old headings

    def repath(self) -> None:
        """Rebuild every heading's ancestor path from the level stack."""
        stack: list[Heading] = []
        for heading in self.headings:
            while stack and stack[-1].level >= heading.level:
                stack.pop()
            heading.path = " > ".join(h.text for h in stack)
            stack.append(heading)
