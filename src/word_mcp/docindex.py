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
    ) -> None:
        self.fingerprint = fingerprint
        self.paragraph_starts = paragraph_starts
        self.headings = headings
        self.heading_starts = [heading.start for heading in headings]
        self.tables = tables
        self.sections = sections
        self.images = images

    @classmethod
    def build(cls, session: Any, fingerprint: int) -> "StructureIndex":
        doc = session.doc()
        # One guard around the whole pass: a busy Word retries the read rather
        # than raising a raw COM error halfway through.
        starts, heading_rows = session.guard(lambda: cls._read_paragraphs(doc))
        tables, sections, images = cls._summaries(session, doc)
        return cls(
            fingerprint,
            starts,
            cls._headings(heading_rows),
            tables,
            sections,
            images,
        )

    @staticmethod
    def _read_paragraphs(doc: Any) -> tuple[list[int], list[tuple[int, int, Any]]]:
        """One pass: every paragraph's start offset, and the headings among them."""
        starts: list[int] = []
        heading_rows: list[tuple[int, int, Any]] = []
        for para in raw_items(doc.Paragraphs):
            rng = PARAGRAPH.get(para, "Range")
            start = RANGE.get(rng, "Start")
            level = int(PARAGRAPH.get(para, "OutlineLevel"))
            starts.append(start)
            if level < WD_OUTLINE_BODY_TEXT:
                heading_rows.append((len(starts) - 1, level, para))
        return starts, heading_rows

    @staticmethod
    def _headings(rows: list[tuple[int, int, Any]]) -> list[Heading]:
        headings: list[Heading] = []
        stack: list[Heading] = []
        for _, level, para in rows:
            rng = PARAGRAPH.get(para, "Range")
            while stack and stack[-1].level >= level:
                stack.pop()
            headings.append(
                Heading(
                    level=level,
                    style=STYLE.get(RANGE.get(rng, "Style"), "NameLocal"),
                    text=RANGE.get(rng, "Text").strip(),
                    start=RANGE.get(rng, "Start"),
                    end=RANGE.get(rng, "End"),
                    path=" > ".join(h.text for h in stack),
                )
            )
            stack.append(headings[-1])
        return headings

    @staticmethod
    def _summaries(session: Any, doc: Any) -> tuple[list[dict], list[dict], list[dict]]:
        tables = [
            {
                "table": index,
                "rows": int(session.guard(lambda t=table: t.Rows.Count)),
                "columns": int(session.guard(lambda t=table: t.Columns.Count)),
            }
            for index, table in enumerate(session.guard(lambda: list(doc.Tables)), start=1)
        ]
        sections = [
            {
                "section": index,
                "page_width": float(session.guard(lambda s=section: s.PageSetup.PageWidth)),
                "page_height": float(session.guard(lambda s=section: s.PageSetup.PageHeight)),
                "orientation": int(session.guard(lambda s=section: s.PageSetup.Orientation)),
            }
            for index, section in enumerate(session.guard(lambda: list(doc.Sections)), start=1)
        ]
        images = [
            {
                "inline_shape": index,
                "type": int(session.guard(lambda s=shape: s.Type)),
                "caption_paragraph": session.guard(
                    lambda s=shape: s.Range.Paragraphs(1).Range.Text.strip()
                ),
            }
            for index, shape in enumerate(
                session.guard(lambda: list(doc.InlineShapes)), start=1
            )
        ]
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
