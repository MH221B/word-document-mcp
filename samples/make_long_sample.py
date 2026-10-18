"""Build samples/long.docx for cold outline timing.

Deterministic. Same seed gives same doc. Reuses the 1x1 png trick
from make_samples.py so no binary needs committing.
"""

import base64
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.shared import Inches

DOT_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII="
)

TARGET_PARAS = 1616
TARGET_TABLES = 42


def _caption(doc, text):
    try:
        return doc.add_paragraph(text, style="Caption")
    except Exception:
        # ponytail: plain style fallback, Caption style is template dependent
        return doc.add_paragraph(text)


def build(path="samples/long.docx"):
    Path("samples/dot.png").write_bytes(DOT_PNG)
    doc = Document()
    tables = 0
    paras = 0

    def add_body(i):
        doc.add_paragraph(f"Body {i} rider reference TOKEN123 filler text.")

    h1_count = 0
    i = 0
    bullets_done = False
    numbers_done = False
    section_added = False
    while paras < TARGET_PARAS or tables < TARGET_TABLES:
        if paras >= TARGET_PARAS and tables >= TARGET_TABLES:
            break
        # H1 every 100 paras, with H2 and H3 nested under it
        if paras % 100 == 0 and paras < TARGET_PARAS:
            h1_count += 1
            doc.add_heading(f"Part {h1_count} Overview", level=1)
            paras += 1
            doc.add_heading(f"Part {h1_count} Detail", level=2)
            paras += 1
            doc.add_heading(f"Part {h1_count} Notes", level=3)
            paras += 1
        # tables spread through, 42 total
        if tables < TARGET_TABLES and (paras % 38 == 0 or paras >= TARGET_PARAS):
            n = tables + 1
            if n == 1:
                _caption(doc, f"Table {n} Above example")
                paras += 1
            tbl = doc.add_table(rows=3, cols=4)
            for r in range(3):
                for c in range(4):
                    tbl.cell(r, c).text = f"R{r}C{c} T{n}"
            tables += 1
            paras += 1  # table paragraph
            if n == 2:
                _caption(doc, f"Table {n} Below example")
                paras += 1
            continue
        if paras >= TARGET_PARAS:
            continue
        # bullet block and numbered block once each
        if not bullets_done and paras >= 200:
            for k in range(5):
                doc.add_paragraph(f"Bullet {k}", style="List Bullet")
                paras += 1
            bullets_done = True
            continue
        if not numbers_done and paras >= 400:
            for k in range(5):
                doc.add_paragraph(f"Step {k}", style="List Number")
                paras += 1
            numbers_done = True
            continue
        # image every 400 paras
        if paras % 400 == 50 and paras < TARGET_PARAS:
            doc.add_picture("samples/dot.png", width=Inches(1.0))
            paras += 1
            continue
        # One section break, roughly halfway, so section 2 is real and
        # container targeting can be tested against it.
        if not section_added and paras >= TARGET_PARAS // 2:
            doc.add_section(WD_SECTION.NEW_PAGE)
            section_added = True
        add_body(i)
        paras += 1
        i += 1

    for number, section in enumerate(doc.sections, start=1):
        header = section.header
        header.is_linked_to_previous = False
        header.paragraphs[0].text = f"Section {number} header"

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"saved {path} paras={paras} tables={tables}")


if __name__ == "__main__":
    build()
