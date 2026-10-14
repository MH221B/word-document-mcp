"""The index's pure lookups. No Word needed."""

from word_mcp.docindex import Heading, StructureIndex


def make_index() -> StructureIndex:
    headings = [
        Heading(1, "Heading 1", "Top", 10, 20, ""),
        Heading(2, "Heading 2", "Sub", 30, 40, "Top"),
    ]
    return StructureIndex(
        fingerprint=1,
        paragraph_starts=[0, 10, 25, 30, 50],
        headings=headings,
        tables=[],
        sections=[],
        images=[],
    )


def test_paragraph_index_maps_offsets_to_paragraphs():
    index = make_index()
    assert index.paragraph_index(0) == 0
    assert index.paragraph_index(9) == 0
    assert index.paragraph_index(10) == 1  # exactly on a boundary
    assert index.paragraph_index(30) == 3
    assert index.paragraph_index(999) == 4


def test_heading_path_names_ancestors_only():
    index = make_index()
    assert index.heading_path(5) == ""  # before any heading
    assert index.heading_path(15) == ""  # inside the level-1 heading
    assert index.heading_path(35) == "Top"  # inside the level-2 heading
    assert index.heading_path(45) == "Top > Sub"  # body under the level-2 heading
