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


def test_patch_shifts_only_offsets_after_the_change():
    index = make_index()
    assert index.patch(start=10, delta=5, fingerprint=99, paragraph_count=5)
    # The paragraph containing the change keeps its start; later ones move.
    assert index.paragraph_starts == [0, 10, 30, 35, 55]
    # The change is at the level-1 heading's own start, so its start stays put
    # but its end grows; the level-2 heading starts after it, so both ends move.
    assert [(h.start, h.end) for h in index.headings] == [(10, 25), (35, 45)]
    assert index.heading_starts == [10, 35]
    assert index.fingerprint == 99


def test_patch_handles_a_deletion_at_the_boundary():
    index = make_index()
    assert index.patch(start=10, delta=-4, fingerprint=99, paragraph_count=5)
    assert index.paragraph_starts == [0, 10, 21, 26, 46]


def test_patch_refuses_a_shift_that_breaks_the_order():
    index = make_index()
    # A deletion larger than the gap to the previous paragraph cannot be right.
    assert not index.patch(start=10, delta=-20, fingerprint=99, paragraph_count=5)
    assert index.paragraph_starts == [0, 10, 25, 30, 50]


def test_repath_rebuilds_ancestors_from_the_levels():
    index = make_index()
    index.headings[0].text = "Renamed"
    index.repath()
    assert [h.path for h in index.headings] == ["", "Renamed"]
