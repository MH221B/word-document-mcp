"""Canned Variant rows and denial mapping. No Word needed."""

from word_mcp import macroscan
from word_mcp.errors import MacrosUnavailable, map_com_error


def test_rows_to_shapes_filters_body_and_strips_mark():
    rows = [
        (0, 10, 1, "Heading 1", "Top\r"),
        (10, 20, 10, "Normal", "Body\r"),
        (20, 30, 2, "Heading 2", "Sub\r"),
    ]
    starts, hrows, detail = macroscan.rows_to_shapes(rows)
    assert starts == [0, 10, 20]
    assert [(i, level) for i, level, _ in hrows] == [(0, 1), (2, 2)]
    assert detail[0] == ("Top", "Heading 1", 9)
    assert detail[2] == ("Sub", "Heading 2", 29)


def test_rows_to_shapes_clamps_empty_heading_end():
    starts, hrows, detail = macroscan.rows_to_shapes([(5, 5, 1, "Heading 1", "H\r")])
    assert starts == [5]
    assert detail[0] == ("H", "Heading 1", 5)


def test_vbom_denial_maps_to_macros_unavailable():
    err = map_com_error(
        Exception("Programmatic access to Office VBA project is denied (6068)")
    )
    assert isinstance(err, MacrosUnavailable)
    assert err.code == "macros_unavailable"


def test_disabled_macros_map_to_macros_unavailable():
    err = map_com_error(Exception("The macros in this project are disabled."))
    assert isinstance(err, MacrosUnavailable)
