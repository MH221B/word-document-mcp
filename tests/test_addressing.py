import pytest

from word_mcp import addressing
from word_mcp.addressing import TargetSpec, TokenStore, normalise
from word_mcp.errors import AmbiguousTarget, StaleToken, TargetNotFound
from tests.fake_session import FakePara, FakeSession


def test_normalise_collapses_whitespace_and_quotes():
    assert normalise('He said \u201chi\nthere\u201d') == 'He said "hi there"'


def test_smart_quotes_match_straight_quotes():
    session = FakeSession([FakePara('The "Risks" section')])
    found = session.locate(TargetSpec(text='The "Risks" section'))
    assert len(found) == 1


def test_multiple_matches_return_candidates_not_a_guess():
    session = FakeSession([FakePara("Total 5"), FakePara("Total 7")])
    with pytest.raises(AmbiguousTarget) as exc:
        addressing.choose("Total", session.locate(TargetSpec(text="Total")), None)
    assert len(exc.value.details["candidates"]) == 2


def test_occurrence_selects_one():
    session = FakeSession([FakePara("Total 5"), FakePara("Total 7")])
    picked = addressing.choose("Total", session.locate(TargetSpec(text="Total")), 1)
    assert picked[0].text == "Total 7"


def test_no_match_reports_not_found():
    session = FakeSession([FakePara("Nothing here")])
    with pytest.raises(TargetNotFound):
        addressing.choose("Total", session.locate(TargetSpec(text="Total")), None)


def test_heading_path_narrows_matches():
    session = FakeSession(
        [
            FakePara("Overview", outline_level=1),
            FakePara("Total 5"),
            FakePara("Risks", outline_level=1),
            FakePara("Total 9"),
        ]
    )
    found = session.locate(TargetSpec(text="Total", heading_path="Risks"))
    assert [m.text for m in found] == ["Total 9"]


def test_token_round_trip_and_unknown_token():
    store = TokenStore()
    session = FakeSession([FakePara("Total 5")])
    match = session.locate(TargetSpec(text="Total"))[0]
    token = store.add(match)
    assert store.get(token).text == "Total 5"
    with pytest.raises(StaleToken):
        store.get("nope")


def test_token_with_deleted_range_is_stale():
    store = TokenStore()
    session = FakeSession([FakePara("Total 5")])
    match = session.locate(TargetSpec(text="Total"))[0]
    token = store.add(match)
    match.handle.Start, match.handle.End = 5, 1
    with pytest.raises(StaleToken):
        store.get(token)


def test_token_store_evicts_the_oldest():
    from word_mcp.addressing import Match, TokenStore

    store = TokenStore(limit=3)
    for i in range(5):
        store.add(
            Match(
                token=f"t{i}", text=str(i), paragraph_index=i, heading_path="",
                style="", container="body", snippet="", handle=None,
            )
        )
    assert sorted(store._items) == ["t2", "t3", "t4"]
