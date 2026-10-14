import pytest

from word_mcp.errors import InvalidArgument, WordMcpError
from word_mcp.history import ChangeStack


class FakeRange:
    def __init__(self, start, end):
        self.Start = start
        self.End = end


def test_records_newest_first_lookup_and_ids():
    stack = ChangeStack()
    first = stack.record("text", "insert", FakeRange(0, 5), text_before="")
    second = stack.record("text", "replace", FakeRange(5, 9), text_before="old")
    assert stack.newest().op_id == second.op_id
    assert stack.by_id(first.op_id).description == "insert"


def test_depth_is_bounded():
    stack = ChangeStack(depth=2)
    for i in range(5):
        stack.record("text", f"op{i}", FakeRange(i, i + 1))
    assert [c.description for c in stack._items] == ["op3", "op4"]


def test_unknown_id_is_rejected():
    stack = ChangeStack()
    stack.record("text", "insert", FakeRange(0, 1))
    with pytest.raises(InvalidArgument):
        stack.by_id(99)


def test_pop_on_empty_is_rejected():
    with pytest.raises(WordMcpError):
        ChangeStack().pop()


def test_only_newest_can_be_reverted():
    from word_mcp import history

    stack = ChangeStack()
    older = stack.record("text", "one", FakeRange(0, 1), text_before="a")
    stack.record("text", "two", FakeRange(2, 3), text_before="b")

    class Session:
        def doc(self):
            return None

        def guard(self, fn):
            return fn()

    with pytest.raises(InvalidArgument) as exc:
        history.revert(Session(), stack, older.op_id)
    assert "newest" in str(exc.value).lower()
