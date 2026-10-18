"""Focused checks for the fast-path fixes. No Word required."""

import pytest

from word_mcp import errors
from word_mcp.comfast import raw_items
from word_mcp.docindex import StructureIndex
from word_mcp.errors import MacrosUnavailable


def test_denial_text_variants_map_to_macros_unavailable():
    assert errors.map_com_error(
        Exception("Cannot run the macro 'X'. The macro may not be available.")
    ).code == "macros_unavailable"
    assert errors.map_com_error(
        Exception("Programmatic access to Visual Basic Project is not trusted.")
    ).code == "macros_unavailable"


def test_denial_word_scode_maps_to_macros_unavailable():
    class _Err(Exception):
        # What pywin32 actually raises for a VBOM denial: DISP_E_EXCEPTION with
        # the Word code 6068 in excepinfo[5], and a localized description.
        excepinfo = (0, "Microsoft Word", "Zugriff verweigert.", None, 0, -2146822220)

    assert errors.map_com_error(_Err("Exception occurred.")).code == "macros_unavailable"


class _StubSession:
    """Just enough session for StructureIndex.build."""

    def __init__(self, macro):
        self._macro = macro
        self.dropped = False

    def macro_available(self):
        return self._macro

    def drop_macro(self):
        self.dropped = True

    def guard(self, fn):
        return fn()


def _index(source="com"):
    return StructureIndex(1, [0], [], [], [], [], 1, source=source)


def test_build_uses_vba_and_does_not_drop(monkeypatch):
    s = _StubSession(macro=True)
    monkeypatch.setattr(StructureIndex, "build_macro", classmethod(lambda cls, *a: _index("vba")))
    assert StructureIndex.build(s, 1, 1).source == "vba"
    assert s.dropped is False


def test_build_drops_only_on_denial(monkeypatch):
    def deny(cls, *a):
        raise MacrosUnavailable("nope")

    s = _StubSession(macro=True)
    monkeypatch.setattr(StructureIndex, "build_macro", classmethod(deny))
    monkeypatch.setattr(StructureIndex, "build_com", classmethod(lambda cls, *a: _index("com")))
    assert StructureIndex.build(s, 1, 1).source == "com"
    assert s.dropped is True


def test_build_falls_back_on_scan_fault_without_dropping(monkeypatch):
    def fault(cls, *a):
        raise RuntimeError("ragged array")

    s = _StubSession(macro=True)
    monkeypatch.setattr(StructureIndex, "build_macro", classmethod(fault))
    monkeypatch.setattr(StructureIndex, "build_com", classmethod(lambda cls, *a: _index("com")))
    assert StructureIndex.build(s, 1, 1).source == "com"
    assert s.dropped is False


def test_raw_items_requests_batches():
    seen = []

    class _Iter:
        def Next(self, celt):
            seen.append(celt)
            batches = [(object(), object()), (object(),), ()]
            return batches[len(seen) - 1]

    class _Enum:
        def QueryInterface(self, iid):
            return _Iter()

    class _Ole:
        def Invoke(self, *a):
            return _Enum()

    class _Collection:
        _oleobj_ = _Ole()

    items = raw_items(_Collection(), celt=2)
    assert len(items) == 3
    assert seen == [2, 2]


def test_manage_style_list_reuses_inventory(monkeypatch):
    from word_mcp import inspect
    from word_mcp.operations import styles as styles_ops

    class _Doc:
        ReadOnly = False
        ProtectionType = -1

    class _Session:
        def doc(self):
            return _Doc()

    monkeypatch.setattr(
        inspect, "inspect_document",
        lambda session, refresh=False: {
            "styles": [{"name": "Normal", "built_in": True, "type": 1, "in_use": True}]
        },
    )
    result = styles_ops.manage_style(_Session(), None, None, action="list")
    assert result == {"styles": [{"name": "Normal", "built_in": True, "in_use": True}]}


def test_registered_tools_match_the_documented_set():
    from word_mcp.server import mcp

    names = {t.name for t in mcp._tool_manager.list_tools()}
    expected = {
        "list_open_documents", "connect_document", "disconnect_document",
        "document_outline", "read_range", "get_formatting", "find_text",
        "inspect_document", "get_ooxml", "edit_text", "format_text", "apply_edits",
        "revert_change", "save_document", "set_document_properties", "manage_style",
        "manage_list", "edit_table", "edit_page_layout", "edit_fields", "edit_media",
        "edit_links", "edit_content_controls", "edit_annotations", "insert_ooxml",
    }
    assert names == expected


def test_skill_md_lists_every_tool():
    from pathlib import Path

    from word_mcp.server import mcp

    skill = Path("skills/word-document-editing/SKILL.md").read_text(encoding="utf-8")
    missing = [t.name for t in mcp._tool_manager.list_tools() if t.name not in skill]
    assert missing == []


def test_target_spec_rejects_unknown_keys_and_container():
    from word_mcp.addressing import TargetSpec
    from word_mcp.errors import InvalidArgument

    assert TargetSpec.from_dict({"text": "x"}).text == "x"
    with pytest.raises(InvalidArgument):
        TargetSpec.from_dict({"text": "x", "bogus": 1})
    with pytest.raises(InvalidArgument):
        TargetSpec.from_dict({"text": "x", "container": "nope"})


def test_change_stack_clear():
    from word_mcp.history import ChangeStack

    stack = ChangeStack()
    stack._items.append(object())
    stack.clear()
    assert stack.newest() is None


def test_guard_retries_while_busy(monkeypatch):
    from pythoncom import com_error

    from word_mcp import session as session_mod

    monkeypatch.setattr(session_mod, "RETRY_SECONDS", 0)
    session = session_mod.ComWordSession()
    calls = {"n": 0}

    def flaky():
        calls["n"] += 1
        if calls["n"] < 2:
            raise com_error(-2147418111, "busy", None, None)
        return "ok"

    assert session.guard(flaky) == "ok"
    assert calls["n"] == 2


def test_guard_times_out_when_always_busy(monkeypatch):
    from pythoncom import com_error

    from word_mcp import session as session_mod
    from word_mcp.errors import WordBusy

    monkeypatch.setattr(session_mod, "RETRY_SECONDS", 0)
    monkeypatch.setattr(session_mod, "CALL_TIMEOUT_SECONDS", 0)
    session = session_mod.ComWordSession()

    def always_busy():
        raise com_error(-2147418111, "busy", None, None)

    with pytest.raises(WordBusy):
        session.guard(always_busy)


def test_macro_denial_is_cached(monkeypatch):
    from word_mcp import macroscan
    from word_mcp import session as session_mod
    from word_mcp.errors import MacrosUnavailable

    session = session_mod.ComWordSession()
    calls = {"n": 0}

    def deny(_session):
        calls["n"] += 1
        raise MacrosUnavailable("no")

    monkeypatch.setattr(macroscan, "probe", deny)
    assert session.macro_available() is False
    assert session.macro_available() is False
    assert calls["n"] == 1
    assert session.macro_denied() is True
