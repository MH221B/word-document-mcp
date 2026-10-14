from word_mcp import errors


class FakeComError(Exception):
    def __init__(self, hresult):
        super().__init__("com")
        self.hresult = hresult


def test_call_rejected_maps_to_busy():
    assert isinstance(
        errors.map_com_error(FakeComError(errors.RPC_E_CALL_REJECTED)), errors.WordBusy
    )


def test_disconnected_maps_to_disconnected():
    mapped = errors.map_com_error(FakeComError(errors.RPC_E_DISCONNECTED))
    assert isinstance(mapped, errors.WordDisconnected)


def test_unknown_hresult_stays_generic():
    mapped = errors.map_com_error(FakeComError(-2147024891))
    assert type(mapped) is errors.WordMcpError


def test_to_dict_includes_code_and_details():
    payload = errors.AmbiguousTarget("many", candidates=[1, 2]).to_dict()
    assert payload["code"] == "ambiguous_target"
    assert payload["candidates"] == [1, 2]
