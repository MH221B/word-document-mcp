"""Typed errors the tools return, and the COM to domain error mapping."""

RPC_E_CALL_REJECTED = 0x80010001 - 0x100000000  # -2147418111
RPC_E_DISCONNECTED = 0x80010108 - 0x100000000  # -2147417848


class WordMcpError(Exception):
    code = "word_error"

    def __init__(self, message: str, **details):
        super().__init__(message)
        self.message = message
        self.details = details

    def to_dict(self) -> dict:
        return {"code": self.code, "message": self.message, **self.details}


class WordNotRunning(WordMcpError):
    code = "word_not_running"


class DocumentNotFound(WordMcpError):
    code = "document_not_found"


class DocumentAmbiguous(WordMcpError):
    code = "document_ambiguous"


class AmbiguousTarget(WordMcpError):
    code = "ambiguous_target"


class TargetNotFound(WordMcpError):
    code = "target_not_found"


class StaleToken(WordMcpError):
    code = "stale_token"


class WordBusy(WordMcpError):
    code = "word_busy"


class WordDisconnected(WordMcpError):
    code = "word_disconnected"


class ProtectedDocument(WordMcpError):
    code = "protected_document"


class ReadOnlyDocument(WordMcpError):
    code = "read_only_document"


class ConfirmationRequired(WordMcpError):
    code = "confirmation_required"


class InvalidArgument(WordMcpError):
    code = "invalid_argument"


class MacrosUnavailable(WordMcpError):
    code = "macros_unavailable"


# Word error 6068: "Programmatic access to Visual Basic Project is not trusted."
# pywin32 surfaces this as DISP_E_EXCEPTION (-2147352567) with the Word code in
# excepinfo[5]; confirmed against live Word. The string list is the English
# fallback.
WORD_MACRO_DENIED_SCODE = 0x800A17B4 - 0x100000000

_MACRO_DENIED = (
    "6068",
    "programmatic access",
    "not trusted",
    "macros in this project are disabled",
    "cannot run the macro",
    "may be disabled",
)


def _is_macro_denial(exc: BaseException) -> bool:
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) > 5 and excepinfo[5] == WORD_MACRO_DENIED_SCODE:
        return True
    text = str(exc).casefold()
    return any(m in text for m in _MACRO_DENIED)


def map_com_error(exc: BaseException) -> WordMcpError:
    """Map a pywin32 com_error to a domain error."""
    hresult = getattr(exc, "hresult", None)
    if hresult is None:
        args = getattr(exc, "args", ())
        hresult = args[0] if args and isinstance(args[0], int) else None
    if hresult == RPC_E_CALL_REJECTED:
        return WordBusy("Word is busy; usually mid-edit or behind a dialog.")
    if hresult == RPC_E_DISCONNECTED:
        return WordDisconnected("Word closed or the document was closed.")
    if _is_macro_denial(exc):
        return MacrosUnavailable(
            "Word refused the macro path. Enable Trust access to the VBA "
            "project object model, or keep using the COM path."
        )
    return WordMcpError(f"Word returned an error: {exc}")
