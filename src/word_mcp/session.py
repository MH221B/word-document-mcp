"""The live Word connection: attach, pin, and the guarded call wrapper."""

from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Any, Callable, Protocol, TypeVar

import pythoncom
import win32com.client
from pythoncom import com_error

from word_mcp.errors import (
    DocumentAmbiguous,
    DocumentNotFound,
    ProtectedDocument,
    ReadOnlyDocument,
    WordMcpError,
    WordNotRunning,
    map_com_error,
)

T = TypeVar("T")
RETRY_SECONDS = 1.0
CALL_TIMEOUT_SECONDS = 10.0


@dataclass
class DocumentInfo:
    name: str
    path: str
    read_only: bool
    protected: bool


class WordSession(Protocol):
    """What operations need from a session. Implemented by ComWordSession and the
    in-memory fake used by unit tests."""

    def documents(self) -> list[DocumentInfo]: ...
    def pin(self, key: str) -> DocumentInfo: ...
    def unpin(self) -> None: ...
    def pinned(self) -> DocumentInfo | None: ...
    def app(self) -> Any: ...
    def doc(self) -> Any: ...
    def guard(self, fn: Callable[[], T]) -> T: ...


class ComWordSession:
    def __init__(self) -> None:
        self._app: Any = None
        self._doc: Any = None
        self._info: DocumentInfo | None = None

    # -- connection ---------------------------------------------------------

    def app(self) -> Any:
        if self._app is None:
            pythoncom.CoInitialize()
            try:
                self._app = win32com.client.GetActiveObject("Word.Application")
            except com_error as exc:
                if getattr(exc, "hresult", None) in (-2147221021, -2147221164):
                    raise WordNotRunning(
                        "Word is not running. Open Word and the target document first."
                    ) from exc
                raise map_com_error(exc) from exc
        return self._app

    def documents(self) -> list[DocumentInfo]:
        app = self.app()
        return [
            self._info_for(doc)
            for doc in self.guard(lambda: list(app.Documents))
        ]

    def pin(self, key: str) -> DocumentInfo:
        app = self.app()
        docs = self.guard(lambda: list(app.Documents))
        matches = [
            d for d in docs if self._same(d, key)
        ]
        if not matches:
            raise DocumentNotFound(
                f"No open document matches {key!r}.",
                open=[self._info_for(d).name for d in docs],
            )
        if len(matches) > 1:
            raise DocumentAmbiguous(
                f"{key!r} matches more than one open document.",
                candidates=[self._info_for(d).name for d in matches],
            )
        self._doc = matches[0]
        self._info = self._info_for(self._doc)
        return self._info

    def unpin(self) -> None:
        self._doc = None
        self._info = None

    def pinned(self) -> DocumentInfo | None:
        return self._info

    def doc(self) -> Any:
        if self._doc is None:
            raise DocumentNotFound("No document is pinned. Call connect_document.")
        info = self._info_for(self._doc)
        if info.path != self._info.path:
            self.unpin()
            raise DocumentNotFound(
                "The pinned document was renamed or saved elsewhere. "
                "Call connect_document again.",
                was=self._info.path,
                now=info.path,
            )
        return self._doc

    # -- guard --------------------------------------------------------------

    def guard(self, fn: Callable[[], T]) -> T:
        """Run a COM call, retrying while Word reports itself busy."""
        deadline = time.monotonic() + CALL_TIMEOUT_SECONDS
        while True:
            try:
                return fn()
            except com_error as exc:
                mapped = map_com_error(exc)
                if isinstance(mapped, WordMcpError) and mapped.code == "word_busy":
                    if time.monotonic() >= deadline:
                        raise WordMcpError(
                            "Word stayed busy for "
                            f"{CALL_TIMEOUT_SECONDS:.0f}s. Close any dialog in Word "
                            "and stop editing, then retry."
                        ) from exc
                    time.sleep(RETRY_SECONDS)
                    continue
                raise mapped from exc

    # -- helpers ------------------------------------------------------------

    @staticmethod
    def _same(doc: Any, key: str) -> bool:
        wanted = key.replace("\\", "/").casefold()
        return (
            doc.FullName.replace("\\", "/").casefold().endswith(wanted)
            or doc.Name.casefold() == key.casefold()
        )

    @staticmethod
    def _info_for(doc: Any) -> DocumentInfo:
        return DocumentInfo(
            name=doc.Name,
            path=doc.FullName,
            read_only=bool(doc.ReadOnly),
            protected=bool(doc.ProtectionType != -1),  # wdNoProtection == -1
        )


def require_writable(doc: Any) -> None:
    if doc.ReadOnly:
        raise ReadOnlyDocument("The document is open read-only.")
    if doc.ProtectionType != -1:
        raise ProtectedDocument(
            "The document has editing restrictions enabled.",
            protection_type=int(doc.ProtectionType),
        )
