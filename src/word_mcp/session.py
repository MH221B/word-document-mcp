"""The live Word connection: attach, pin, and the guarded call wrapper."""

from __future__ import annotations

import time
from contextlib import contextmanager
from dataclasses import dataclass
from typing import Any, Callable, Protocol, TypeVar

import pythoncom
import win32com.client
from pythoncom import com_error

from word_mcp.docindex import StructureIndex
from word_mcp.errors import (
    DocumentAmbiguous,
    DocumentNotFound,
    InvalidArgument,
    ProtectedDocument,
    ReadOnlyDocument,
    WordMcpError,
    WordNotRunning,
    map_com_error,
)

T = TypeVar("T")
RETRY_SECONDS = 1.0
CALL_TIMEOUT_SECONDS = 10.0
LOOKUP_ATTEMPTS = 3
LOOKUP_RETRY_SECONDS = 0.5
# Word's automation server is gone, still starting, or not registered.
DEAD_SERVER_CODES = (-2147023174, -2146959355, -2147221021, -2147221164)


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
        self._doc: Any = None
        self._info: DocumentInfo | None = None
        self._index: StructureIndex | None = None

    # -- connection ---------------------------------------------------------

    def applications(self) -> list[Any]:
        """Every Word instance reachable on this machine, the user's one first.

        A machine runs several WINWORD processes, and GetActiveObject returns an
        arbitrary registered one, which is often a windowless background instance
        holding no documents. Each open document registers a moniker in the running
        object table, and its Application is the window the user is actually looking
        at, so the documents lead and GetActiveObject is the fallback.

        The table is a moving target while Word starts and stops, and a stale entry
        answers with "RPC server is unavailable", so the whole probe is retried
        before giving up.
        """
        pythoncom.CoInitialize()
        last_error: BaseException | None = None
        for _ in range(LOOKUP_ATTEMPTS):
            candidates = self._rot_applications()
            try:
                candidates.append(win32com.client.GetActiveObject("Word.Application"))
            except Exception as exc:
                last_error = exc
            found = self._dedupe(candidates)
            if found:
                # A visible window with documents open, then the most documents.
                found.sort(key=self._prominence, reverse=True)
                return found
            time.sleep(LOOKUP_RETRY_SECONDS)
        raise self._no_word(last_error)

    @staticmethod
    def _rot_applications() -> list[Any]:
        """Applications found through the open documents registered in the table.

        Binding an arbitrary entry can fail in more ways than COM errors, since the
        table is machine-wide, so every failure here is skipped rather than raised.
        """
        out: list[Any] = []
        try:
            context = pythoncom.CreateBindCtx(0)
            table = pythoncom.GetRunningObjectTable()
        except Exception:
            return out
        for moniker in table:
            try:
                name = moniker.GetDisplayName(context, None)
                if name.startswith("!"):
                    continue  # class monikers and cloud sync objects
                bound = moniker.BindToObject(context, None, pythoncom.IID_IDispatch)
                out.append(win32com.client.Dispatch(bound).Application)
            except Exception:
                continue
        return out

    @classmethod
    def _dedupe(cls, candidates: list[Any]) -> list[Any]:
        seen: set[tuple] = set()
        out: list[Any] = []
        for candidate in candidates:
            key = cls._signature(candidate)
            if key is None or key in seen:
                continue
            seen.add(key)
            out.append(candidate)
        return out

    @staticmethod
    def _no_word(last_error: BaseException | None) -> WordMcpError:
        hresult = getattr(last_error, "hresult", None)
        if hresult in DEAD_SERVER_CODES:
            return WordNotRunning(
                "Word is open but its automation server is not answering, which "
                "usually means it is still starting or shutting down. Retry in a "
                "moment."
            )
        return WordNotRunning(
            "Word is not running. Open Word and the target document first."
        )

    def app(self) -> Any:
        """The pinned document's application, or the most prominent one."""
        if self._doc is not None:
            return self._doc.Application
        return self.applications()[0]

    def documents(self) -> list[DocumentInfo]:
        infos: list[DocumentInfo] = []
        seen: set[str] = set()
        for app in self.applications():
            for doc in self._documents_of(app):
                info = self._info_for(doc)
                if info.path not in seen:
                    seen.add(info.path)
                    infos.append(info)
        return infos

    def pin(self, key: str) -> DocumentInfo:
        matches: list[tuple[Any, Any]] = []
        open_names: list[str] = []
        for app in self.applications():
            for doc in self._documents_of(app):
                open_names.append(self._info_for(doc).name)
                if self._same(doc, key):
                    matches.append((app, doc))
        if not matches:
            raise DocumentNotFound(
                f"No open document matches {key!r}.", open=open_names
            )
        if len(matches) > 1:
            raise DocumentAmbiguous(
                f"{key!r} matches more than one open document.",
                candidates=[self._info_for(doc).name for _, doc in matches],
            )
        self._doc = matches[0][1]
        self._info = self._info_for(self._doc)
        self._index = None
        return self._info

    def unpin(self) -> None:
        self._doc = None
        self._info = None
        self._index = None

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

    def index(self) -> StructureIndex:
        """The structure index for the pinned document.

        Rebuilt only when the document's character count changed, so a repeated
        read costs a single cheap property read. The build writes into a local
        and is assigned only when complete, so a failure publishes nothing.
        """
        doc = self.doc()
        fingerprint = int(self.guard(lambda: doc.Content.End))
        if self._index is None or self._index.fingerprint != fingerprint:
            self._index = StructureIndex.build(self, fingerprint)
        return self._index

    def invalidate_index(self) -> None:
        """Drop the cached index; the next index() rebuilds it."""
        self._index = None

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
    def _documents_of(app: Any) -> list[Any]:
        try:
            return list(app.Documents)
        except Exception:
            return []

    @staticmethod
    def _signature(app: Any) -> tuple | None:
        """What identifies an instance: its open documents and its windows."""
        try:
            docs = tuple(sorted(str(doc.FullName) for doc in app.Documents))
            windows = int(app.Windows.Count)
        except Exception:
            return None
        return (docs, windows)

    @staticmethod
    def _prominence(app: Any) -> tuple[int, int, int]:
        try:
            visible = 1 if app.Visible and app.Windows.Count > 0 else 0
            return (visible, len(list(app.Documents)), int(app.Windows.Count))
        except Exception:
            return (0, 0, 0)

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


def story_range(doc: Any, container: str):
    """The range a container's text lives in.

    Offsets inside a header or footer range are relative to that story, not to the
    body, so a range built with doc.Range() lands in the body. Every caller that
    needs a range inside a container goes through here.
    """
    if container == "body":
        return doc.Content
    if container == "header":
        return doc.Sections(1).Headers(1).Range
    if container == "footer":
        return doc.Sections(1).Footers(1).Range
    raise InvalidArgument(
        f"Container {container!r} is not supported yet.",
        supported=["body", "header", "footer"],
    )


@contextmanager
def quiet_batch(app: Any, session: "ComWordSession"):
    """Suppress alerts and screen updating for one batch, then restore them.

    A modal dialog cannot be dismissed by a background process, so leaving these
    set would hide the problem rather than fix it.
    """
    prev_alerts = app.DisplayAlerts
    prev_updating = app.ScreenUpdating
    session.guard(lambda: setattr(app, "DisplayAlerts", 0))  # wdAlertsNone
    session.guard(lambda: setattr(app, "ScreenUpdating", False))
    try:
        yield
    finally:
        session.guard(lambda: setattr(app, "DisplayAlerts", prev_alerts))
        session.guard(lambda: setattr(app, "ScreenUpdating", prev_updating))
