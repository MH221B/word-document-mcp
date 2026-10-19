"""The live Word connection: attach, pin, and the guarded call wrapper."""

from __future__ import annotations

import sys
import time
from dataclasses import dataclass
from typing import Any, Callable, TypeVar

import pythoncom
import win32com.client
from pythoncom import com_error

from word_mcp.docindex import StructureIndex
from word_mcp.addressing import split_container
from word_mcp.errors import (
    DocumentAmbiguous,
    DocumentNotFound,
    InvalidArgument,
    MacrosUnavailable,
    ProtectedDocument,
    ReadOnlyDocument,
    WordBusy,
    WordDisconnected,
    WordMcpError,
    WordNotRunning,
    map_com_error,
)

T = TypeVar("T")
RETRY_SECONDS = 1.0
CALL_TIMEOUT_SECONDS = 10.0
LOOKUP_ATTEMPTS = 3
LOOKUP_RETRY_SECONDS = 0.5
MACRO_RETRY_SECONDS = 30.0
# Word's automation server is gone, still starting, or not registered.
DEAD_SERVER_CODES = (-2147023174, -2146959355, -2147221021, -2147221164)


@dataclass
class DocumentInfo:
    name: str
    path: str
    read_only: bool
    protected: bool


class ComWordSession:
    def __init__(self) -> None:
        self._doc: Any = None
        self._info: DocumentInfo | None = None
        self._index: StructureIndex | None = None
        self._inventory: dict | None = None
        self._end_before: int | None = None
        self._change: tuple[int, int] | None = None
        self._macro_ok: bool | None = None
        self._macro_denied = False
        self._macro_retry_after = 0.0
        self._macro_noted = False
        self._change_count = 0
        self._result_range: tuple | None = None
        self._effect = "structure"
        self._validated = False

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
        matches: list[tuple[Any, DocumentInfo]] = []
        open_names: list[str] = []
        for app in self.applications():
            for doc in self._documents_of(app):
                info = self._info_for(doc)
                open_names.append(info.name)
                if self._same(info, key):
                    matches.append((doc, info))
        if not matches:
            raise DocumentNotFound(
                f"No open document matches {key!r}.", open=open_names
            )
        if len(matches) > 1:
            raise DocumentAmbiguous(
                f"{key!r} matches more than one open document.",
                candidates=[info.name for _, info in matches],
            )
        self._doc = matches[0][0]
        self._info = matches[0][1]
        self._index = None
        self._inventory = None
        self._macro_ok = None
        self._macro_retry_after = 0.0
        return self._info

    def unpin(self) -> None:
        self._doc = None
        self._info = None
        self._index = None
        self._inventory = None
        self._validated = False

    def pinned(self) -> DocumentInfo | None:
        return self._info

    def begin_call(self) -> None:
        """Called once per tool call, so the next doc() re-validates the pin.

        A single tool call reaches doc() several times; re-reading FullName each
        time is a round-trip per call. No COM here.
        """
        self._validated = False

    def doc(self) -> Any:
        if self._doc is None:
            raise DocumentNotFound("No document is pinned. Call connect_document.")
        if not self._validated:
            self._validated = True
            path = str(self.guard(lambda: self._doc.FullName))
            if path != self._info.path:
                was = self._info.path
                self.unpin()
                raise DocumentNotFound(
                    "The pinned document was renamed or saved elsewhere. "
                    "Call connect_document again.",
                    was=was,
                    now=path,
                )
        return self._doc

    def require_writable(self) -> None:
        """Raise unless the pinned document can be edited.

        Uses the info captured at pin time. A document does not become read-only
        or protected mid-session, and Word rejects the edit itself if it did.
        """
        info = self._info
        if info is None:
            raise DocumentNotFound("No document is pinned. Call connect_document.")
        if info.read_only:
            raise ReadOnlyDocument("The document is open read-only.")
        if info.protected:
            raise ProtectedDocument(
                "The document has editing restrictions enabled.",
                protection_type=int(self.guard(lambda: self._doc.ProtectionType)),
            )

    def index(self) -> StructureIndex:
        """The structure index for the pinned document.

        Rebuilt only when the document's character count changed, so a repeated
        read costs a single cheap property read. The build writes into a local
        and is assigned only when complete, so a failure publishes nothing.
        """
        doc = self.doc()
        fingerprint = int(self.guard(lambda: doc.Content.End))
        if self._index is None or self._index.fingerprint != fingerprint:
            count = int(self.guard(lambda: doc.Paragraphs.Count))
            self._index = StructureIndex.build(self, fingerprint, count)
        return self._index

    def invalidate_index(self) -> None:
        """Drop the cached index and inventory; the next read rebuilds them."""
        self._index = None
        self._inventory = None

    def peek_index(self) -> StructureIndex | None:
        """The cached index if it still matches the document. Never builds."""
        index = self._index
        if index is None:
            return None
        doc = self.doc()
        if int(self.guard(lambda: doc.Content.End)) != index.fingerprint:
            self.invalidate_index()
            return None
        return index

    def inventory(self) -> dict | None:
        """The cached document inventory, or None when it must be re-read."""
        return self._inventory

    def set_inventory(self, payload: dict | None) -> None:
        """Cache the inventory, or clear it with None."""
        self._inventory = payload

    def macro_available(self) -> bool:
        """Whether the batched VBA scan may run. Probed once, then cached.

        A denial caches False. Any other failure is treated as transient: it is
        not cached, and a cooldown keeps a broken probe from running on every
        build.
        """
        if self._macro_ok is not None:
            return self._macro_ok
        now = time.monotonic()
        if now < self._macro_retry_after:
            return False
        try:
            from word_mcp import macroscan

            self._macro_ok = macroscan.probe(self)
            self._macro_denied = False
        except MacrosUnavailable:
            self._macro_ok = False
            self._macro_denied = True
            self._note_macro_fallback()
        except Exception:
            self._macro_retry_after = now + MACRO_RETRY_SECONDS
            return False
        return self._macro_ok

    def drop_macro(self) -> None:
        """Disable the macro path after a confirmed denial."""
        self._macro_ok = False
        self._macro_denied = True
        self._note_macro_fallback()

    def macro_denied(self) -> bool:
        """Whether the macro path was last classified as denied (not transient)."""
        return self._macro_denied

    def _note_macro_fallback(self) -> None:
        if self._macro_noted:
            return
        self._macro_noted = True
        print(
            "word-mcp: VBA trust unavailable; using the slower COM scan.",
            file=sys.stderr,
        )

    def begin_change(self, effect: str = "structure") -> None:
        """Snapshot the length, so a mutation can be described to the index.

        `effect` is what the mutation can touch: "structure" (text or structure
        moved), "inventory" (only comments/bookmarks changed), or "none" (nothing
        observable changed). Never raises: a mutation that cannot be described
        simply drops the index, which is what happened before this existed.
        """
        self._effect = effect
        self._change = None
        self._change_count = 0
        self._end_before = None
        self._result_range = None
        self._container = "body"
        if self._index is None:
            return
        try:
            self._end_before = int(self.guard(lambda: self.doc().Content.End))
        except Exception:
            self._end_before = None

    def note_change(self, start: int, end: int, container: str = "body") -> None:
        """Record the span a mutating operation changed, in post-edit offsets.

        `container` says which story those offsets belong to. A header edit's
        offsets are header-relative, so without this they would look like a body
        edit and the body index would be thrown away for no reason.
        """
        if self._end_before is None and self._index is not None:
            # The index was built during resolve, before this mutation, so its
            # fingerprint is exactly the pre-edit Content.End.
            self._end_before = self._index.fingerprint
        self._change = (start, end)
        self._change_count += 1
        self._container = container

    def note_result_range(self, handle: Any, container: str) -> None:
        """Stash the post-edit range so `mutate` can build `at` after settling."""
        self._result_range = (handle, container)

    def take_result_range(self) -> tuple | None:
        out = self._result_range
        self._result_range = None
        return out

    def has_index(self) -> bool:
        return self._index is not None

    def settle_index(self) -> None:
        """Patch the cached index if the recorded change verifies, else drop it.

        A wrong index is worse than a slow one, so nothing here is trusted: the
        index must have been fresh, the paragraph count must be unchanged, and the
        touched paragraphs must still be headings exactly as stored. Any failure,
        including a read that raises, drops the index and costs a rebuild.
        """
        try:
            self._settle()
        except (WordDisconnected, WordNotRunning):
            self.invalidate_index()
            raise
        except Exception:
            self.invalidate_index()

    def _settle(self) -> None:
        index = self._index
        if index is None:
            return
        if self._effect in ("none", "inventory") or self._container != "body":
            # The operation promised no body character moved, or it happened in a
            # header or footer, whose offsets are a different space entirely. One
            # read confirms the body is untouched.
            doc = self.doc()
            if int(self.guard(lambda: doc.Content.End)) == index.fingerprint:
                if self._effect == "inventory":
                    self._inventory = None
                return
            self.invalidate_index()
            return
        if self._change is None or self._end_before is None or self._change_count != 1:
            self.invalidate_index()
            return
        start, end = self._change
        doc = self.doc()
        end_now = int(self.guard(lambda: doc.Content.End))
        if self._end_before != index.fingerprint:
            self.invalidate_index()  # the index was already stale before the edit
            return
        count = int(self.guard(lambda: doc.Paragraphs.Count))
        if count != index.paragraph_count:
            self.invalidate_index()
            return
        # Patch first: it shifts the stored offsets into post-edit space, which is
        # the space levels_match reads the document in. Any check below that fails
        # drops the index anyway, so shifting before verifying costs nothing.
        if not index.patch(start, end_now - self._end_before, end_now, count):
            self.invalidate_index()
            return
        if not index.levels_match(self, doc, start, end):
            self.invalidate_index()
            return
        index.refresh_span(self, doc, start, end)

    # -- guard --------------------------------------------------------------

    def guard(self, fn: Callable[[], T]) -> T:
        """Run a COM call, retrying while Word reports itself busy."""
        deadline: float | None = None
        while True:
            try:
                return fn()
            except com_error as exc:
                mapped = map_com_error(exc)
                if isinstance(mapped, WordMcpError) and mapped.code == "word_busy":
                    if deadline is None:
                        deadline = time.monotonic() + CALL_TIMEOUT_SECONDS
                    if time.monotonic() >= deadline:
                        raise WordBusy(
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
    def _same(info: DocumentInfo, key: str) -> bool:
        wanted = key.replace("\\", "/").casefold()
        return (
            info.path.replace("\\", "/").casefold().endswith(wanted)
            or info.name.casefold() == key.casefold()
        )

    @staticmethod
    def _info_for(doc: Any) -> DocumentInfo:
        return DocumentInfo(
            name=doc.Name,
            path=doc.FullName,
            read_only=bool(doc.ReadOnly),
            protected=bool(doc.ProtectionType != -1),  # wdNoProtection == -1
        )


def story_range(doc: Any, container: str):
    """The range a container's text lives in.

    Offsets inside a header or footer range are relative to that story, not to the
    body, so a range built with doc.Range() lands in the body. Every caller that
    needs a range inside a container goes through here. A container may name its
    section, as in "header:2"; a bare name means section 1.
    """
    name, section = split_container(container)
    if name == "body":
        return doc.Content
    if name == "header":
        return doc.Sections(section).Headers(1).Range
    if name == "footer":
        return doc.Sections(section).Footers(1).Range
    raise InvalidArgument(
        f"Container {container!r} is not supported yet.",
        supported=["body", "header", "footer"],
    )
