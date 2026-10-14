"""In-memory WordSession for unit tests."""

from dataclasses import dataclass

from word_mcp.addressing import Match, TargetSpec, choose, normalise
from word_mcp.session import DocumentInfo


@dataclass
class FakePara:
    text: str
    style: str = "Normal"
    outline_level: int = 10  # 10 == body text


class FakeHandle:
    def __init__(self, index: int):
        self.index = index
        self.Start = index
        self.End = index + 1


class FakeSession:
    def __init__(self, paragraphs: list[FakePara], name: str = "fake.docx"):
        self._paras = paragraphs
        self._info = DocumentInfo(name=name, path=f"C:/docs/{name}", read_only=False, protected=False)

    def documents(self):
        return [self._info]

    def pin(self, key):
        return self._info

    def unpin(self):
        pass

    def pinned(self):
        return self._info

    def app(self):
        raise NotImplementedError("fake session has no Word application")

    def doc(self):
        raise NotImplementedError("fake session has no document object")

    def guard(self, fn):
        return fn()

    def heading_path_at(self, index: int) -> str:
        parts = []
        level_cap = 10
        for i in range(index - 1, -1, -1):
            para = self._paras[i]
            if para.outline_level < level_cap:
                level_cap = para.outline_level
                parts.insert(0, para.text)
                if level_cap == 1:
                    break
        return " > ".join(parts)

    def locate(self, spec: TargetSpec) -> list[Match]:
        out: list[Match] = []
        wanted = normalise(spec.text) if spec.text else None
        for i, para in enumerate(self._paras):
            if spec.style and para.style != spec.style:
                continue
            if wanted is not None and wanted.casefold() not in normalise(para.text).casefold():
                continue
            path = self.heading_path_at(i)
            if spec.heading_path and spec.heading_path not in path:
                continue
            out.append(
                Match(
                    token="",
                    text=para.text,
                    paragraph_index=i,
                    heading_path=path,
                    style=para.style,
                    container="body",
                    snippet=para.text[:80],
                    handle=FakeHandle(i),
                )
            )
        if spec.occurrence is not None:
            return choose(spec.text or "", out, spec.occurrence)
        return out
