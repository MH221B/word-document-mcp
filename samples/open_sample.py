"""Reopen the regenerated sample so the integration suite can pin it."""

from pathlib import Path

from word_mcp.session import ComWordSession

sample = Path(__file__).with_name("plain.docx")
session = ComWordSession()
app = session.applications()[0]
doc = app.Documents.Open(str(sample))
print("opened:", str(doc.Name), "|", str(doc.FullName))
