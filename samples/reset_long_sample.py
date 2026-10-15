"""Close ONLY samples/long.docx, so it can be regenerated.

Mirrors reset_sample.py. Refuses unless exactly one open doc
is named long.docx inside a folder named samples.
Takes --reopen to open it again in the first Word instance.
"""

import os
import sys

from word_mcp.session import ComWordSession

reopen = "--reopen" in sys.argv

session = ComWordSession()
found = []
for app in session.applications():
    for doc in session._documents_of(app):
        name = str(doc.Name)
        full = str(doc.FullName)
        if (
            name == "long.docx"
            and os.path.basename(full).lower() == "long.docx"
            and os.path.basename(os.path.dirname(full)).lower() == "samples"
        ):
            found.append((doc, full))

if len(found) != 1:
    print(
        "REFUSING: expected exactly one samples/long.docx, found",
        len(found),
        [full for _, full in found],
    )
    sys.exit(1)

doc, full = found[0]
print("closing only:", full)
doc.Close(SaveChanges=0)
print("closed")

if reopen:
    app = session.applications()[0]
    opened = app.Documents.Open(os.path.join(os.getcwd(), "samples", "long.docx"))
    print("reopened:", str(opened.FullName))
