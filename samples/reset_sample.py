"""Close ONLY samples/plain.docx, so it can be regenerated.

Authorised one-off reset. Refuses unless exactly one open document has
Name == "plain.docx" and sits in a folder named "samples".
"""

import os
import sys

from word_mcp.session import ComWordSession

session = ComWordSession()
found = []
for app in session.applications():
    for doc in session._documents_of(app):
        name = str(doc.Name)
        full = str(doc.FullName)
        if (
            name == "plain.docx"
            and os.path.basename(full).lower() == "plain.docx"
            and os.path.basename(os.path.dirname(full)).lower() == "samples"
        ):
            found.append((doc, full))

if len(found) != 1:
    print(
        "REFUSING: expected exactly one samples/plain.docx, found",
        len(found),
        [full for _, full in found],
    )
    sys.exit(1)

doc, full = found[0]
print("closing only:", full)
doc.Close(SaveChanges=0)
print("closed")
