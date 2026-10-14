---
name: word-document-editing
description: Read and edit a Word document that is currently open in Word desktop on Windows, through the local word-mcp MCP server (tools list_open_documents, connect_document, document_outline, read_range, find_text, get_formatting, edit_text, format_text, edit_table, edit_page_layout, edit_fields, edit_media, edit_links, edit_content_controls, edit_annotations, manage_style, manage_list, insert_ooxml, revert_change). Use when a task involves changing or inspecting a live Word document: inserting or replacing text, applying or auditing styles, tables, captions, headers, footers, page setup, fields, tables of contents, images, comments, tracked changes, or content controls in an open .docx, or when the user refers to word-mcp, "the document I have open", or wants edits to appear in Word right now.
---

# Editing a live Word document

`word-mcp` edits the document the user has open in Word. It has no file access, so
every operation acts on the live document in memory, and only the pinned document
is reachable.

## Quick start

```
list_open_documents()                      # what Word has open
connect_document("report.docx")             # pin one for the session
document_outline()                          # headings, tables, sections, images, each with a token
read_range({"text": "Total riders"})        # text, plus a fresh token
edit_text({"token": "be4b28bb"}, "replace", text="Total riders: 9")
```

## The loop

Read first, then edit by token. Every read returns a token per location, and the
server holds the resolved Word range behind it, so an edit lands where you read
even if the document shifted in between.

Name a target in this order: a token from a read, then text with context, then a
style, then a heading path. Add `occurrence` only after seeing the candidates.

## Workflows

**Change existing text.** `read_range` for the token, then `edit_text` with that
token. Actions are `insert` (with `position` of `start`, `end`, or `replace`),
`replace`, and `delete`. Prefer `replace` over delete then insert.

**Add content to a templated document.** `format_mode` decides the formatting.
`inherit` is the default and behaves as if the user typed at that point, direct
formatting included. `style_only` applies the style and clears direct formatting,
which is what you want when the template's style definition should govern.
`explicit` applies exactly the `formatting` dict you pass. Run `get_formatting`
first when it matters whether a paragraph has drifted: it reports
`deviations_from_style`.

**Tables and figures.** `edit_table("insert", target={...}, rows=, columns=,
values=)` creates one, then address it by `table_index`. `edit_table("caption", ...)`
and `edit_media("caption", image_index=1, ...)` insert real captions, which are SEQ
fields, so numbers renumber themselves and `edit_fields("table_of_figures")` can
collect them.

**Page furniture.** `edit_page_layout` handles page setup, columns, section
breaks, and headers and footers by `side` and `which`. Put a page number in a
footer by targeting the footer text and calling `edit_fields("insert",
field="page")`.

**Comments and revisions.** `edit_annotations` adds, replies to, and resolves
comments, and toggles track changes. Agent edits are not tracked unless you call
`track_changes` with `mode="on"` first, so do that before a review pass when the
user wants to accept or reject your work.

**Undo a mistake.** `revert_change(operation_id)` undoes the newest change only,
in reverse order. Each mutation returns its operation id.

## Rules that keep you out of trouble

- One document is pinned per session, and the pin is checked before every
  operation. Call `connect_document` again to switch documents.
- Ambiguity is never guessed. `ambiguous_target` returns the candidates; narrow by
  `heading_path`, `container`, `occurrence`, or a candidate token.
- `confirm=true` is required for deleting more than one paragraph or 500
  characters, for ranges over 2,000 characters, and for accepting or rejecting all
  revisions. Ask the user before passing it.
- Mutations fail with `read_only_document` or `protected_document` on a document
  that cannot be edited. Report that to the user.
- `word_busy` means Word is mid-edit or behind a dialog. The server already retries
  for ten seconds, so wait for the user rather than looping.
- `word_disconnected` means Word closed. Ask the user to reopen the document, then
  connect again.

## Reference

Tool catalogue with parameters, error codes, and recipes for watermarks, tables of
contents, cross-references, numbering, and formatting audits: [REFERENCE.md](REFERENCE.md).
