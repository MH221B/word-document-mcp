---
name: word-document-editing
description: 'Read and edit a Word document that is currently open in Word desktop on Windows, through the local word-mcp MCP server (tools list_open_documents, connect_document, disconnect_document, document_outline, read_range, find_text, get_formatting, inspect_document, get_ooxml, edit_text, apply_edits, format_text, edit_table, edit_page_layout, edit_fields, edit_media, edit_links, edit_content_controls, edit_annotations, manage_style, manage_list, insert_ooxml, revert_change, save_document, set_document_properties). Use when a task involves changing or inspecting a live Word document: inserting or replacing text, applying or auditing styles, tables, captions, headers, footers, page setup, fields, tables of contents, images, comments, tracked changes, or content controls in an open .docx, or when the user refers to word-mcp, "the document I have open", or wants edits to appear in Word right now.'
---

# Editing a live Word document

`word-mcp` edits the document the user has open in Word. It never opens or
saves-as a document, so every operation acts on the live document in memory, and
only the pinned document is reachable.

## Quick start

```
list_open_documents()                      # what Word has open
connect_document("report.docx")             # pin one for the session
document_outline()                          # headings, tables, sections, images, each with a token
edit_text({"text": "Total riders"}, "replace", text="Total riders: 9")   # returns an "at" token for the result
```

## The loop

Name a target in this order: a token from a read, then text with context, then a
style, then a heading path. Add `occurrence` only after seeing the candidates.

Read first when you need candidates or formatting. For a single, immediately
applied edit, targeting by text directly is fine: the first text-targeted edit is
the probe, and if you get `ambiguous_target`, call `find_text` and narrow with
`heading_path`, `occurrence`, or a candidate token. Every read returns a token
per location, and the server holds the resolved Word range behind it, so an edit
lands where you read even if the document shifted in between. Every text edit
returns an `at` token for the result, so edits chain without a verification read.
Use `apply_edits` for several changes in one call.

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
breaks, and headers and footers by `side` and `which`. `header_footer` writes
literal text. Header and footer offsets belong to their own story, so a target
with `container="footer"` stays in the footer. Name a section after a colon,
`container="footer:2"`; a bare name means section 1. `which` is not reachable
from a container, so `container` always means the primary header or footer.

Putting a page number in a footer takes two steps, because inserting a field
replaces the text it was aimed at. Insert the field over a placeholder, then put
the literal text back in front of it:

```
edit_page_layout("header_footer", side="footer", which="primary", text="NN")
res = edit_fields("insert", target={"text": "NN", "container": "footer"}, field="page")
edit_text(res["at"]["token"], "insert", text="Page ", position="start")
```

The footer then reads "Page 1", with the number still a live field.

**Comments and revisions.** `edit_annotations` adds, replies to, and deletes
comments, and toggles track changes. Marking a comment done (`resolve_comment`)
is not available on this Word build. Agent edits are not tracked unless you call
`track_changes` with `mode="on"` first, so do that before a review pass when the
user wants to accept or reject your work.

**Inventory, metadata, and raw XML.** `inspect_document` lists styles, comments,
revisions, bookmarks, and content controls in one cached call. Use it for a
formatting or review audit instead of enumerating collections yourself.
`set_document_properties` sets built-in (Title, Author, ...) or custom document
properties. `get_ooxml` returns WordprocessingML for a target when the object
model has no property for what you need. `save_document` saves the pinned
document. `disconnect_document` forgets the pin when you are done with it.

**Undo a mistake.** `revert_change(operation_id)` undoes the newest change only,
in reverse order. Only `edit_text`, `format_text`, `manage_style("apply")`, and
`insert_ooxml` return an operation id; the other mutations are not recorded and
cannot be reverted.

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
- Reads come from a structure index cached per document version, so a second
  `document_outline` on an unchanged document is nearly free. A normal edit made
  through a tool keeps the structure warm, because the edit patches the index in
  place. A multi-change `apply_edits` drops the index instead, so the next read is
  a full rebuild. A change made directly in Word, rather than through a tool, may
  need `document_outline(refresh=true)` before the new headings appear.

## Reference

Tool catalogue with parameters, error codes, and recipes for watermarks, tables of
contents, cross-references, numbering, and formatting audits: [REFERENCE.md](REFERENCE.md).
