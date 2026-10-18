# word-mcp

An MCP server that reads and edits the document you have open in Word desktop on
Windows.

It does not open documents and it does not save copies. It attaches to a document
that is already open, so your edits appear in Word as you make them. Word keeps
character offsets only in its object model, so the server cannot reconstruct them
from the file. Every read and every edit is a live COM call.

## Setup

    uv venv
    uv pip install -e ".[dev]"

Word must run in the same user session as the server, with the target document
open.

## Register with an MCP client

OpenCode reads `mcp` entries from its config. The global config is at
`~/.config/opencode/opencode.jsonc`. A project can also carry one.

Replace the placeholder with the path to this checkout, in the form your machine
uses (`D:/word-document-mcp`, `/home/you/word-document-mcp`). Forward slashes work
everywhere.

    {
      "mcp": {
        "word": {
          "type": "local",
          "command": ["uv", "run", "--directory", "<path to this checkout>", "word-mcp"],
          "enabled": true
        }
      }
    }

`uv run` creates the environment when it is missing, and resolves dependencies at
launch. To skip that work on every start, point at the environment interpreter
directly. Use `<path to this checkout>/.venv/Scripts/python.exe` on Windows, and
`<path to this checkout>/.venv/bin/python` on macOS and Linux.

    "command": ["<path to this checkout>/.venv/Scripts/python.exe", "-m", "word_mcp"]

## How a session runs

`connect_document` pins one open document. Every other tool acts on that pinned
document, and on nothing else.

```text
connect_document        pin one open document
  document_outline      headings, tables, sections, images, and a token for each
  inspect_document      styles, comments, revisions, bookmarks, content controls,
                        properties
  find_text             search for text, and get candidate tokens
  read_range            read a target, optionally with its formatting
  get_formatting        font and paragraph values, and deviations_from_style
  edit_text             insert, replace, or delete, then report the new range as
                        `at`
  revert_change         undo the newest change
```

A `target` names what to act on. A `token` from an earlier read is the most
precise form, and it stays valid until the document changes. You can also target
by `text`, by `style`, by `heading_path`, or by `occurrence`. Ambiguity is never
guessed. The server returns the candidates instead.

One pass looks like this.

```text
connect_document("quarterly.docx")                      pin the document
document_outline()                                      a token per heading
find_text("Draft", limit=5)                             candidates, not a guess
edit_text({"text": "Draft"}, "replace", text="Final")    the result carries `at`
revert_change(1)                                        undo it
```

Two caches make a repeat read cheap. `document_outline` builds a structure index.
`inspect_document` builds an inventory. Each cache stays valid until the document
changes underneath it. Most edits keep both.

```text
both caches kept
  edit_text, in the body or in a header or footer
  apply_edits with a single edit
  edit_page_layout header_footer
  edit_media size or wrap
  manage_list

index kept, inventory re-read
  edit_annotations add_comment or delete_comment
  set_document_properties

both caches re-read
  edit_page_layout page_setup
  apply_edits with two or more edits
  insert_ooxml, edit_table, edit_fields, manage_style
```

An edit that can describe the span it changed patches the index in place. An edit
that cannot drops the index, and dropping the index drops the inventory with it.
An edit that only changes the inventory keeps the index, and the next
`inspect_document` re-reads the inventory.

A restyle you make by hand in Word is invisible to the index. Pass
`document_outline(refresh=true)` to force a rebuild.

## Tests

    uv run pytest                    # unit tests, no Word needed
    uv run pytest -m integration     # needs Word open with samples/plain.docx

The integration tests edit the open sample document. The last test saves a
temporary copy and puts the original back. To start from a clean sample, close the
document in Word, run `uv run python samples/make_samples.py`, then reopen
`samples/plain.docx`.

## Manual checks

These checks need a person at the keyboard. Nobody has run them here, so each one
states the result to expect.

- A modal dialog is open in Word, for example File > Options, and a tool is
  called. Expected: a `word_busy` error after the 10 second call timeout. The
  server must not hang. Not run, because the dialog blocks the COM server and a
  person must leave it open during the call.
- Word is killed while a tool call is in flight. Expected: a `word_disconnected`
  error, or `word_not_running` on the next reconnect. The pin is not cleared, so
  call `connect_document` when Word is back. Not run, because the Word processes on
  this machine hold unrelated documents, and killing them would risk that work.

## Timings

Measured on a 2,207-paragraph, 2-section, 42-table document.

| Call | First | Repeat |
|---|---|---|
| `document_outline`, VBA scan | 2.5 s | 0.006 s |
| `document_outline`, COM fallback | 11 to 22 s | 0.006 s |
| `inspect_document` | 1.2 s | 0.002 s |
| `find_text`, `limit=5` | 0.2 s | |
| `find_text`, `limit=20` | 0.7 s | |

An edit that drops a cache pays for the rebuild on the next read. That is about
2.7 s for the index and about 1.1 s for the inventory, so an edit in the `both
caches re-read` group costs about 3.8 s across the next two reads.

Turn on Trust access to the VBA project object model. It is the difference between
a 2.5 second outline and an 11 second one. Without it the server uses COM, reports
`structure_source: "com"`, and adds a `structure_hint` that says how to enable the
faster scan.

`find_text` costs about 30 ms per match it returns, and it skips the style unless
the search filters by style. The `limit` argument bounds the search itself, so
`limit=20` stops after 20 matches instead of visiting all 47.

The COM fallback walks collections with `IEnumVARIANT.Next(256)`. That halves the
paragraph enumeration, from 0.86 s to 0.44 s on this document.

## Known limitations

Each item below was checked against the sample document.

- Only `body`, `header`, and `footer` containers work. The names `footnote`,
  `endnote`, and `comment` pass target validation, and then every tool rejects them
  with `invalid_argument: Container 'footnote' is not supported yet.`
- `container="footer:2"` names a section. A bare name means section 1. The `which`
  argument is not reachable from a container, so a container always means the
  primary header or footer. `edit_page_layout("header_footer", which="first_page")`
  writes a first-page header that no container can read back.
- Only the newest change reverts, in reverse order, and only a change that reports
  an `operation_id` is recorded. These tools report one: `edit_text`,
  `format_text`, `manage_style("apply")`, `insert_ooxml`, and each entry of
  `apply_edits`. These tools do not, so they cannot be reverted: `edit_table`,
  `edit_page_layout`, `edit_fields`, `manage_list`, `edit_media`, `edit_links`,
  `edit_annotations`, `edit_content_controls`. Revert restores text and formatting.
  It does not restore deleted images, fields, or comments.
- `find_text` returns at most 50 matches, whatever `limit` says.
- `resolve_comment` does not work on this Word build. It returns `word_error` with
  Word's message "This command is not available". Adding, replying to, and deleting
  comments all work.
- One document is pinned per session, and only the pinned document is reachable.
  Call `connect_document` to switch.
- `word_property` is not registered. It was skipped in v1, so no tool reads or
  writes a single document property on its own.
