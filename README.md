# word-mcp

MCP server that reads and edits the Word document you have open in Word desktop
on Windows.

## Setup

    uv venv
    uv pip install -e ".[dev]"

## Register with an MCP client

OpenCode reads `mcp` entries from its config, global at
`~/.config/opencode/opencode.jsonc` or per project. Put the checkout's own path
where the placeholder is, in the form that machine uses (`D:/word-document-mcp`,
`/home/you/word-document-mcp`); forward slashes work everywhere:

    {
      "mcp": {
        "word": {
          "type": "local",
          "command": ["uv", "run", "--directory", "<path to this checkout>", "word-mcp"],
          "enabled": true
        }
      }
    }

`uv run` creates the environment if it is missing and resolves dependencies at
launch, so nothing else has to be set up first. To skip that resolution on every
start, point at the environment's interpreter directly, which is
`<path to this checkout>/.venv/Scripts/python.exe` on Windows and
`<path to this checkout>/.venv/bin/python` on macOS and Linux:

    "command": ["<path to this checkout>/.venv/Scripts/python.exe", "-m", "word_mcp"]

Word must be running with the target document open, in the same user session.

## Tests

    uv run pytest                      # unit tests, no Word needed
    uv run pytest -m integration       # requires Word open with samples/plain.docx

The integration tests edit the open sample document, and the last one saves it to
a temporary copy and then puts it back. Regenerate a clean sample before a full
run: close the document in Word, run `uv run python samples/make_samples.py`, then
reopen `samples/plain.docx`.

## Known behaviour

Manual checks. These need a human at the keyboard, so they are recorded as not
executed in this environment, with the behaviour to expect.

- Modal dialog open in Word (for example File > Options) while a tool is called.
  Expected: a `word_busy` error after the ~10s call timeout, not a hang. Not run:
  opening the dialog blocks the COM server and needs a human to leave it open
  mid-call.
- Word killed while a tool call is in flight.
  Expected: a `word_disconnected` (or `word_not_running` on reconnect) error. The
  pin is not cleared automatically, so call `connect_document` once Word is back.
  Not run: the Word instances on this machine also hold unrelated open documents,
  and killing them would risk that work.
- Reads are cached per document version. On a 2,207-paragraph, 2-section, 42-table
  document the first `document_outline` takes about 2.5 seconds through the batched
  VBA scan and 11-22 seconds through plain COM depending on Word state, with every
  later call on the
  unchanged document about 0.006 seconds; the first `inspect_document` takes about
  1.2 seconds and a repeat about 0.002 seconds. The VBA scan needs Trust access
  to the VBA project object model in Word; without it the server falls back to
  COM and reports `structure_source: "com"` with a `structure_hint`. `find_text` pays Word's scan plus
  about 30 ms per match it returns (its style is skipped unless the search
  filters by style), so `limit=5` takes about 0.2 seconds and
  `limit=20` about 0.7 seconds rather than visiting all 47 matches. An edit that
  can describe what it changed patches the cached index in place, so the read
  after an edit is warm instead of a full rebuild; anything it cannot describe
  drops the cache, which is rebuilt when the document's character count changes.
  A restyle made by hand in Word is invisible to that check: pass
  `document_outline(refresh=true)` to force a rebuild. A multi-change
  `apply_edits` drops the cached index, so the next read rebuilds it. The COM
  fallback walks collections with `IEnumVARIANT.Next(256)`, which halves the
  paragraph enumeration (~0.86 s to ~0.44 s on the 2,207-paragraph document).

## Known limitations

Each line below was checked against the sample document, not inferred.

- **Containers.** Only `body`, `header`, and `footer` work. `footnote`, `endnote`,
  and `comment` pass target validation and are then rejected by every tool with
  `invalid_argument: Container 'footnote' is not supported yet.`
- **One header and footer per section.** `container="footer:2"` names a section (a
  bare name means section 1), but `which` is not reachable from a container, so a
  container always means the *primary* header or footer.
  `edit_page_layout("header_footer", which="first_page")` writes a first-page
  header that no container can read back.
- **Revert is narrow.** Only the newest change reverts, in reverse order, and only
  changes that report an `operation_id` are recorded: `edit_text`, `format_text`,
  `manage_style("apply")`, `insert_ooxml`, and each entry of `apply_edits`.
  `edit_table`, `edit_page_layout`, `edit_fields`, `manage_list`, `edit_media`,
  `edit_links`, `edit_annotations`, and `edit_content_controls` are not recorded
  and cannot be reverted. Revert restores text and formatting only: deleted
  images, fields, and comments are not restored.
- **`find_text` returns at most 50 matches** (`MAX_CANDIDATES`), whatever `limit`
  says.
- **`resolve_comment` does not work on this Word build.** It returns `word_error`
  carrying Word's "This command is not available". Adding, replying to, and
  deleting comments all work.
- **One document is pinned per session**, and only the pinned document is
  reachable. Call `connect_document` to switch.
- **`insert_ooxml` never reports the span it changed**, so the read after it is a
  full rebuild rather than an in-place patch.
- **`word_property` is not registered** (skipped in v1), so no tool reads or writes
  a single document property in isolation.

## Client registration check

Confirmed programmatically: the server registers 25 tools and does not register
`word_property` (skipped in v1).

