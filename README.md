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
  Expected: a `word_busy` error naming the dialog after the ~10s call timeout,
  not a hang. Not run: opening the dialog blocks the COM server and needs a
  human to leave it open mid-call.
- Word killed while a tool call is in flight.
  Expected: a `word_disconnected` (or `word_not_running` on reconnect) error, and
  the session reports no pinned document. Not run: the Word instances on this
  machine also hold unrelated open documents, and killing them would risk that
  work.
- A co-authored document with another author holding a lock.
  Expected: the tool reports a conflict rather than retrying forever. Not run:
  no co-authored document with a second author lock was available.
- Reads are cached per document version. On a 1,616-paragraph, 42-table document
  the first `document_outline` takes about 23 seconds and every later call on the
  unchanged document about 0.006 seconds; the first `inspect_document` takes about
  3.4 seconds and a repeat about 0.002 seconds. `find_text` pays Word's scan plus
  about 120 ms per match it returns, so `limit=5` takes about 0.7 seconds and
  `limit=20` about 1.8 seconds rather than visiting all 47 matches. An edit that
  can describe what it changed patches the cached index in place, so the read
  after an edit is warm instead of a full rebuild; anything it cannot describe
  drops the cache, which is rebuilt when the document's character count changes.
  A restyle made by hand in Word is invisible to that check: pass
  `document_outline(refresh=true)` to force a rebuild.

## Client registration check

Confirmed programmatically: the server registers 24 tools and does not register
`word_property` (skipped in v1).

