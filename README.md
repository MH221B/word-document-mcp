# word-mcp

MCP server that reads and edits the Word document you have open in Word desktop
on Windows. See `docs/superpowers/specs/2026-10-14-word-document-mcp-design.md`.

## Setup

    uv venv
    uv pip install -e ".[dev]"

## Register with an MCP client

    {
      "mcpServers": {
        "word": { "command": "uv", "args": ["run", "--directory", "D:/word-document-mcp", "word-mcp"] }
      }
    }

Word must be running with the target document open, in the same user session.

## Tests

    uv run pytest                      # unit tests, no Word needed
    uv run pytest -m integration       # requires Word open with samples/plain.docx

## Known behaviour

Manual checks from the plan. These need a human at the keyboard, so they are
recorded as not executed in this environment, with the behaviour to expect.

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

## Client registration check

Confirmed programmatically: the server registers 24 tools and does not register
`word_property` (skipped in v1).

