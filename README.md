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
