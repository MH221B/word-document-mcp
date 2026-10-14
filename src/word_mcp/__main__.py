"""stdio entry point."""

from word_mcp.server import mcp


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
