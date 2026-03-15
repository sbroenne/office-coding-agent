"""Word MCP server for pytest-skill-engineering AI tool tests."""

from __future__ import annotations

from pathlib import Path

from manifest_mcp_common import run_manifest_server
from word_sim import WordSimulator


def main() -> None:
    run_manifest_server(
        "word-ai-addin-test-server",
        WordSimulator(),
        default_manifest=Path(__file__).parent / "manifests" / "word-tools-manifest.json",
    )


if __name__ == "__main__":
    main()
