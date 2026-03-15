"""Outlook MCP server for pytest-skill-engineering AI tool tests."""

from __future__ import annotations

from pathlib import Path

from manifest_mcp_common import run_manifest_server
from outlook_sim import OutlookSimulator


def main() -> None:
    run_manifest_server(
        "outlook-ai-addin-test-server",
        OutlookSimulator(),
        default_manifest=Path(__file__).parent / "manifests" / "outlook-tools-manifest.json",
    )


if __name__ == "__main__":
    main()
