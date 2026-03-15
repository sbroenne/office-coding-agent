"""PowerPoint MCP server for pytest-skill-engineering AI tool tests."""

from __future__ import annotations

from pathlib import Path

from manifest_mcp_common import run_manifest_server
from powerpoint_sim import PowerPointSimulator


def main() -> None:
    run_manifest_server(
        "powerpoint-ai-addin-test-server",
        PowerPointSimulator(),
        default_manifest=Path(__file__).parent / "manifests" / "powerpoint-tools-manifest.json",
    )


if __name__ == "__main__":
    main()
