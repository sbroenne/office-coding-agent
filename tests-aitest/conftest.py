"""Configuration for pytest-skill-engineering Office tool tests.

These tests validate that LLMs can correctly use the Office host tool schemas
exposed through manifest-driven MCP servers backed by in-memory simulators.

Run with: uv run pytest tests-aitest/ -v
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

from pytest_skill_engineering import MCPServer, Wait

# ---------------------------------------------------------------------------
# Environment setup
# ---------------------------------------------------------------------------

# LiteLLM bug workaround: their logging code expects AZURE_API_BASE but
# Azure SDK uses AZURE_OPENAI_ENDPOINT. Set both to silence the warning.
if os.environ.get("AZURE_OPENAI_ENDPOINT") and not os.environ.get("AZURE_API_BASE"):
    os.environ["AZURE_API_BASE"] = os.environ["AZURE_OPENAI_ENDPOINT"]


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_MODEL = "gpt-5-mini"
DEFAULT_RPM = 10
DEFAULT_TPM = 10000
DEFAULT_MAX_TURNS = 10

_PROJECT_ROOT = Path(__file__).parent.parent
AITEST_MANIFEST_DIR = _PROJECT_ROOT / "tests-aitest" / "manifests"
EXCEL_MANIFEST_PATH = AITEST_MANIFEST_DIR / "excel-tools-manifest.json"
POWERPOINT_MANIFEST_PATH = AITEST_MANIFEST_DIR / "powerpoint-tools-manifest.json"
WORD_MANIFEST_PATH = AITEST_MANIFEST_DIR / "word-tools-manifest.json"
OUTLOOK_MANIFEST_PATH = AITEST_MANIFEST_DIR / "outlook-tools-manifest.json"

BASE_PROMPT_PATH = _PROJECT_ROOT / "src" / "services" / "ai" / "BASE_PROMPT.md"
PROMPTS_DIR = _PROJECT_ROOT / "src" / "services" / "ai" / "prompts"
APP_PROMPT_PATHS = {
    "excel": PROMPTS_DIR / "EXCEL_APP_PROMPT.md",
    "powerpoint": PROMPTS_DIR / "POWERPOINT_APP_PROMPT.md",
    "word": PROMPTS_DIR / "WORD_APP_PROMPT.md",
    "outlook": PROMPTS_DIR / "OUTLOOK_APP_PROMPT.md",
}


def load_system_prompt(host: str) -> str:
    """Compose the production system prompt for a specific Office host."""
    base_prompt = BASE_PROMPT_PATH.read_text(encoding="utf-8").strip()
    app_prompt = APP_PROMPT_PATHS[host].read_text(encoding="utf-8").strip()
    return f"{base_prompt}\n\n{app_prompt}"


SYSTEM_PROMPTS = {host: load_system_prompt(host) for host in APP_PROMPT_PATHS}


# ---------------------------------------------------------------------------
# Pytest configuration
# ---------------------------------------------------------------------------


def pytest_configure(config: pytest.Config) -> None:
    """Register AI-test markers used in this directory."""
    for marker in [
        "integration: live LLM integration test",
        "excel: Excel host AI test",
        "powerpoint: PowerPoint host AI test",
        "word: Word host AI test",
        "outlook: Outlook host AI test",
        "token_efficiency: token efficiency experiment",
        "adversarial: adversarial AI eval that probes tool quality edge cases",
    ]:
        config.addinivalue_line("markers", marker)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _build_server(script_name: str, manifest_path: Path, wait_for: list[str]) -> MCPServer:
    if not manifest_path.exists():
        pytest.skip(f"Manifest not found: {manifest_path}. Run 'npm run manifest' first.")

    return MCPServer(
        command=[
            sys.executable,
            "-u",
            str(Path(__file__).parent / script_name),
            "--manifest",
            str(manifest_path),
        ],
        wait=Wait.for_tools(wait_for),
    )


@pytest.fixture(scope="module")
def excel_server() -> MCPServer:
    """Excel MCP server backed by the in-memory spreadsheet simulator."""
    return _build_server(
        "excel_mcp.py",
        EXCEL_MANIFEST_PATH,
        ["get_range_values", "set_range_values", "list_sheets", "get_used_range"],
    )


@pytest.fixture(scope="module")
def powerpoint_server() -> MCPServer:
    """PowerPoint MCP server backed by an in-memory presentation simulator."""
    return _build_server(
        "powerpoint_mcp.py",
        POWERPOINT_MANIFEST_PATH,
        ["get_presentation_overview", "set_presentation_content", "add_slide_from_code"],
    )


@pytest.fixture(scope="module")
def word_server() -> MCPServer:
    """Word MCP server backed by an in-memory document simulator."""
    return _build_server(
        "word_mcp.py",
        WORD_MANIFEST_PATH,
        ["get_document_overview", "insert_content_at_selection", "find_and_replace"],
    )


@pytest.fixture(scope="module")
def outlook_server() -> MCPServer:
    """Outlook MCP server backed by an in-memory mailbox simulator."""
    return _build_server(
        "outlook_mcp.py",
        OUTLOOK_MANIFEST_PATH,
        ["get_mail_item", "display_new_message", "reply_to_mail"],
    )

