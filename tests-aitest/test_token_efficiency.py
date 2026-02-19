"""Token efficiency experiments for range read operations.

These tests measure token usage across different strategies for reading
range data, to identify which approach minimises LLM costs on large sheets.

Strategies compared:
  1. Baseline   — get_range_values reads entire range at once (no paging)
  2. maxRows    — get_used_range with maxRows=N to preview headers then page
  3. Chunked    — agent pages through data in explicit row chunks via get_range_values
  4. Dimensions — agent calls get_used_range first (no values) to decide what to read

Results are printed after each test so they can be compared.

Run with: uv run pytest tests-aitest/test_token_efficiency.py -v -s
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

from pytest_aitest import Agent, MCPServer, Provider, Wait

from conftest import (
    DEFAULT_MAX_TURNS,
    DEFAULT_MODEL,
    DEFAULT_RPM,
    DEFAULT_TPM,
    MANIFEST_PATH,
    SYSTEM_PROMPT_PATH,
)

pytestmark = [pytest.mark.integration, pytest.mark.token_efficiency]

EXCEL_PROMPT = SYSTEM_PROMPT_PATH.read_text(encoding="utf-8").strip()

# ─── Shared dataset: 50 rows × 6 columns (~300 cells) ────────────────────────

HEADERS = ["Product", "Region", "Q1", "Q2", "Q3", "Q4"]
PRODUCTS = ["Widget A", "Widget B", "Gadget X", "Gadget Y", "Module Z"]
REGIONS = ["North", "South", "East", "West"]


def _make_dataset(num_rows: int = 50) -> list[list]:
    """Build a realistic sales dataset with header + num_rows data rows."""
    rows = [HEADERS]
    for i in range(num_rows):
        prod = PRODUCTS[i % len(PRODUCTS)]
        region = REGIONS[i % len(REGIONS)]
        q1, q2, q3, q4 = 1000 + i * 50, 1100 + i * 55, 900 + i * 45, 1050 + i * 52
        rows.append([prod, region, q1, q2, q3, q4])
    return rows


def _end_cell(num_rows: int, num_cols: int = 6) -> str:
    col = chr(ord("A") + num_cols - 1)
    return f"A1:{col}{num_rows + 1}"  # +1 for header


# ─── Fixture ─────────────────────────────────────────────────────────────────


@pytest.fixture(scope="module")
def excel_server():
    """Excel MCP server for token efficiency tests."""
    if not MANIFEST_PATH.exists():
        pytest.skip(f"Manifest not found: {MANIFEST_PATH}. Run 'npm run manifest' first.")
    return MCPServer(
        command=[
            sys.executable,
            "-u",
            str(Path(__file__).parent / "excel_mcp.py"),
            "--manifest",
            str(MANIFEST_PATH),
        ],
        wait=Wait.for_tools(["get_range_values", "get_used_range", "set_range_values"]),
    )


def _agent(excel_server: MCPServer, name: str, allowed_tools: list[str]) -> Agent:
    return Agent(
        name=name,
        provider=Provider(model=f"azure/{DEFAULT_MODEL}", rpm=DEFAULT_RPM, tpm=DEFAULT_TPM),
        mcp_servers=[excel_server],
        system_prompt=EXCEL_PROMPT,
        max_turns=DEFAULT_MAX_TURNS,
        allowed_tools=allowed_tools,
    )


def _print_tokens(label: str, token_usage: dict) -> None:
    prompt = token_usage.get("prompt", 0)
    completion = token_usage.get("completion", 0)
    print(f"\n[TOKEN REPORT] {label}")
    print(f"  prompt:     {prompt:,}")
    print(f"  completion: {completion:,}")
    print(f"  total:      {prompt + completion:,}")


# =============================================================================
# Strategy 1: Baseline — read entire range at once
# =============================================================================


class TestBaselineFullRead:
    """Baseline: agent writes data then reads the entire range at once.

    This is the current behaviour — no paging, all values returned in one
    tool response. Establishes the token ceiling to beat.
    """

    async def test_full_read_20_rows(self, aitest_run, excel_server):
        """Read 20 rows × 6 cols (120 cells) in a single get_range_values call."""
        dataset = _make_dataset(20)
        addr = _end_cell(20)

        agent = _agent(excel_server, "baseline-20", ["set_range_values", "get_range_values"])

        result = await aitest_run(
            agent,
            f"Write this data to {addr}: {dataset}. "
            "Then read back the entire range and tell me the total Q1 sales.",
        )

        assert result.success
        assert result.tool_was_called("get_range_values")
        _print_tokens("Baseline 20 rows × 6 cols (full read)", result.token_usage)

    async def test_full_read_50_rows(self, aitest_run, excel_server):
        """Read 50 rows × 6 cols (300 cells) in a single get_range_values call.

        With 50 rows the response JSON is significantly larger.
        """
        dataset = _make_dataset(50)
        addr = _end_cell(50)

        agent = _agent(excel_server, "baseline-50", ["set_range_values", "get_range_values"])

        result = await aitest_run(
            agent,
            f"Write this data to {addr}: {dataset}. "
            "Then read back the entire range and tell me which product appears most often.",
        )

        assert result.success
        assert result.tool_was_called("get_range_values")
        _print_tokens("Baseline 50 rows × 6 cols (full read)", result.token_usage)


# =============================================================================
# Strategy 2: Dimensions-first + targeted read
# =============================================================================


class TestDimensionsFirstRead:
    """Dimensions-first: agent calls get_used_range (no values) to discover
    shape, then calls get_range_values only for the rows/columns it needs.

    Goal: reduce tokens by letting the LLM decide what subset to read.
    """

    async def test_dimensions_then_targeted_read(self, aitest_run, excel_server):
        """Write 50 rows, then use get_used_range to discover shape before reading."""
        dataset = _make_dataset(50)
        addr = _end_cell(50)

        agent = _agent(
            excel_server,
            "dimensions-first",
            ["set_range_values", "get_used_range", "get_range_values"],
        )

        result = await aitest_run(
            agent,
            f"Write this data to {addr}: {dataset}. "
            "Use get_used_range first to find the sheet dimensions. "
            "Then read only the Q1 column (column C) and tell me the total Q1 sales.",
        )

        assert result.success
        _print_tokens("Dimensions-first + targeted column read (50 rows)", result.token_usage)
        # Report whether agent read the full range or just what it needed
        all_calls = result.all_tool_calls
        get_range_calls = [c for c in all_calls if c.name == "get_range_values"]
        print(f"  get_range_values calls: {len(get_range_calls)}")
        for c in get_range_calls:
            print(f"    address arg: {c.arguments.get('address', '?')}")


# =============================================================================
# Strategy 3: maxRows preview on get_used_range
# =============================================================================


class TestMaxRowsPreview:
    """maxRows preview: agent uses get_used_range(maxRows=N) to see headers
    and a few rows, then decides whether it needs to read more.

    This is the existing mechanism — tests how well the LLM exploits it.
    """

    async def test_maxrows_preview_summarise(self, aitest_run, excel_server):
        """Write 50 rows then ask for a summary using maxRows preview."""
        dataset = _make_dataset(50)
        addr = _end_cell(50)

        agent = _agent(
            excel_server,
            "maxrows-preview",
            ["set_range_values", "get_used_range"],
        )

        result = await aitest_run(
            agent,
            f"Write this data to {addr}: {dataset}. "
            "Use get_used_range with maxRows=5 to preview the sheet, "
            "then decide if you need to read more. "
            "Tell me which product has the highest Q1 sales.",
        )

        assert result.success
        assert result.tool_was_called("get_used_range")
        _print_tokens("maxRows preview (50 rows) - summarise task", result.token_usage)

        all_calls = result.all_tool_calls
        used_range_calls = [c for c in all_calls if c.name == "get_used_range"]
        for c in used_range_calls:
            max_rows = c.arguments.get("maxRows", "not set")
            print(f"  get_used_range(maxRows={max_rows})")

    async def test_maxrows_vs_full_read_token_delta(self, aitest_run, excel_server):
        """Compare: ask the agent to read all data vs use maxRows.

        This reveals whether the LLM self-selects a paged approach or
        always defaults to full reads.
        """
        dataset = _make_dataset(50)
        addr = _end_cell(50)

        # Agent with only get_used_range (forces maxRows path)
        agent_paged = _agent(
            excel_server,
            "forced-paged",
            ["set_range_values", "get_used_range"],
        )

        result_paged = await aitest_run(
            agent_paged,
            f"Write this data to {addr}: {dataset}. "
            "Read the sheet and count how many rows belong to the 'North' region.",
        )

        assert result_paged.success
        _print_tokens("Paged (get_used_range only, 50 rows)", result_paged.token_usage)

        # Agent with only get_range_values (forces full read path)
        agent_full = _agent(
            excel_server,
            "forced-full",
            ["set_range_values", "get_range_values"],
        )

        result_full = await aitest_run(
            agent_full,
            f"Write this data to {addr}: {dataset}. "
            "Read the sheet and count how many rows belong to the 'North' region.",
        )

        assert result_full.success
        _print_tokens("Full read (get_range_values only, 50 rows)", result_full.token_usage)

        paged_total = sum(result_paged.token_usage.values())
        full_total = sum(result_full.token_usage.values())
        saving_pct = (full_total - paged_total) / full_total * 100 if full_total else 0
        print(f"\n  [DELTA] paged={paged_total:,} vs full={full_total:,} → {saving_pct:.0f}% saving")


# =============================================================================
# Strategy 4: Natural LLM behaviour (no hints)
# =============================================================================


class TestNaturalBehaviour:
    """Does the LLM naturally pick an efficient strategy when given all tools?

    Exposes get_used_range AND get_range_values with no instructions on
    which to use, and measures what the model naturally selects.
    """

    async def test_natural_tool_selection_50_rows(self, aitest_run, excel_server):
        """Give LLM all read tools and ask a question requiring data inspection.

        Uses 20 rows to stay within TPM limits when both tools are available.
        """
        dataset = _make_dataset(20)
        addr = _end_cell(20)

        agent = _agent(
            excel_server,
            "natural-50",
            ["set_range_values", "get_used_range", "get_range_values"],
        )

        result = await aitest_run(
            agent,
            f"Write this data to {addr}: {dataset}. "
            "Which product has the highest average quarterly sales?",
        )

        assert result.success
        all_calls = result.all_tool_calls
        tools_used = [c.name for c in all_calls if c.name in ("get_used_range", "get_range_values")]
        _print_tokens("Natural tool selection (50 rows)", result.token_usage)
        print(f"  Tools chosen: {tools_used}")

        used_range_calls = [c for c in all_calls if c.name == "get_used_range"]
        range_calls = [c for c in all_calls if c.name == "get_range_values"]
        print(f"  get_used_range calls: {len(used_range_calls)}")
        print(f"  get_range_values calls: {len(range_calls)}")
