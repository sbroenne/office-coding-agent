"""Excel MCP server integration tests.

Test that an LLM can correctly use the Excel tools exposed via the manifest-driven
MCP server backed by ExcelSimulator. These tests validate the tool schemas are
understandable to the LLM — the key goal of the codegen system.

Run with: uv run pytest tests-aitest/ -v

Note: These tests use real LLM calls and cost money. Use --lf to re-run failures only.
"""

from __future__ import annotations

import pytest

from pytest_aitest import Agent, MCPServer, Provider

from conftest import (
    DEFAULT_MAX_TURNS,
    DEFAULT_MODEL,
    DEFAULT_RPM,
    DEFAULT_TPM,
    SYSTEM_PROMPT_PATH,
)

pytestmark = [pytest.mark.integration, pytest.mark.excel]

# Load the real production system prompt — the same one users get
EXCEL_PROMPT = SYSTEM_PROMPT_PATH.read_text(encoding="utf-8").strip()


def _make_agent(
    excel_server: MCPServer,
    name: str,
    *,
    allowed_tools: list[str] | None = None,
    max_turns: int = DEFAULT_MAX_TURNS,
) -> Agent:
    """Create an Excel agent with standard config.

    Use allowed_tools to limit which tools the LLM sees, reducing token usage
    from ~50k (all 59 tools) to ~2-5k (focused set).
    """
    return Agent(
        name=name,
        provider=Provider(model=f"azure/{DEFAULT_MODEL}", rpm=DEFAULT_RPM, tpm=DEFAULT_TPM),
        mcp_servers=[excel_server],
        system_prompt=EXCEL_PROMPT,
        max_turns=max_turns,
        allowed_tools=allowed_tools,
    )


# =============================================================================
# Range Operations — Core Read/Write
# =============================================================================


class TestRangeOperations:
    """Test that the LLM can read and write range data correctly."""

    async def test_write_and_read_range(self, aitest_run, excel_server):
        """Write values then read them back — validates round-trip data flow."""
        agent = _make_agent(
            excel_server, "write-read",
            allowed_tools=["set_range_values", "get_range_values"],
        )

        result = await aitest_run(
            agent,
            "Write the values [[10, 20], [30, 40]] to range A1:B2, "
            "then read back the values from that range and show them to me.",
        )

        assert result.success
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("get_range_values")

    async def test_clear_range(self, aitest_run, excel_server):
        """Write data then clear it using the clear_range tool."""
        agent = _make_agent(
            excel_server, "clear-range",
            allowed_tools=["set_range_values", "clear_range"],
        )

        result = await aitest_run(
            agent,
            "Write [[1, 2, 3]] to A1:C1. "
            "Then use the clear_range tool to clear that same range.",
        )

        assert result.success
        assert result.tool_was_called("set_range_values")


# =============================================================================
# Sheet Operations
# =============================================================================


class TestSheetOperations:
    """Test that the LLM can manage worksheets."""

    async def test_create_and_list_sheets(self, aitest_run, excel_server):
        """Create a new sheet, then list all sheets."""
        agent = _make_agent(
            excel_server, "create-sheet",
            allowed_tools=["create_sheet", "list_sheets"],
        )

        result = await aitest_run(
            agent,
            "Create a new worksheet called 'SalesData', then list all worksheets.",
        )

        assert result.success
        assert result.tool_was_called("create_sheet")
        assert result.tool_was_called("list_sheets")


# =============================================================================
# Table Operations
# =============================================================================


class TestTableOperations:
    """Test that the LLM can create and manage tables."""

    async def test_create_table(self, aitest_run, excel_server):
        """Write data then create a table over it."""
        agent = _make_agent(
            excel_server, "create-table",
            allowed_tools=["set_range_values", "create_table"],
        )

        result = await aitest_run(
            agent,
            "Write [['Name', 'Age'], ['Alice', 30], ['Bob', 25]] to A1:B3. "
            "Then create a table named 'People' over range A1:B3.",
        )

        assert result.success
        assert result.tool_was_called("create_table")


# =============================================================================
# Conditional Format — Tool Selection (Decomposed)
# =============================================================================

# All 6 conditional format tools — tests must pick ONE from this set
_CF_TOOLS = [
    "add_color_scale", "add_data_bar", "add_cell_value_format",
    "add_top_bottom_format", "add_contains_text_format", "add_custom_format",
]


class TestConditionalFormatSelection:
    """Test that the LLM picks the RIGHT conditional format tool.

    These tests are the key validation for the decomposition strategy.
    The old mega-tool required a complex 'type' discriminator. The new
    decomposed tools should be self-selecting based on description alone.
    """

    async def test_selects_color_scale(self, aitest_run, excel_server):
        """LLM should pick add_color_scale for gradient formatting."""
        agent = _make_agent(excel_server, "cf-color-scale", allowed_tools=_CF_TOOLS)

        result = await aitest_run(
            agent,
            "Apply a color scale (red to green gradient) to cells A1:A10.",
        )

        assert result.success
        assert result.tool_was_called("add_color_scale")

    async def test_selects_data_bar(self, aitest_run, excel_server):
        """LLM should pick add_data_bar for bar chart formatting."""
        agent = _make_agent(excel_server, "cf-data-bar", allowed_tools=_CF_TOOLS)

        result = await aitest_run(
            agent,
            "Add data bars to cells B1:B20 to visualize the values.",
        )

        assert result.success
        assert result.tool_was_called("add_data_bar")

    async def test_selects_cell_value_format(self, aitest_run, excel_server):
        """LLM should pick add_cell_value_format for value-based rules."""
        agent = _make_agent(excel_server, "cf-cell-value", allowed_tools=_CF_TOOLS)

        result = await aitest_run(
            agent,
            "Highlight cells in C1:C10 with a red background if the value is greater than 100.",
        )

        assert result.success
        assert result.tool_was_called("add_cell_value_format")

    async def test_selects_contains_text(self, aitest_run, excel_server):
        """LLM should pick add_contains_text_format for text matching."""
        agent = _make_agent(excel_server, "cf-text", allowed_tools=_CF_TOOLS)

        result = await aitest_run(
            agent,
            "Highlight cells in D1:D50 that contain the word 'Error' with a yellow background.",
        )

        assert result.success
        assert result.tool_was_called("add_contains_text_format")


# =============================================================================
# Data Validation — Tool Selection (Decomposed)
# =============================================================================

# All 5 data validation tools
_DV_TOOLS = [
    "set_list_validation", "set_number_validation", "set_date_validation",
    "set_text_length_validation", "set_custom_validation",
]


class TestDataValidationSelection:
    """Test that the LLM picks the right validation tool."""

    async def test_selects_list_validation(self, aitest_run, excel_server):
        """LLM should pick set_list_validation for dropdown lists."""
        agent = _make_agent(excel_server, "dv-list", allowed_tools=_DV_TOOLS)

        result = await aitest_run(
            agent,
            "Set a dropdown list validation on A1:A20 with options: Yes, No, Maybe.",
        )

        assert result.success
        assert result.tool_was_called("set_list_validation")

    async def test_selects_number_validation(self, aitest_run, excel_server):
        """LLM should pick set_number_validation for numeric constraints."""
        agent = _make_agent(excel_server, "dv-number", allowed_tools=_DV_TOOLS)

        result = await aitest_run(
            agent,
            "Set whole number validation on B1:B10 so only values between 1 and 100 are allowed.",
        )

        assert result.success
        assert result.tool_was_called("set_number_validation")


# =============================================================================
# Comment Operations
# =============================================================================


class TestCommentOperations:
    """Test that the LLM can add and manage comments."""

    async def test_add_comment(self, aitest_run, excel_server):
        """Add a comment to a cell."""
        agent = _make_agent(
            excel_server, "comments",
            allowed_tools=["add_comment", "list_comments", "edit_comment", "delete_comment"],
        )

        result = await aitest_run(
            agent,
            "Add a comment 'Review this formula' to cell A5.",
        )

        assert result.success
        assert result.tool_was_called("add_comment")


# =============================================================================
# Multi-Step Workflows
# =============================================================================


class TestMultiStepWorkflows:
    """Complex workflows that chain multiple tool categories."""

    async def test_write_and_format(self, aitest_run, excel_server):
        """Write data then format it — tests chaining range tools."""
        agent = _make_agent(
            excel_server, "write-format",
            allowed_tools=["set_range_values", "get_range_values", "format_range"],
        )

        result = await aitest_run(
            agent,
            "Write [['Revenue', 'Cost'], [50000, 35000]] to A1:B2. "
            "Then make the header row A1:B1 bold.",
        )

        assert result.success
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("format_range")


# =============================================================================
# Large Data — Token Efficiency
# =============================================================================


class TestLargeDataResponse:
    """Test tool responses with larger datasets.

    Exercises get_used_range on a populated sheet to measure token
    efficiency. The maxRows parameter allows the LLM to preview
    headers + a few rows before deciding whether it needs the full
    dataset.
    """

    async def test_populate_and_summarize(self, aitest_run, excel_server):
        """Write a 20-row dataset, then ask the LLM to read and summarize it.

        The LLM can choose to use maxRows to preview headers first,
        or read everything at once. We measure which approach it picks
        and how many tokens it costs.
        """
        # Build a realistic 21-row dataset (1 header + 20 data rows)
        headers = ["Product", "Region", "Q1 Sales", "Q2 Sales", "Q3 Sales", "Q4 Sales", "Total"]
        products = ["Widget A", "Widget B", "Gadget X", "Gadget Y", "Module Z"]
        regions = ["North", "South", "East", "West"]
        rows = [headers]
        for i in range(20):
            prod = products[i % len(products)]
            region = regions[i % len(regions)]
            q1, q2, q3, q4 = 1000 + i * 100, 1200 + i * 110, 900 + i * 90, 1100 + i * 105
            rows.append([prod, region, q1, q2, q3, q4, q1 + q2 + q3 + q4])

        agent = _make_agent(
            excel_server, "large-data",
            allowed_tools=["set_range_values", "get_used_range"],
            max_turns=5,
        )

        result = await aitest_run(
            agent,
            f"Write this data to A1:G21: {rows}. "
            "Then use get_used_range to read the full sheet and tell me "
            "which product has the highest total sales.",
        )

        assert result.success
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("get_used_range")
