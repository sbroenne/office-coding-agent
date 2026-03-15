"""Adversarial Excel MCP server integration tests.

These scenarios are intentionally less scripted than the happy-path evals in
``test_excel_tools.py``. They probe ambiguous phrasing, error recovery,
multistep sequencing, boundary inputs, tool confusion, and cross-category
workflows that are more likely to expose product-quality issues.
"""

from __future__ import annotations

import pytest

from pytest_skill_engineering import Eval, MCPServer, Provider

from conftest import DEFAULT_MAX_TURNS, DEFAULT_MODEL, DEFAULT_RPM, DEFAULT_TPM, SYSTEM_PROMPTS

pytestmark = [pytest.mark.integration, pytest.mark.excel, pytest.mark.adversarial]

EXCEL_PROMPT = SYSTEM_PROMPTS["excel"]


def _make_eval(
    excel_server: MCPServer,
    name: str,
    *,
    allowed_tools: list[str] | None = None,
    max_turns: int = DEFAULT_MAX_TURNS,
) -> Eval:
    """Create an Excel eval with the production prompt and scoped tool visibility."""
    return Eval(
        name=name,
        provider=Provider(model=f"azure/{DEFAULT_MODEL}", rpm=DEFAULT_RPM, tpm=DEFAULT_TPM),
        mcp_servers=[excel_server],
        system_prompt=EXCEL_PROMPT,
        max_turns=max_turns,
        allowed_tools=allowed_tools,
    )


def _tool_names(result) -> list[str]:
    return [call.name for call in result.all_tool_calls]


def _assert_called_in_order(result, *tool_names: str) -> None:
    calls = _tool_names(result)
    previous_index = -1
    for tool_name in tool_names:
        assert tool_name in calls, f"Expected {tool_name} in call trace {calls}"
        current_index = calls.index(tool_name)
        assert current_index > previous_index, f"Expected {tool_name} after {calls[previous_index] if previous_index >= 0 else 'start'} in call trace {calls}"
        previous_index = current_index


class TestAmbiguousInstructions:
    """Prompts where the right tool is semantically subtle, not keyword-driven."""

    async def test_percentage_display_prefers_number_format(self, eval_run, excel_server):
        """Stress: 'show percentages' should drive number formatting, not generic cell styling."""
        agent = _make_eval(
            excel_server,
            "ambiguous-percentage-format",
            allowed_tools=["set_range_values", "set_number_format", "format_range"],
        )

        result = await eval_run(
            agent,
            "Write [[0.125], [0.5], [0.875]] to A1:A3. "
            "Then make A1:A3 show percentages with one decimal place without changing the values themselves.",
        )

        tool_names = _tool_names(result)
        assert result.success
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("set_number_format")
        assert "format_range" not in tool_names


class TestErrorRecovery:
    """Prompts that deliberately create a recoverable failure and check whether the agent continues."""

    async def test_duplicate_sheet_name_recovers_by_listing(self, eval_run, excel_server):
        """Stress: the second create_sheet should fail, forcing the agent to recover instead of stalling."""
        agent = _make_eval(
            excel_server,
            "error-recovery-duplicate-sheet",
            allowed_tools=["create_sheet", "list_sheets"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Create a worksheet named 'Budget'. Then try to create another worksheet with the exact same name. "
            "If Excel says it already exists, recover by listing the worksheets and confirming the workbook is still usable.",
        )

        tool_names = _tool_names(result)
        assert result.success
        assert tool_names.count("create_sheet") >= 2
        assert result.tool_was_called("list_sheets")

    async def test_missing_table_recovery_builds_then_retries(self, eval_run, excel_server):
        """Stress: the agent must handle a missing-table failure and continue with a corrective plan."""
        agent = _make_eval(
            excel_server,
            "error-recovery-missing-table",
            allowed_tools=["add_table_rows", "set_range_values", "create_table", "get_table_data"],
            max_turns=12,
        )

        result = await eval_run(
            agent,
            "First, try to append [[\"Carol\", 32]] to a table named People. "
            "If that fails because the table does not exist, recover by writing [['Name', 'Age'], ['Alice', 30], ['Bob', 25]] to A1:B3, "
            "creating a table named People over A1:B3, retrying the append, and then reading the table back.",
        )

        tool_names = _tool_names(result)
        assert result.success
        assert tool_names.count("add_table_rows") >= 2
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("create_table")
        assert result.tool_was_called("get_table_data")


class TestMultiStepReasoning:
    """Tasks that require chaining multiple tool categories in a sensible order."""

    async def test_sales_summary_requires_ordered_tool_chain(self, eval_run, excel_server):
        """Stress: writing data, tabularizing it, charting it, and formatting headers requires 4-step planning."""
        agent = _make_eval(
            excel_server,
            "multi-step-sales-summary",
            allowed_tools=["set_range_values", "create_table", "create_chart", "format_range"],
            max_turns=8,
        )

        result = await eval_run(
            agent,
            "Create a sales summary on Sheet1: write [['Product', 'Q1', 'Q2'], ['Alpha', 120, 150], ['Beta', 90, 140]] to A1:C3, "
            "turn A1:C3 into a table named SalesSummary, create a clustered column chart from that data, and make the header row bold.",
        )

        tool_names = _tool_names(result)
        assert result.success
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("create_table")
        assert result.tool_was_called("create_chart")
        assert result.tool_was_called("format_range")
        _assert_called_in_order(result, "set_range_values", "create_table", "create_chart")
        assert tool_names.index("format_range") > tool_names.index("set_range_values")


class TestEdgeCasesAndBoundaries:
    """Boundary scenarios that are easy for tools to mishandle."""

    async def test_special_character_sheet_name_round_trip(self, eval_run, excel_server):
        """Stress: quoted sheet names with spaces and ampersands must survive write + read flows."""
        agent = _make_eval(
            excel_server,
            "edge-special-sheet-name",
            allowed_tools=["create_sheet", "set_range_values", "get_range_values"],
        )

        result = await eval_run(
            agent,
            "Create a worksheet named 'Q4 Sales & Ops'. Write [['ready']] to 'Q4 Sales & Ops'!B2, then read back 'Q4 Sales & Ops'!B2.",
        )

        assert result.success
        assert result.tool_was_called("create_sheet")
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("get_range_values")

    async def test_max_excel_address_single_cell_round_trip(self, eval_run, excel_server):
        """Stress: single-cell operations at Excel's extreme bottom-right address."""
        agent = _make_eval(
            excel_server,
            "edge-max-address",
            allowed_tools=["set_range_values", "get_range_values"],
        )

        result = await eval_run(
            agent,
            "Write [['edge']] to cell XFD1048576, then read back that exact cell and confirm the text is edge.",
        )

        assert result.success
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("get_range_values")


class TestToolConfusion:
    """Prompts that sound similar but should map to different categories of tools."""

    async def test_blocking_negative_input_prefers_validation(self, eval_run, excel_server):
        """Stress: 'prevent entry' should choose validation, not conditional formatting."""
        agent = _make_eval(
            excel_server,
            "tool-confusion-negative-input",
            allowed_tools=["set_number_validation", "add_cell_value_format"],
        )

        result = await eval_run(
            agent,
            "Prevent users from entering negative numbers in D2:D20. Block bad input entirely; do not just color the cells.",
        )

        tool_names = _tool_names(result)
        assert result.success
        assert result.tool_was_called("set_number_validation")
        assert "add_cell_value_format" not in tool_names

    async def test_highlighting_outliers_prefers_conditional_format(self, eval_run, excel_server):
        """Stress: 'highlight values' should choose formatting, not a blocking validation rule."""
        agent = _make_eval(
            excel_server,
            "tool-confusion-highlight-outliers",
            allowed_tools=["set_number_validation", "add_cell_value_format"],
        )

        result = await eval_run(
            agent,
            "Highlight values above 100 in E2:E20 with a red background, but do not restrict what users are allowed to type.",
        )

        tool_names = _tool_names(result)
        assert result.success
        assert result.tool_was_called("add_cell_value_format")
        assert "set_number_validation" not in tool_names


class TestCrossCategoryWorkflows:
    """Composite requests that span range, formatting, validation, tables, and conditional formats."""

    async def test_kpi_tracker_crosses_multiple_tool_categories(self, eval_run, excel_server):
        """Stress: one prompt requires coordinated range, table, format, validation, and conditional-format decisions."""
        agent = _make_eval(
            excel_server,
            "cross-category-kpi-tracker",
            allowed_tools=[
                "set_range_values",
                "create_table",
                "format_range",
                "set_number_format",
                "set_list_validation",
                "add_cell_value_format",
            ],
            max_turns=12,
        )

        result = await eval_run(
            agent,
            "Build a KPI tracker on Sheet1. Write [['Metric', 'Owner', 'Completion', 'Status'], ['Launch', 'Ava', 0.8, 'On Track'], ['Migration', 'Ben', 0.45, 'At Risk'], ['Training', 'Chen', 0.6, 'On Track']] to A1:D4. "
            "Convert that range into a table named KPIs. Make the header row bold. Format C2:C4 as percentages. "
            "Add a dropdown on D2:D4 with options On Track, At Risk, Off Track. Highlight completion values below 50% in red.",
        )

        tool_names = _tool_names(result)
        assert result.success
        assert result.tool_was_called("set_range_values")
        assert result.tool_was_called("create_table")
        assert result.tool_was_called("format_range")
        assert result.tool_was_called("set_number_format")
        assert result.tool_was_called("set_list_validation")
        assert result.tool_was_called("add_cell_value_format")
        _assert_called_in_order(result, "set_range_values", "create_table")
        assert tool_names.index("format_range") > tool_names.index("set_range_values")
        assert tool_names.index("set_number_format") > tool_names.index("set_range_values")
        assert tool_names.index("set_list_validation") > tool_names.index("set_range_values")
        assert tool_names.index("add_cell_value_format") > tool_names.index("set_range_values")
