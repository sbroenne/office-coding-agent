"""PowerPoint MCP server integration tests using pytest-skill-engineering."""

from __future__ import annotations

import pytest

from pytest_skill_engineering import Eval, MCPServer, Provider

from conftest import DEFAULT_MAX_TURNS, DEFAULT_MODEL, DEFAULT_RPM, DEFAULT_TPM, SYSTEM_PROMPTS

pytestmark = [pytest.mark.integration, pytest.mark.powerpoint]

POWERPOINT_PROMPT = SYSTEM_PROMPTS["powerpoint"]


def _make_eval(
    powerpoint_server: MCPServer,
    name: str,
    *,
    allowed_tools: list[str] | None = None,
    max_turns: int = DEFAULT_MAX_TURNS,
) -> Eval:
    return Eval(
        name=name,
        provider=Provider(model=f"azure/{DEFAULT_MODEL}", rpm=DEFAULT_RPM, tpm=DEFAULT_TPM),
        mcp_servers=[powerpoint_server],
        system_prompt=POWERPOINT_PROMPT,
        max_turns=max_turns,
        allowed_tools=allowed_tools,
    )


class TestPowerPointToolSelection:
    """Validate that models can choose the right PowerPoint tools."""

    async def test_gets_presentation_info(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-overview",
            allowed_tools=["get_presentation_overview"],
        )

        result = await eval_run(agent, "Inspect the active presentation and tell me how many slides it has.")

        assert result.success
        assert result.tool_was_called("get_presentation_overview")

    async def test_adds_text_to_a_new_slide(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-add-text",
            allowed_tools=[
                "get_presentation_overview",
                "set_presentation_content",
                "get_presentation_content",
            ],
        )

        result = await eval_run(
            agent,
            "First call get_presentation_overview. The sample deck starts with 1 slide, so add a second slide by calling set_presentation_content with slideIndex 1 and text 'Q3 Review'. Then call get_presentation_content for slideIndex 1 to verify the new slide text.",
            max_turns=8,
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("set_presentation_content")
        assert result.tool_was_called("get_presentation_content")

    async def test_adds_a_geometric_shape(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-add-shape",
            allowed_tools=["get_presentation_overview", "add_geometric_shape"],
        )

        result = await eval_run(
            agent,
            "Inspect the deck, then add a blue rectangle to slide 1 near the top-left corner.",
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("add_geometric_shape")

    async def test_adds_an_image_via_slide_code(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-add-image",
            allowed_tools=["get_presentation_overview", "add_slide_from_code"],
        )

        result = await eval_run(
            agent,
            "First call get_presentation_overview. Then use add_slide_from_code to add a new slide containing an image with this data URI: data:image/png;base64,ZmFrZS1pbWFnZS1ieXRlcw==. Use a simple PptxGenJS slide.addImage call with the data URI and a small image box near the top-left of the slide.",
            max_turns=8,
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("add_slide_from_code")

    async def test_adds_a_chart_via_slide_code(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-add-chart",
            allowed_tools=["get_presentation_overview", "add_slide_from_code"],
        )

        result = await eval_run(
            agent,
            "First call get_presentation_overview. Then use add_slide_from_code to add a new slide with a bar chart titled 'Revenue by Quarter' using categories Q1, Q2, Q3, Q4 and values 12, 18, 24, 30.",
            max_turns=8,
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("add_slide_from_code")


# =============================================================================
# Slide Content — Read/Inspect
# =============================================================================


class TestSlideContentReading:
    """Validate that models can read slide content at various granularities."""

    async def test_reads_slide_count_from_overview(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-slide-count",
            allowed_tools=["get_presentation_overview"],
        )

        result = await eval_run(
            agent,
            "Call get_presentation_overview and report the total number of slides in the deck.",
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")

    async def test_reads_slide_content_by_index(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-slide-by-index",
            allowed_tools=["get_presentation_overview", "get_presentation_content"],
        )

        result = await eval_run(
            agent,
            "First call get_presentation_overview. Then call get_presentation_content with slideIndex 0 to read the content of the first slide.",
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("get_presentation_content")

    async def test_reads_all_slides_without_index(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-read-all-slides",
            allowed_tools=["get_presentation_content"],
        )

        result = await eval_run(
            agent,
            "Call get_presentation_content without specifying any slide index to retrieve content for every slide.",
        )

        assert result.success
        assert result.tool_was_called("get_presentation_content")

    async def test_reads_slide_range_with_start_and_end(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-slide-range",
            allowed_tools=["get_presentation_overview", "set_presentation_content", "get_presentation_content"],
            max_turns=8,
        )

        result = await eval_run(
            agent,
            "Call get_presentation_overview first. Add a second slide at index 1 with text 'Slide Two' using set_presentation_content. "
            "Then call get_presentation_content with startIndex 0 and endIndex 1 to read both slides at once.",
            max_turns=8,
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("set_presentation_content")
        assert result.tool_was_called("get_presentation_content")


# =============================================================================
# Slide Content — Write Operations
# =============================================================================


class TestSlideContentWriting:
    """Validate that models can add and modify slide content correctly."""

    async def test_adds_text_box_and_verifies(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-text-box-verify",
            allowed_tools=["set_presentation_content", "get_presentation_content"],
        )

        result = await eval_run(
            agent,
            "Add a text box with text 'Annual Review' to slideIndex 0 using set_presentation_content. "
            "Then call get_presentation_content for slideIndex 0 to confirm the text was added.",
            max_turns=6,
        )

        assert result.success
        assert result.tool_was_called("set_presentation_content")
        assert result.tool_was_called("get_presentation_content")

    async def test_adds_a_table_via_slide_code(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-add-table",
            allowed_tools=["get_presentation_overview", "add_slide_from_code"],
        )

        result = await eval_run(
            agent,
            "First call get_presentation_overview. Then use add_slide_from_code to add a new slide containing a "
            "table with columns (Month, Revenue) and three rows of data, using a PptxGenJS slide.addTable call.",
            max_turns=8,
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("add_slide_from_code")

    async def test_replaces_existing_slide_via_code(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-replace-slide",
            allowed_tools=["get_presentation_overview", "add_slide_from_code"],
        )

        result = await eval_run(
            agent,
            "Call get_presentation_overview first. Then use add_slide_from_code with replaceSlideIndex 0 "
            "to replace the first slide with one containing the title text 'Updated Cover'.",
            max_turns=8,
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("add_slide_from_code")

    async def test_fetches_image_then_embeds_in_slide(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-fetch-and-embed",
            allowed_tools=["fetch_image_as_base64", "add_slide_from_code"],
        )

        result = await eval_run(
            agent,
            "Fetch the image at URL 'https://example.com/logo.png' using fetch_image_as_base64. "
            "Then pass the returned base64 data URI into add_slide_from_code to embed the image in a new slide.",
            max_turns=8,
        )

        assert result.success
        assert result.tool_was_called("fetch_image_as_base64")
        assert result.tool_was_called("add_slide_from_code")

    async def test_adds_multiple_shapes_to_slide(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-multiple-shapes",
            allowed_tools=["get_presentation_overview", "add_geometric_shape"],
            max_turns=8,
        )

        result = await eval_run(
            agent,
            "Inspect the deck first. Then add a red ellipse and a blue rectangle to slide 0, "
            "each using a separate call to add_geometric_shape.",
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("add_geometric_shape")


# =============================================================================
# PowerPoint — Adversarial
# =============================================================================


class TestPowerPointAdversarial:
    """Edge cases and error-handling scenarios for PowerPoint tools."""

    async def test_out_of_bounds_slide_index_returns_error(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-adversarial-oob-index",
            allowed_tools=["get_presentation_overview", "get_presentation_content"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Call get_presentation_overview to see the slide count. Then call get_presentation_content "
            "with slideIndex 999 (which is out of range) and report the error returned.",
        )

        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("get_presentation_content")

    async def test_overview_called_before_modifying_slide(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-adversarial-read-before-write",
            allowed_tools=["get_presentation_overview", "add_geometric_shape"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Add a blue triangle to slide 0. You must call get_presentation_overview first to verify "
            "the deck has at least one slide before modifying it.",
        )

        tool_calls = [call.name for call in result.all_tool_calls]
        assert result.success
        assert result.tool_was_called("get_presentation_overview")
        assert result.tool_was_called("add_geometric_shape")
        assert tool_calls.index("get_presentation_overview") < tool_calls.index("add_geometric_shape")

    async def test_negative_slide_index_is_gracefully_handled(self, eval_run, powerpoint_server):
        agent = _make_eval(
            powerpoint_server,
            "powerpoint-adversarial-negative-index",
            allowed_tools=["set_presentation_content", "get_presentation_content"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Try to add text 'Test' to slideIndex -1 using set_presentation_content. "
            "Report the error returned, then call get_presentation_content without any index "
            "to confirm the existing slides are still readable.",
        )

        assert result.success
        assert result.tool_was_called("set_presentation_content")
        assert result.tool_was_called("get_presentation_content")
