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
