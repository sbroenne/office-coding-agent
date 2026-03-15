"""Outlook MCP server integration tests using pytest-skill-engineering."""

from __future__ import annotations

import pytest

from pytest_skill_engineering import Eval, MCPServer, Provider

from conftest import DEFAULT_MAX_TURNS, DEFAULT_MODEL, DEFAULT_RPM, DEFAULT_TPM, SYSTEM_PROMPTS

pytestmark = [pytest.mark.integration, pytest.mark.outlook]

OUTLOOK_PROMPT = SYSTEM_PROMPTS["outlook"]


def _make_eval(
    outlook_server: MCPServer,
    name: str,
    *,
    allowed_tools: list[str] | None = None,
    max_turns: int = DEFAULT_MAX_TURNS,
) -> Eval:
    return Eval(
        name=name,
        provider=Provider(model=f"azure/{DEFAULT_MODEL}", rpm=DEFAULT_RPM, tpm=DEFAULT_TPM),
        mcp_servers=[outlook_server],
        system_prompt=OUTLOOK_PROMPT,
        max_turns=max_turns,
        allowed_tools=allowed_tools,
    )


class TestOutlookToolSelection:
    """Validate that models can choose the right Outlook tools."""

    async def test_reads_current_email_content(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-read-email",
            allowed_tools=["get_mail_item", "get_mail_body"],
        )

        result = await eval_run(
            agent,
            "Review the current email, then read the full message body so you can summarize it.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_item")
        assert result.tool_was_called("get_mail_body")

    async def test_composes_a_new_email(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-compose-email",
            allowed_tools=["display_new_message", "get_mail_item", "get_mail_body"],
        )

        result = await eval_run(
            agent,
            "Compose a new email to finance@example.com with the subject 'Budget review' and the HTML body '<p>Please review the draft budget.</p>'. Then inspect the draft.",
        )

        assert result.success
        assert result.tool_was_called("display_new_message")
        assert result.tool_was_called("get_mail_item")

    async def test_replies_to_the_current_email(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-reply-email",
            allowed_tools=["get_mail_item", "reply_to_mail", "get_mail_body"],
        )

        result = await eval_run(
            agent,
            "Read the current email, then draft a reply that says '<p>Thanks, I will review this today.</p>' and inspect the reply body.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_item")
        assert result.tool_was_called("reply_to_mail")
        assert result.tool_was_called("get_mail_body")
