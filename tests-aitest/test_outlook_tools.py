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


# =============================================================================
# Mail Item Inspection
# =============================================================================


class TestMailItemInspection:
    """Validate that models can inspect the active mail item."""

    async def test_reads_email_subject_from_overview(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-read-subject",
            allowed_tools=["get_mail_item"],
        )

        result = await eval_run(
            agent,
            "Call get_mail_item and report the subject line of the current email.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_item")

    async def test_reads_email_body_as_html(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-read-body-html",
            allowed_tools=["get_mail_body"],
        )

        result = await eval_run(
            agent,
            "Call get_mail_body with format 'html' and show me the full HTML body of the current email.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_body")

    async def test_reads_email_body_as_plain_text(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-read-body-text",
            allowed_tools=["get_mail_body"],
        )

        result = await eval_run(
            agent,
            "Call get_mail_body with format 'text' and summarize the plain-text content.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_body")

    async def test_reads_sender_before_replying(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-sender-then-reply",
            allowed_tools=["get_mail_item", "reply_to_mail"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "First read the current email using get_mail_item to find out who sent it. "
            "Then reply with '<p>Thank you for the update.</p>'.",
        )

        tool_calls = [call.name for call in result.all_tool_calls]
        assert result.success
        assert result.tool_was_called("get_mail_item")
        assert result.tool_was_called("reply_to_mail")
        assert tool_calls.index("get_mail_item") < tool_calls.index("reply_to_mail")


# =============================================================================
# Compose Operations
# =============================================================================


class TestComposeOperations:
    """Validate compose-mode tools work correctly in sequence."""

    async def test_sets_subject_after_opening_compose(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-set-subject",
            allowed_tools=["display_new_message", "set_mail_subject", "get_mail_item"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Open a new blank compose form using display_new_message. "
            "Then set the subject to 'Action Required: Q3 Budget'. "
            "Finally call get_mail_item to verify the subject was set.",
        )

        assert result.success
        assert result.tool_was_called("display_new_message")
        assert result.tool_was_called("set_mail_subject")
        assert result.tool_was_called("get_mail_item")

    async def test_sets_body_after_opening_compose(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-set-body",
            allowed_tools=["display_new_message", "set_mail_body", "get_mail_body"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Open a new compose form using display_new_message. "
            "Then set the body to '<p>Please find the report attached.</p>' using set_mail_body. "
            "Read back the body with get_mail_body to confirm.",
        )

        assert result.success
        assert result.tool_was_called("display_new_message")
        assert result.tool_was_called("set_mail_body")
        assert result.tool_was_called("get_mail_body")

    async def test_adds_cc_recipient_to_compose(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-add-cc",
            allowed_tools=["display_new_message", "add_mail_recipient", "get_mail_item"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Open a new compose form to hr@example.com with subject 'Leave request'. "
            "Then add manager@example.com as a CC recipient using add_mail_recipient. "
            "Call get_mail_item to verify the recipients.",
        )

        assert result.success
        assert result.tool_was_called("display_new_message")
        assert result.tool_was_called("add_mail_recipient")
        assert result.tool_was_called("get_mail_item")

    async def test_reply_all_to_current_email(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-reply-all",
            allowed_tools=["get_mail_item", "reply_to_mail"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Read the current email with get_mail_item, then reply-all with "
            "'<p>Acknowledged. Will follow up by end of week.</p>' — set replyAll to true.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_item")
        assert result.tool_was_called("reply_to_mail")

    async def test_compose_full_email_with_all_fields(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-compose-full",
            allowed_tools=["display_new_message", "set_mail_subject", "set_mail_body", "get_mail_item"],
            max_turns=8,
        )

        result = await eval_run(
            agent,
            "Open a new compose form to team@example.com using display_new_message. "
            "Then set the subject to 'Weekly sync' using set_mail_subject. "
            "Then set the body to '<p>Agenda: review sprint progress and blockers.</p>' using set_mail_body. "
            "Finally call get_mail_item to confirm all fields.",
        )

        assert result.success
        assert result.tool_was_called("display_new_message")
        assert result.tool_was_called("set_mail_subject")
        assert result.tool_was_called("set_mail_body")
        assert result.tool_was_called("get_mail_item")


# =============================================================================
# Outlook — Adversarial
# =============================================================================


class TestOutlookAdversarial:
    """Edge cases and error-handling scenarios for Outlook tools."""

    async def test_set_subject_in_read_mode_returns_error(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-adversarial-set-subject-read-mode",
            allowed_tools=["get_mail_item", "set_mail_subject"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Without opening a compose form first, try to set the subject of the current email "
            "to 'Modified Subject' using set_mail_subject. Report the error you receive. "
            "Then call get_mail_item to confirm the original subject is unchanged.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_item")
        assert result.tool_was_called("set_mail_subject")

    async def test_set_body_in_read_mode_returns_error(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-adversarial-set-body-read-mode",
            allowed_tools=["get_mail_body", "set_mail_body"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Without opening a compose form first, try to set the email body to '<p>Overwritten</p>' "
            "using set_mail_body. Report the error, then call get_mail_body to verify the original body is intact.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_body")
        assert result.tool_was_called("set_mail_body")

    async def test_reads_item_before_composing_reply(self, eval_run, outlook_server):
        agent = _make_eval(
            outlook_server,
            "outlook-adversarial-inspect-then-reply",
            allowed_tools=["get_mail_item", "get_mail_body", "reply_to_mail"],
            max_turns=8,
        )

        result = await eval_run(
            agent,
            "Inspect the current email fully: call get_mail_item for the header info and get_mail_body "
            "for the text content. Then compose a contextual reply that references the email subject.",
        )

        assert result.success
        assert result.tool_was_called("get_mail_item")
        assert result.tool_was_called("get_mail_body")
        assert result.tool_was_called("reply_to_mail")
