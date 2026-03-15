"""Word MCP server integration tests using pytest-skill-engineering."""

from __future__ import annotations

import pytest

from pytest_skill_engineering import Eval, MCPServer, Provider

from conftest import DEFAULT_MAX_TURNS, DEFAULT_MODEL, DEFAULT_RPM, DEFAULT_TPM, SYSTEM_PROMPTS

pytestmark = [pytest.mark.integration, pytest.mark.word]

WORD_PROMPT = SYSTEM_PROMPTS["word"]


def _make_eval(
    word_server: MCPServer,
    name: str,
    *,
    allowed_tools: list[str] | None = None,
    max_turns: int = DEFAULT_MAX_TURNS,
) -> Eval:
    return Eval(
        name=name,
        provider=Provider(model=f"azure/{DEFAULT_MODEL}", rpm=DEFAULT_RPM, tpm=DEFAULT_TPM),
        mcp_servers=[word_server],
        system_prompt=WORD_PROMPT,
        max_turns=max_turns,
        allowed_tools=allowed_tools,
    )


class TestWordToolSelection:
    """Validate that models can choose the right Word tools."""

    async def test_reads_document_content(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-read-content",
            allowed_tools=["get_document_overview", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Inspect the active document, then read the document content so you can summarize it.",
        )

        assert result.success
        assert result.tool_was_called("get_document_overview")
        assert result.tool_was_called("get_document_content")

    async def test_inserts_text_and_reads_it_back(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-insert-text",
            allowed_tools=[
                "get_document_overview",
                "insert_content_at_selection",
                "get_document_content",
            ],
        )

        result = await eval_run(
            agent,
            "Review the document first. Then insert the text '<p>Launch checklist complete.</p>' at the current selection and read back the document content.",
        )

        assert result.success
        assert result.tool_was_called("get_document_overview")
        assert result.tool_was_called("insert_content_at_selection")
        assert result.tool_was_called("get_document_content")

    async def test_formats_inserted_text(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-format-text",
            allowed_tools=[
                "get_document_overview",
                "insert_content_at_selection",
                "apply_style_to_selection",
            ],
        )

        result = await eval_run(
            agent,
            "Inspect the document, insert '<p>Ready for release.</p>' at the selection, then make that inserted text bold.",
        )

        assert result.success
        assert result.tool_was_called("insert_content_at_selection")
        assert result.tool_was_called("apply_style_to_selection")

    async def test_inserts_a_table(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-insert-table",
            allowed_tools=["get_document_overview", "insert_table", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Review the document, then insert a 2 by 2 table with the header row [['Task', 'Owner'], ['Review', 'Mark']] and confirm the document content.",
        )

        assert result.success
        assert result.tool_was_called("get_document_overview")
        assert result.tool_was_called("insert_table")
        assert result.tool_was_called("get_document_content")

    async def test_replaces_text_in_the_document(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-replace-text",
            allowed_tools=["set_document_content", "find_and_replace", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Replace the whole document with '<p>Draft agenda draft agenda</p>', then replace every occurrence of the word draft with final and read back the document.",
        )

        assert result.success
        assert result.tool_was_called("set_document_content")
        assert result.tool_was_called("find_and_replace")
        assert result.tool_was_called("get_document_content")
