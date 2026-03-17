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


# =============================================================================
# Document Structure — Inspection
# =============================================================================


class TestDocumentStructure:
    """Validate that models can inspect document structure and metadata."""

    async def test_reads_document_overview_only(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-overview-only",
            allowed_tools=["get_document_overview"],
        )

        result = await eval_run(
            agent,
            "Call get_document_overview and tell me the paragraph count and any headings you find.",
        )

        assert result.success
        assert result.tool_was_called("get_document_overview")

    async def test_sets_and_reads_document_content(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-set-then-read",
            allowed_tools=["set_document_content", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Replace the entire document with '<h1>Status Report</h1><p>All systems nominal.</p>', "
            "then read the document content back to confirm.",
        )

        assert result.success
        assert result.tool_was_called("set_document_content")
        assert result.tool_was_called("get_document_content")

    async def test_overview_reflects_inserted_content(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-overview-after-insert",
            allowed_tools=["insert_content_at_selection", "get_document_overview"],
        )

        result = await eval_run(
            agent,
            "Insert '<p>New section added.</p>' at the selection. "
            "Then call get_document_overview to show the updated paragraph count.",
        )

        assert result.success
        assert result.tool_was_called("insert_content_at_selection")
        assert result.tool_was_called("get_document_overview")


# =============================================================================
# Formatting Operations
# =============================================================================


class TestFormattingOperations:
    """Validate that models apply the correct formatting tools."""

    async def test_applies_italic_style(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-italic-style",
            allowed_tools=["insert_content_at_selection", "apply_style_to_selection"],
        )

        result = await eval_run(
            agent,
            "Insert '<p>Important note.</p>' at the selection, then apply italic formatting to the current selection.",
        )

        assert result.success
        assert result.tool_was_called("insert_content_at_selection")
        assert result.tool_was_called("apply_style_to_selection")

    async def test_applies_font_color(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-font-color",
            allowed_tools=["insert_content_at_selection", "apply_style_to_selection"],
        )

        result = await eval_run(
            agent,
            "Insert '<p>Warning: deadline approaching.</p>' at the selection. "
            "Then apply a red font color to the selection using apply_style_to_selection.",
        )

        assert result.success
        assert result.tool_was_called("insert_content_at_selection")
        assert result.tool_was_called("apply_style_to_selection")

    async def test_applies_underline_style(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-underline-style",
            allowed_tools=["insert_content_at_selection", "apply_style_to_selection"],
        )

        result = await eval_run(
            agent,
            "Insert '<p>Section heading text.</p>' at the selection, "
            "then apply underline formatting to the selection.",
        )

        assert result.success
        assert result.tool_was_called("insert_content_at_selection")
        assert result.tool_was_called("apply_style_to_selection")


# =============================================================================
# Find & Replace Variants
# =============================================================================


class TestFindAndReplaceVariants:
    """Extended find-and-replace coverage."""

    async def test_find_and_replace_case_sensitive(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-replace-case-sensitive",
            allowed_tools=["set_document_content", "find_and_replace", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Set the document to '<p>Draft DRAFT draft</p>'. Then do a case-sensitive find-and-replace "
            "to replace only 'Draft' (capitalised) with 'Version', then read back the result.",
        )

        assert result.success
        assert result.tool_was_called("set_document_content")
        assert result.tool_was_called("find_and_replace")
        assert result.tool_was_called("get_document_content")

    async def test_find_and_replace_multiple_occurrences(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-replace-multiple",
            allowed_tools=["set_document_content", "find_and_replace", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Replace the document with '<p>alpha beta alpha gamma alpha</p>'. "
            "Then replace every occurrence of 'alpha' with 'omega' and read back the document.",
        )

        assert result.success
        assert result.tool_was_called("set_document_content")
        assert result.tool_was_called("find_and_replace")
        assert result.tool_was_called("get_document_content")


# =============================================================================
# Table Operations
# =============================================================================


class TestTableOperations:
    """Validate table insertion variants."""

    async def test_inserts_table_with_header_row(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-table-header-row",
            allowed_tools=["get_document_overview", "insert_table", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Review the document, then insert a 3-column table with a header row: "
            "[['Sprint', 'Status', 'Owner'], ['Sprint 1', 'Done', 'Alice'], ['Sprint 2', 'In Progress', 'Bob']]. "
            "Confirm the result by reading back the document.",
        )

        assert result.success
        assert result.tool_was_called("get_document_overview")
        assert result.tool_was_called("insert_table")
        assert result.tool_was_called("get_document_content")

    async def test_inserts_empty_table(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-empty-table",
            allowed_tools=["insert_table", "get_document_content"],
        )

        result = await eval_run(
            agent,
            "Insert a blank 4-row by 3-column table into the document, then read the document to confirm.",
        )

        assert result.success
        assert result.tool_was_called("insert_table")
        assert result.tool_was_called("get_document_content")


# =============================================================================
# Word — Adversarial
# =============================================================================


class TestWordAdversarial:
    """Edge cases and error-handling scenarios for Word tools."""

    async def test_overview_called_before_inserting_content(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-adversarial-read-before-insert",
            allowed_tools=["get_document_overview", "insert_content_at_selection"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Before inserting anything, call get_document_overview to understand the current document state. "
            "Then insert '<p>Appended paragraph.</p>' at the selection.",
        )

        tool_calls = [call.name for call in result.all_tool_calls]
        assert result.success
        assert result.tool_was_called("get_document_overview")
        assert result.tool_was_called("insert_content_at_selection")
        assert tool_calls.index("get_document_overview") < tool_calls.index("insert_content_at_selection")

    async def test_replace_nonexistent_text_reports_zero_count(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-adversarial-no-match",
            allowed_tools=["find_and_replace", "get_document_content"],
            max_turns=6,
        )

        result = await eval_run(
            agent,
            "Use find_and_replace to replace 'xyzzy_does_not_exist' with 'found'. "
            "Report how many replacements were made, then read the document content to confirm it is unchanged.",
        )

        assert result.success
        assert result.tool_was_called("find_and_replace")
        assert result.tool_was_called("get_document_content")

    async def test_chain_insert_then_replace(self, eval_run, word_server):
        agent = _make_eval(
            word_server,
            "word-adversarial-insert-then-replace",
            allowed_tools=["insert_content_at_selection", "find_and_replace", "get_document_content"],
            max_turns=8,
        )

        result = await eval_run(
            agent,
            "Insert '<p>The quick brown fox jumps.</p>' at the selection. "
            "Then replace 'quick' with 'slow'. Finally read the document content back to verify the change.",
        )

        assert result.success
        assert result.tool_was_called("insert_content_at_selection")
        assert result.tool_was_called("find_and_replace")
        assert result.tool_was_called("get_document_content")
