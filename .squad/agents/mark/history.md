# Project Context

- **Owner:** Stefan Broenner
- **Project:** office-coding-agent — Microsoft Office add-in bringing GitHub Copilot into Excel, PowerPoint, Word, and Outlook with full Copilot CLI plugin support
- **Stack:** React 18, TypeScript, Node.js, Tailwind CSS v4, Copilot SDK, WebSocket + JSON-RPC, Vite 7, Vitest, Playwright, Mocha (E2E)
- **Created:** 2026-03-15

## Learnings

<!-- Append new learnings below. Each entry is something lasting about the project. -->

- 2026-03-15: `tests-aitest/` now uses `pytest-skill-engineering` end to end. In this repo, `uv sync` is required after switching from `pytest-aitest` or pytest will autoload the stale `pytest_aitest` plugin and force HTML/AI analysis with the old `aitest_run` fixture behavior.
- 2026-03-15: AI eval manifests now live under `tests-aitest/manifests/` and are host-specific. Excel still needs a decomposed config-driven manifest so legacy Excel simulator routes keep working, while PowerPoint, Word, and Outlook use manifests generated from runtime `Tool[]` definitions.
- 2026-03-15: The Azure OpenAI eval resource `stbrnner1` is in subscription `f036a9c9-6d6c-4d28-8d2c-3b68997cd99b` / tenant `16b3c013-d300-468d-ac64-7eda0820b6d3`. If evals start failing with `Tenant provided in token does not match resource tenant`, switch the Azure CLI default subscription back to that tenant before running `uv run pytest tests-aitest ...`.
- 2026-03-15: New host eval suites were added and validated: `test_powerpoint_tools.py`, `test_word_tools.py`, and `test_outlook_tools.py`. PowerPoint prompts need to be highly explicit for `add_slide_from_code` scenarios (provide chart values, concrete slide index, or direct PptxGenJS intent) to avoid turn-limit/time-out failures.
- 2026-03-15: Adversarial Excel evals now live in `tests-aitest/test_excel_adversarial.py` and use the same `_make_eval(...)` pattern as `tests-aitest/test_excel_tools.py`, plus local helpers for ordered-call assertions and wrong-tool absence checks. Register the suite with `pytest.mark.adversarial` in `tests-aitest/conftest.py` so targeted runs and report filtering stay consistent.
- 2026-03-15: Tool-confusion prompts are productive bug-finders in this repo. A prompt that explicitly said to highlight values without restricting input still led the model to call both `add_cell_value_format` and `set_number_validation`, so adversarial evals should keep strict negative assertions when the wrong tool would change workbook behavior.
- 2026-03-16: Integration coverage is broad (53 files) but the Vitest `unit` project still includes `tests/**/*.test.*`, so many `tests/integration/` files are eligible for duplicate execution unless callers explicitly scope to `--project integration`.
- 2026-03-16: Host E2E depth is uneven. Excel exercises the major config families end to end, but current runners only cover a subset of runtime tools in other hosts: PowerPoint 13 of 37 named tools, Word 12 of 35, and Outlook 6 of 22.
- 2026-03-16: Playwright coverage is mixed-quality because several UI tests rely on `page.routeWebSocket`, pre-seeded localStorage, and synthetic tool events in `tests-ui/fixtures.ts`, which conflicts with the repo rule that Playwright should verify real end-to-end behavior without mocks.
- 2026-03-16: `tests-ui/fixtures.ts` must stay free of `page.routeWebSocket` and synthetic JSON-RPC helpers. Playwright coverage in this repo is only valid when it exercises the live proxy/Copilot path; disconnected/error-path assertions belong in integration tests unless they can be reproduced without mocking.
- 2026-03-17: Expanded AI eval test counts — PowerPoint: +12 tests (5→17), Word: +13 tests (5→18), Outlook: +12 tests (3→15). New classes cover slide-reading, slide-writing, compose operations, formatting variants, find-and-replace variants, table operations, and adversarial error-handling cases per host.
- 2026-03-17: Tools skipped from eval coverage because the in-memory simulator has no backing method: `get_slide_image`, `get_slide_notes`, `delete_slide`, `move_slide`, `clear_slide`, `update_slide_shape` (PPT); `insert_paragraph`, `insert_break`, `apply_paragraph_style`, `set_paragraph_format`, `get_document_properties`, `insert_image`, `get_comments` (Word); `get_mail_attachments`, `get_attachment_content`, `forward_mail`, `get_user_profile`, `add_file_attachment`, `add_notification`, `remove_notification`, `save_draft`, `get_mail_headers`, `display_new_appointment`, `get_diagnostics` (Outlook). Simulator gaps are the primary blocker for further coverage expansion.
