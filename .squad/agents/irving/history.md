# Project Context

- **Owner:** Stefan Broenner
- **Project:** office-coding-agent — Microsoft Office add-in bringing GitHub Copilot into Excel, PowerPoint, Word, and Outlook with full Copilot CLI plugin support
- **Stack:** React 18, TypeScript, Node.js, Tailwind CSS v4, Copilot SDK, WebSocket + JSON-RPC, Vite 7, Vitest, Playwright, Mocha (E2E)
- **Created:** 2026-03-15

## Learnings

### 2026-03-15 — @github/copilot 0.0.414 → 1.0.5 upgrade (GHSA-g8r9-g2v8-jv6f fix)

**What changed:**
- Upgraded `@github/copilot` from ^0.0.414 to ^1.0.5 (major version bump)
- Upgraded `@github/copilot-sdk` from ^0.1.30 to ^0.1.32 (compatible update)
- This addresses critical security vulnerability GHSA-g8r9-g2v8-jv6f (Dangerous Shell Expansion Patterns Enable Arbitrary Code Execution)

**Breaking changes:** None detected. This is a clean upgrade path:
- All imports are from `@github/copilot-sdk` (not the base package), so no import path changes needed
- The SDK API surface is stable between these versions
- Build passes (`npm run build:dev` — clean, no errors)
- TypeScript type checking passes (`npm run typecheck` — zero errors)
- Integration tests pass except for live WebSocket tests that require `npm run dev` running (expected behavior)

**Security posture:**
- The critical GHSA-g8r9-g2v8-jv6f vulnerability is no longer present in `npm audit` output
- Remaining vulnerabilities are in dev dependencies (mocha, electron, flatted, undici, tar) — not runtime blockers

**Recommendation:** This upgrade is ready to commit. No code changes required beyond package.json/package-lock.json.

### 2026-03-15 — Python AI evals migrated to pytest-skill-engineering

**What changed:**
- Swapped the Python eval project dependency from `pytest-aitest` to `pytest-skill-engineering`
- Removed the standalone `pytest-llm-assert` dependency because `llm_assert` now ships with the new framework
- Updated `tests-aitest` imports and fixture usage from `Agent`/`aitest_run` to `Eval`/`eval_run`
- Regenerated `uv.lock` and verified the suite still collects with `uv run pytest tests-aitest/ --collect-only`

**Why it matters:**
- The project now aligns with Stefan's newer skill-engineering eval stack instead of the older AI test plugin
- The migration keeps the existing MCP-backed Excel eval structure intact while using the renamed API surface
- Collection succeeded without running paid evals, so the migration is validated at the wiring level before any live model spend

### 2026-03-15 — Excel AI eval descriptions now come from the aggregate manifest

**What changed:**
- `tests-aitest/excel_mcp.py` now loads `tests-aitest/manifests/excel-tools-manifest.json` at startup and maps decomposed tool names back to aggregate manifest tool/action metadata.
- Tool docs now use manifest-backed descriptions, and parameter docs reuse manifest descriptions with targeted aliases like `searchValue -> searchText`, `formatCode -> format`, `newText -> text`, and `values -> filterValues`.
- `tests-aitest/conftest.py` now sets `DEFAULT_MAX_TURNS = 10` so multi-step evals have enough request budget to call tools and respond.

**Patterns / paths to remember:**
- Treat `tests-aitest/manifests/excel-tools-manifest.json` as the description source of truth for Excel AI eval tools; avoid inventing docs from method names when the manifest already has richer copy.
- The decomposed conditional-format and data-validation tools in `tests-aitest/excel_mcp.py` encode manifest `type` values in the tool name, so descriptions should call out the fixed rule/validation type.
- Fast verification path for this area: `uv run python tests-aitest/excel_mcp.py --help`, `uv run pytest tests-aitest/test_excel_tools.py::TestRangeOperations::test_write_and_read_range -v -x`, then `uv run pytest tests-aitest/ --collect-only`.
