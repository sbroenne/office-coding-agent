# Project Context

- **Owner:** Stefan Broenner
- **Project:** office-coding-agent — Microsoft Office add-in bringing GitHub Copilot into Excel, PowerPoint, Word, and Outlook with full Copilot CLI plugin support
- **Stack:** React 18, TypeScript, Node.js, Tailwind CSS v4, Copilot SDK, WebSocket + JSON-RPC, Vite 7, Vitest, Playwright, Mocha (E2E)
- **Created:** 2026-03-15

## Learnings

### 2026-03-16 — Localhost server hardening shipped

**What changed:**
- Added shared origin and browse-path guards in `src/serverSecurity.mjs` and used them in both `src/server.mjs` and `src/server-prod.mjs`.
- Replaced wildcard CORS with an allowlist for localhost plus trusted Microsoft/Office hosts, and reject unexpected `Origin` headers before `/api` routes run.
- Added WebSocket `Origin` validation in `src/copilotProxy.mjs` so `/api/copilot` now returns HTTP 403 for untrusted browser origins.
- Locked `/api/env` down to non-sensitive runtime metadata only, and removed the PermissionManager dependency on path-bearing env data.
- Hardened `/api/browse` with trusted-caller checks, canonical path checks, and explicit `..` traversal rejection while still allowing project/home browsing.
- Updated live WebSocket integration tests to send `Origin: https://localhost:3000`, which matches real browser behaviour after the new upgrade gate.

**Verification:**
- `npm run build:dev`
- `npm run test:integration`
- Manual probes confirmed 403 responses for untrusted REST origins, path traversal on `/api/browse`, and untrusted WebSocket upgrades.

### 2026-03-16 — Backend/server review findings

**Key findings:**
- `src/copilotProxy.mjs` is generally defensive on disconnect cleanup, but browser-routed `tool.call` requests have no timeout, so a hung browser tool can stall a session indefinitely.
- The browser WebSocket stack (`src/lib/websocket-client.ts` + `src/lib/websocket-transport.ts`) has no native reconnect loop; recovery mostly happens in `useOfficeChat` during send failures, so idle disconnects can leave stale clients behind until the next action.
- `src/server.mjs` exposes `GET /api/env`, `GET /api/browse`, and permissive `cors({ origin: '*' })` with no auth/origin check. Combined with the prefix-based path allowlist, localhost attack surface is broader than it should be.
- Tool configs are a strong source of truth, but `src/tools/codegen/factory.ts` does not enforce runtime argument validation or `additionalProperties: false`, so bad tool calls fall through to Office runtime errors instead of clean validation failures.
- `src/mcpClient.mjs` is isolated and not currently wired into the reviewed proxy path, but if revived it needs connection timeouts and guaranteed transport cleanup on partial startup failures.

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

### 2026-03-15 — npm dependency refresh across lint/build/test stack

**What changed:**
- Updated `lucide-react` to `^0.577.0`, `skillpm` to `^0.0.12`, `@eslint/js` to `^10.0.1`, `eslint` to `^10.0.3`, `vite` to `^8.0.0`, `@vitejs/plugin-react` to `^6.0.1`, `jsdom` to `^29.0.0`, and `electron` to `^41.0.2`.
- Replaced `eslint-plugin-vitest` with `@vitest/eslint-plugin` because the older plugin only supports ESLint 8/9 and breaks under ESLint 10.
- Fixed new ESLint 10 `no-useless-assignment` findings in `src/tools/powerpoint/index.ts` and `src/tools/word/index.ts` so the upgraded lint stack stays green.
- Verified the dependency refresh with `npm run build:dev`, `npm run lint`, `npm run test:integration`, `npm run test:e2e:ppt`, and `npm run test:e2e:word`.

**Patterns / paths to remember:**
- Keep Vitest lint rules wired through `eslint.config.mjs` using `@vitest/eslint-plugin` when the repo is on ESLint 10+.
- `vite.config.ts` and `vitest.config.ts` both remained compatible with Vite 8 / jsdom 29; dependency refreshes here are mostly package-level unless a future plugin introduces its own migration.
- Backend dependency touchpoints for this area are `package.json`, `package-lock.json`, `eslint.config.mjs`, and host tool implementations under `src/tools/`.
