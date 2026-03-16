# Squad Decisions

## Active Decisions

### User Directive: Python Testing Framework & Expansion (2026-03-15T10:27Z)
**Owner:** Stefan Broenner (via Copilot)  
**Status:** Captured  
Keep the Python test project (pyproject.toml + tests-aitest/). Switch to `pytest-skill-engineering` as the testing framework. Add tests for ALL tools (PowerPoint, Word, Outlook, not just Excel). Document in README.

### Harmony: Squad v0.8.25 Reference Templates (2026-03-15)
**Owner:** Harmony  
**Status:** Completed  
Created missing v0.8.25 reference templates and plugin marketplace state file:
- `.squad/templates/ralph-reference.md`, `casting-reference.md`, `copilot-agent.md`, `human-members.md`, `issue-lifecycle.md`, `prd-intake.md`, `ceremony-reference.md`
- `.squad/plugins/marketplaces.json`

**Why:** The coordinator governance file (`.github/agents/squad.agent.md`) referenced these files as on-demand detail expansions but they did not exist. Restores lookup paths and ensures plugin marketplace has a defined empty-state file.

### Irving: Adaptive Cards MCP Integration (2026-03-15)
**Owner:** Irving  
**Status:** Completed  
Added `adaptive-cards-mcp` to `.copilot/mcp-config.json` as a stdio MCP server via `npx -y adaptive-cards-mcp`. Exposes seven tools: `generate_card`, `validate_card`, `data_to_card`, `optimize_card`, `template_card`, `transform_card`, `suggest_layout`.

**Why:** Lightweight way to generate and validate Adaptive Card payloads for Office workflows. Best fit for Outlook actionable messages and Excel-to-card data visualization flows.

**Note:** npm 404 for package on current environment; follow-up may be needed to publish or point config at different distributable.

### Irving: Upgrade @github/copilot to 1.0.5 (2026-03-15)
**Owner:** Irving  
**Status:** Completed  
Upgraded critical security packages:
- `@github/copilot`: ^0.0.414 → ^1.0.5 (major version bump)
- `@github/copilot-sdk`: ^0.1.30 → ^0.1.32

**Why:** GHSA-g8r9-g2v8-jv6f vulnerability (arbitrary code execution via shell expansion). No breaking API changes affecting codebase. Build, typecheck, and integration tests all pass.

**Impact:** None required beyond package.json/package-lock.json. Fully backward compatible.

### Irving: ESLint 10 Vitest Plugin (2026-03-15)
**Owner:** Irving  
**Status:** Completed  
Switched from `eslint-plugin-vitest` (fails under ESLint 10) to `@vitest/eslint-plugin` in `eslint.config.mjs`.

**Why:** `eslint-plugin-vitest` 0.5.4 only supports ESLint 8/9. `@vitest/eslint-plugin` declares support for ESLint >=8.57.0, covers ESLint 10, preserves flat-config usage.

**Future:** Treat `@vitest/eslint-plugin` as the supported Vitest ESLint integration going forward.

### Irving: Manifest-Backed Excel MCP Descriptions (2026-03-15)
**Owner:** Irving  
**Status:** Proposed  
Use `tests-aitest/manifests/excel-tools-manifest.json` as the source of truth for Excel AI eval tool descriptions. `tests-aitest/excel_mcp.py` maps decomposed tool names back to aggregate manifest tools and reuses parameter descriptions.

**Why:** Keeps eval MCP aligned with product tool definitions. Avoids second, lower-quality description system in Python. Makes decomposed tools self-describing by attaching the conditional-format or validation type.

### Irving: Standardize Python Evals on pytest-skill-engineering (2026-03-15)
**Owner:** Irving  
**Status:** Proposed  
Move `tests-aitest/` from `pytest-aitest` to `pytest-skill-engineering`. Remove `pytest-llm-assert` dependency (now built in).

**Why:** `pytest-skill-engineering` is current framework for MCP/prompt/skill evals. Preserves MCPServer/Provider/Wait concepts; renames Agent→Eval and aitest_run→eval_run. Consolidates Python eval dependencies.

### Mark: Adversarial Excel AI Evals (2026-03-15)
**Owner:** Mark  
**Status:** Proposed  
Added `tests-aitest/test_excel_adversarial.py` with `pytest.mark.adversarial` marker. Decision: keep adversarial evals strict when wrong tool materially changes workbook behavior.

**Why:** Live execution showed prompt "Highlight values above 100 but don't restrict input" still called both `add_cell_value_format` AND `set_number_validation`, changing user behavior in the sheet. Real product-quality issue, not just prompt quirk.

**Recommendation:** Treat confusion tests as behavior tests. Assert correct tool called AND wrong tool not called when side effects differ. Keep `allowed_tools` narrow to identify genuine routing problems.

### Mark: AI Test Manifest Strategy (2026-03-15)
**Owner:** Mark  
**Status:** Proposed  
Use per-host manifest files under `tests-aitest/manifests/` instead of single shared manifest:
- **Excel:** Generated from decomposed config arrays in `src/tools/configs/` (aligns with legacy simulator and test expectations for names like `set_range_values`, `add_color_scale`)
- **PowerPoint/Word/Outlook:** Generated from runtime `Tool[]` objects (no config-array equivalent)

**Why:** Keeps eval schema aligned with real host surfaces. Preserves existing Excel simulator routing. Enables new host suites without handwritten manifests.

**Follow-up:** If project standardizes all hosts on single manifest shape, revisit Excel simulator to move off decomposed config manifest onto runtime `Tool[]` surface.

### Harmony: Host Registry / Plugin Architecture (2026-03-16)
**Owner:** Harmony  
**Status:** Proposed  
Consolidate per-host configuration into a single registry centralizing:
- prompt sources
- tool definitions
- agent defaults
- optional orchestrators
- capability metadata

**Why:** Current architecture has proxy boundary too wide, host fallback leaks Excel assumptions, tool definitions split three ways (Excel config, PowerPoint/Word inline, Outlook handwritten), and `useOfficeChat.ts` carries excessive responsibilities. Centralized registry would shrink `useOfficeChat.ts`, remove scattered host switches, and unify host targeting including Outlook.

### Irving: Lock Down Localhost APIs & WebSocket (2026-03-16)
**Owner:** Irving  
**Status:** Proposed  
Security hardening roadmap:
1. Remove permissive CORS from `src/server.mjs`; add Origin validation to WebSocket upgrades
2. Fix prefix bypass vulnerability in `/api/browse` path checks
3. Add per-tool execution deadlines to prevent blocking tool calls
4. Improve WebSocket lifecycle with complete reconnect strategy and buffer management

**Why:** Current implementation has permissive `cors({ origin: '*' })`, unauthenticated API endpoints (`/api/env`, `/api/browse`), and no tool execution timeouts. `src/copilotProxy.mjs` accepts upgrades without Origin validation.

### Irving: Harden WebSocket Lifecycle & Failure Containment (2026-03-16)
**Owner:** Irving  
**Status:** Proposed  
Improve WebSocket transport:
- Add timeout and failure containment around `sendRequest('tool.call', ...)` to prevent indefinite blocking
- Harden `src/lib/websocket-client.ts` reconnect strategy
- Fix buffer ceiling and malformed header handling in `src/lib/websocket-transport.ts`
- Handle idle disconnect scenarios better (currently under-handled outside `useOfficeChat`)

**Why:** Hanging browser tool handlers can block proxy indefinitely. WebSocket lifecycle is incomplete; recovery mostly happens during active send failures.

### Dylan: Refactor ModelPicker Session Lifecycle (2026-03-16)
**Owner:** Dylan  
**Status:** Proposed  
ModelPicker currently calls `useOfficeChat(host)` directly, creating a duplicate session outside the main app tree. Refactor to:
- Consume session context from App instead of creating its own WebSocket connection
- Eliminate risk of WebSocket churn, stale messages, and model switching against wrong session

**Why:** Current implementation bypasses the main session lifecycle, risking out-of-sync state and duplicate socket connections.

### Dylan: Audit & Harmonize VS Code Design Tokens (2026-03-16)
**Owner:** Dylan  
**Status:** Proposed  
UI fidelity hardening:
1. Remove `lucide-react` from App-level banners and loading states
2. Replace all color values with `--vscode-*` CSS custom properties
3. Replace all icons with codicons (`@vscode/codicons`)
4. Declare and validate all referenced tokens (e.g., `--vscode-icon-foreground`, `--vscode-badge-*`, `--vscode-keybindingLabel-*`, `--vscode-quickInput-background`, `--vscode-widget-shadow`)

**Why:** App-level UI breaks VS Code Copilot Chat fidelity. Components reference undeclared tokens. This is the clearest visual break from the design system.

### Dylan: Remove Dead Control Handlers & Improve Accessibility (2026-03-16)
**Owner:** Dylan  
**Status:** Proposed  
UI quality improvements:
1. Either wire `regenerate` and `feedback` button handlers in `ActionBar` or remove the buttons from `ChatPanel`
2. Make `UserMessage` edit affordance keyboard accessible (currently mouse-centric)
3. Add combobox semantics to slash command listbox (`aria-controls`, `aria-activedescendant`, option IDs)
4. Enhance `ChatHeader` with Copilot-style title/header treatment

**Why:** Exposing dead controls creates confusion. Accessibility gaps limit usability for keyboard-only and assistive-device users.

### Mark: Expand Non-Excel Host E2E Coverage (2026-03-16)
**Owner:** Mark  
**Status:** Proposed  
Extend E2E test depth for non-Excel hosts:
- **PowerPoint:** Cover 37 named tools, not just 13; prioritize write/edit/shape/layout flows
- **Word:** Cover 35 named tools, not just 12; prioritize selection awareness and editing workflows
- **Outlook:** Cover 22 tools and real workflows (reply chains, calendar lookups, multi-recipient handling, advanced search, categories, flags, archives), not just shallow read operations

**Why:** Non-Excel hosts have significantly shallower coverage (PowerPoint 13/37, Word 12/35, Outlook 6/22) compared to Excel (233 tests). Real workflows are underrepresented.

### Mark: Split Playwright Tests into Smoke vs Live E2E (2026-03-16)
**Owner:** Mark  
**Status:** Proposed  
Clarify test boundaries:
- `tests-ui/fixtures.ts` currently mocks WebSocket, uses synthetic JSON-RPC responses, and pre-seeds state
- This directly violates the repo rule that Playwright should not mock network/WebSocket/server behavior
- Solution: Split into two suites — smoke tests (mocked, Vitest) and true E2E (live server, Playwright)

**Why:** Current Playwright tests contradict their stated purpose of verifying end-to-end flows through the real Copilot API.

### Mark: Add Real-Path Coverage for Critical Chat Flows (2026-03-16)
**Owner:** Mark  
**Status:** Proposed  
Critical orchestration flows lack end-to-end coverage:
- Permission approval workflows
- MCP tool registration and execution
- Reconnect and session recovery
- Slash-command-to-plugin execution

Current coverage is mostly mocked in Vitest (`use-office-chat.test.tsx`, `management-tools.test.ts`, `chat-composer-slash.test.tsx`). Add one real-path test per flow before expanding long-tail component coverage.

**Why:** These flows are user-critical and currently untested in live conditions.

### Mark: Clean Up Test Artifacts & Unify vitest Config (2026-03-16)
**Owner:** Mark  
**Status:** Proposed  
Test infrastructure improvements:
1. Remove generated artifacts from `tests-e2e/`, `tests-aitest/` (dist, node_modules, .pycache, report.html)
2. Add to `.gitignore` to prevent future commits
3. Remove `unit` project from `vitest.config.ts` (conflicts with repo policy against unit tests)
4. Consolidate `tests/` structure (currently overlaps between `unit` and `integration`)

**Why:** Test directories currently contain noise that makes review harder and can hide real test changes.

### Ellis: Strengthen Outlook Host Support (2026-03-16)
**Owner:** Ellis  
**Status:** Proposed  
Outlook expansion (2–3 sprint effort):
1. Expand Outlook agent AGENT.md to detail level matching PowerPoint/Word (include use-case narratives)
2. Add 50+ E2E tests covering:
   - Reply chains and threading
   - Calendar lookups for scheduling
   - Multi-recipient handling (To/Cc/Bcc)
   - Advanced search + filters
   - Categories, flags, archives
3. Ensure all 22 Outlook tools have integration test coverage
4. Add Outlook AI eval tests to catch LLM confusion on email semantics

**Success Metric:** Outlook E2E tests ≥ 50; agent quality matches PowerPoint.

**Why:** Outlook is the thinnest host (8 tests vs 233 Excel). Agent definition is skeletal. Real workflows are underrepresented.

### Ellis: Add Welcome Screen for Onboarding (2026-03-16)
**Owner:** Ellis  
**Status:** Proposed  
New user onboarding (1 sprint effort):
1. Design welcome screen (Copilot-style, centered, friendly) with:
   - Host-specific prompt suggestions
   - Link to "Learn more" → host-specific docs
   - Dismiss toggle (persist welcome state)
2. Add "?" help button in header → links to agent instructions + docs
3. Host-specific guidance:
   - **Excel:** "Read values from A1:B10" / "Create a pivot table"
   - **PowerPoint:** "Create a 5-slide presentation on…" / "Redesign this slide"
   - **Word:** "Summarize this document" / "Format this heading"
   - **Outlook:** "Draft a reply to this email" / "Find my meetings next week"

**Success Metric:** New users convert (don't abandon at blank chat).

**Why:** Add-in ships with blank chat. No prompt examples, no host-specific guidance, no quick-start help. Most critical UX gap.

### Ellis: Create Developer Documentation Suite (2026-03-16)
**Owner:** Ellis  
**Status:** Proposed  
New extensibility guides (1 sprint effort):
1. `docs/TOOL_API_REFERENCE.md` — auto-generated from `src/tools/` configs, per-host tooling with parameters, return types, examples
2. `docs/SKILL_DEVELOPMENT.md` — example skill, YAML frontmatter spec, local testing
3. `docs/AGENT_DEVELOPMENT.md` — agent template, host declaration, custom agent testing
4. `docs/PLUGIN_DEVELOPMENT.md` — full plugin structure, build/publish workflow, link to CLI spec
5. Link all four guides from README.md under new "Extending" section

**Success Metric:** Third-party plugins submitted; reduced support burden for "how do I add a tool?" questions.

**Why:** Developers who want to extend the add-in (plugins, skills, agents) have no documented starting points. Tool API exists only in code.

### Dylan: ModelPicker Session Ownership Fix (2026-03-16T121900Z)
**Owner:** Dylan  
**Status:** In Progress (PR #144)  
`ModelPicker` should not call `useOfficeChat()` directly. It remains a presentational picker backed by `useSettingsStore` for model list/current selection, while `ChatPanel` owns the session context and passes `hasActiveSession` plus `onSwitchModel` as props.

**Why:** Mounting `useOfficeChat()` inside the picker created a second WebSocket/session lifecycle that could drift from the visible conversation. Passing session-aware props keeps model switching aligned with the active session while preserving no-session behavior (store-only model updates for the next conversation).

**UI Note:** App-level connection, error, and permission banners were also normalized to codicons and `--vscode-*` tokens so the task pane matches VS Code Copilot Chat more closely.

### Mark: Remove Playwright WebSocket Mocking (2026-03-16T121900Z)
**Owner:** Mark  
**Status:** In Progress (PR #143)  
Remove all Playwright WebSocket mocking helpers from `tests-ui/fixtures.ts`. Keep only real-session fixtures (`taskpane`, `configuredTaskpane`) and a lightweight live-server availability helper. Rewrite mocked connection/tool-card specs to use the live proxy path.

**Why:** Project policy says Playwright tests in `tests-ui/` must verify real end-to-end behavior through the live proxy and Copilot API. The prior fixture layer used `page.routeWebSocket` plus synthetic JSON-RPC events, which invalidated that coverage boundary.

**Follow-up:** Playwright now requires `npm run dev` to be running. Offline/error-path coverage should remain in integration tests unless reproduced without mocks.

### Irving: Lock Down Localhost APIs & WebSocket (2026-03-16T121900Z)
**Owner:** Irving  
**Status:** In Progress (PR #145)  
Add Origin validation to WebSocket upgrades, per-tool execution deadlines, and close permissive CORS. Fix prefix-bypass vulnerability in `/api/browse` path checks.

**Why:** Current implementation has permissive `cors({ origin: '*' })`, unauthenticated API endpoints, and no tool execution timeouts. Hanging browser tool handlers can block proxy indefinitely.

**Impact:** Transparent to browser clients; improves proxy isolation and resilience.

## Governance

- All meaningful changes require team consensus
- Document architectural decisions here
- Keep history focused on work, decisions focused on direction
