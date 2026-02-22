# Copilot Instructions for office-coding-agent

## Project Overview

**office-coding-agent** is a Microsoft Office add-in with a single task pane UI and host-routed AI runtime behavior. The current implementation is **Excel-first**, but tools and prompts are selected by host (`excel`, `powerpoint`, etc.) to support future hosts without changing the UI.

## Key Technologies

- **React 18** + **assistant-ui** + **Radix UI** + **Tailwind CSS v4** — task pane UI (Thread, ToolFallback, Popovers)
- **GitHub Copilot SDK** (`@github/copilot-sdk`) — session management, streaming events, tool registration
- **WebSocket + JSON-RPC** — browser-to-proxy transport (`src/lib/websocket-client.ts`, `src/lib/websocket-transport.ts`)
- **Express + HTTPS** — local proxy server (`src/server.mjs`) that bridges WebSocket to the Copilot CLI
- **Zustand 5** — state management with persistence via `officeStorage` (OfficeRuntime.storage)
- **Vite 7** — bundling, dev server (HMR via middleware mode in Express)
- **TypeScript 5** — type safety
- **Vitest** — unit + integration testing (jsdom env)
- **Mocha** — E2E tests inside Excel Desktop (current host runtime E2E)
- **Playwright** — browser UI tests for task pane flows

## Architecture

The add-in routes messages through a **local proxy server** — the browser cannot call the Copilot API directly.

```
Browser task pane (React + assistant-ui)
         ↓ WebSocket (wss://localhost:3000/api/copilot)
Node.js proxy server  (src/server.mjs + src/copilotProxy.mjs)
         ↓ @github/copilot-sdk (manages CLI lifecycle internally)
GitHub Copilot API
```

The `useOfficeChat` hook creates a `WebSocketCopilotClient`, opens a `BrowserCopilotSession`, and maps incoming `SessionEvent` objects to `ThreadMessage[]` for assistant-ui via `useExternalStoreRuntime`.

### Agent System

The AI agent uses a split prompt architecture with host targeting:

- **`src/services/ai/BASE_PROMPT.md`** — universal base prompt (progress narration + presenting choices)
- **`src/services/ai/prompts/*_APP_PROMPT.md`** — host-level app prompt (Excel/PowerPoint)
- **`src/agents/*/AGENT.md`** — agent-specific instructions (default + custom)
- Instructions = `buildSystemPrompt(host) + resolvedAgent.instructions + skillContext`

The `agentService` (`src/services/agents/agentService.ts`) parses agent YAML frontmatter to `AgentConfig` objects and filters by host.

### Agent Frontmatter Contract (required)

Agents are targeted per host via frontmatter fields:

- `name`
- `description`
- `version`
- `hosts` (array; supported values: `excel`, `powerpoint`, `word`)
- `defaultForHosts` (array; subset of supported hosts)

Example:

```yaml
---
name: Excel
description: Default Excel agent
version: 1.1.0
hosts: [excel]
defaultForHosts: [excel]
---
```

Rules:

- If current host is not in `hosts`, the agent must not be shown/selected.
- Invalid/unknown host values are ignored.
- `resolveActiveAgent(activeAgentId, host)` should be used so host-default fallback is applied safely.

### Skills System

Bundled skill files in `src/skills/` provide additional context injected into the system prompt. Skills are **host-targeted** via an optional `hosts` field in their YAML frontmatter (same pattern as agents — empty `hosts` = available to all hosts). `buildSkillContext(activeNames?, host?)` filters skills by both active state and host compatibility. Skills can be toggled on/off via the SkillPicker. Active skills are stored as `activeSkillNames` in the settings store (`null` = all ON).

### The Host Runtime Boundary

**Critical concept:** everything that calls real Office host runtime APIs belongs below the runtime boundary.

```
┌──────────────────────────────────────────────────────┐
│  Testable with Vitest/Playwright (no Excel host)     │
│  ─────────────────────────────                       │
│  • Pure functions (parseFrontmatter,                 │
│    buildSkillContext, toolResultSummary, generateId,  │
│    humanizeToolName, zipImportService)               │
│  • Host routing (detectOfficeHost,                   │
│    getToolsForHost, buildSystemPrompt)               │
│  • Agent targeting and default resolution            │
│  • Zustand store logic (settingsStore)               │
│  • JSON Schema tool configs (toCopilotTools)         │
│  • React component wiring (integration)              │
│  • WebSocket client + session (mocked in unit tests) │
│  • Agent/skill service parsing                       │
├──────────────────────────────────────────────────────┤
│  Excel.run() boundary (current host implementation)  │
├──────────────────────────────────────────────────────┤
│  E2E only (Mocha + real Excel Desktop)               │
│  ─────────────────────────────                       │
│  • rangeCommands, tableCommands, sheetCommands       │
│  • chartCommands, workbookCommands, commentCommands  │
│  • conditionalFormatCommands, dataValidationCommands │
│  • pivotTableCommands                                │
│  • PowerPoint / Word commands                        │
│  • OfficeRuntime.storage (real runtime)              │
└──────────────────────────────────────────────────────┘
```

## UI Layout

The task pane is split into three areas:

- **ChatHeader** — "AI Chat" title + SkillPicker (icon-only with badge) + New Conversation button + Settings gear (SettingsDialog)
- **ChatPanel** — message list (Thread), Copilot-style progress indicators, choice cards, error bar, Composer, and an **input toolbar** below the input box with AgentPicker + ModelPicker (GitHub Copilot-style)
- **App** — owns settings dialog state, detects system theme and Office host; no setup wizard (Copilot CLI handles auth)

## Testing Strategy

> ### ⛔ CRITICAL RULE: DO NOT WRITE UNIT TESTS
>
> Unit tests that mock Office APIs or fabricate fake contexts provide zero confidence that code works in a real host. They test the mock, not the code.
> **Integration tests and E2E tests are the ONLY acceptable test forms for new functionality.**
>
> - Writing a unit test when an integration or E2E test is possible is forbidden.
> - If you are tempted to write a unit test, write an integration test instead.
> - If the feature touches Office APIs, write an E2E test.

### Test Tiers

| Tier            | Runner     | Directory            | Count | What it tests                                                                        |
| --------------- | ---------- | -------------------- | ----- | ------------------------------------------------------------------------------------ |
| **Integration** | Vitest     | `tests/integration/` | 36    | Component wiring; tool schemas; stores; hooks; live Copilot WebSocket |
| **UI**          | Playwright | `tests-ui/`          |       | Browser task pane flows                                                              |
| **E2E (Excel)** | Mocha      | `tests-e2e/`         | ~187  | Excel commands inside real Excel Desktop                                             |
| **E2E (PPT)**   | Mocha      | `tests-e2e-ppt/`     | ~13   | PowerPoint commands inside real PowerPoint Desktop                                   |
| **E2E (Word)**  | Mocha      | `tests-e2e-word/`    | ~12   | Word commands inside real Word Desktop                                               |
| ~~Unit~~        | ~~Vitest~~ | ~~`tests/unit/`~~    |       | ~~DO NOT ADD NEW UNIT TESTS~~                                                        |

### Required Test Execution After Any Code Change

**ALWAYS run integration and E2E tests after making code changes — these are the only tests that matter:**

1. `npm run test:integration` — integration tests (**ALL must pass — 0 failures is the only acceptable result**)
2. `npm run test:e2e` — E2E tests inside real Excel Desktop (requires `npm run start:desktop` first; **must pass before marking work complete**)
3. `npm run test:e2e:ppt` — E2E tests inside real PowerPoint Desktop (requires PPT open; **must pass before marking PPT work complete**)
4. `npm run test:e2e:word` — E2E tests inside real Word Desktop (requires Word open; **must pass before marking Word work complete**)
5. `npm run test:ui` — Playwright UI tests when task pane flows are changed

**Never consider work done until integration and E2E tests pass for the affected host(s).** If E2E tests cannot be run (Office app not open), explicitly flag this as a blocker to the user — do not silently skip them.

> ### ⛔ ZERO FAILURES POLICY
>
> **0 test failures is the only acceptable result.** Any failure — including live Copilot WebSocket tests — is a blocker that must be flagged to the user. Never dismiss failures as "expected" or "needs server". If live Copilot tests fail because the dev server isn't running, tell the user: "9 live Copilot tests are failing because `npm run dev` is not running. Start the server or these failures block completion."

> `tests/unit/` is **empty** — all logic has been migrated to `tests/integration/`. There are no unit tests in this codebase.

### Current Integration Test Files (36)

| File                                    | Category                            | Requires server? |
| --------------------------------------- | ----------------------------------- | ---------------- |
| `agent-manager-dialog.test.tsx`         | Component wiring                    | No               |
| `agent-picker.test.tsx`                 | Component wiring                    | No               |
| `agent-service.test.ts`                 | Agent service + frontmatter parsing | No               |
| `app-error-boundary.test.tsx`           | Component wiring                    | No               |
| `app-session-error.test.tsx`            | Component wiring                    | No               |
| `app-state.test.tsx`                    | Component wiring                    | No               |
| `chat-error-boundary.test.tsx`          | Component wiring                    | No               |
| `chat-header-settings-flow.test.tsx`    | Component wiring                    | No               |
| `chat-panel.test.tsx`                   | Component wiring                    | No               |
| `chat-store.test.ts`                    | Chat message store                  | No               |
| `copilot-custom-agent.integration.test.ts` | Live Copilot custom agent + skills | Yes (fails without server) |
| `copilot-websocket.integration.test.ts` | Live Copilot WebSocket E2E          | Yes (fails without server) |
| `excel-tools.test.ts`                   | Tool schema + factory (Excel)       | No               |
| `general-tools.test.ts`                 | General-purpose tool definitions    | No               |
| `host-tools-limit.test.ts`              | Host tool count limits              | No               |
| `humanize-tool-name.test.ts`            | Tool-name → human-readable labels   | No               |
| `id.test.ts`                            | `generateId` utility                | No               |
| `management-tools.test.ts`              | Management tool schemas + handlers  | No               |
| `manifest.test.ts`                      | Office manifest / host assumptions  | No               |
| `mcp-manager-dialog.test.tsx`           | Component wiring                    | No               |
| `mcp-service.test.ts`                   | MCP server config parsing           | No               |
| `model-manager.test.tsx`                | Component wiring                    | No               |
| `model-picker-interactions.test.tsx`    | Component wiring                    | No               |
| `office-storage.test.ts`                | `officeStorage` with OfficeRuntime  | No               |
| `powerpoint-tools.test.ts`              | Tool schema + factory (PPT)         | No               |
| `settings-dialog.test.tsx`              | Component wiring                    | No               |
| `settings-store.test.ts`                | Zustand store (model/agent/skills)  | No               |
| `skill-manager-dialog.test.tsx`         | Component wiring                    | No               |
| `skill-picker.test.tsx`                 | Component wiring                    | No               |
| `skill-service.test.ts`                 | Skill service + context building    | No               |
| `stale-state.test.tsx`                  | Store hydration                     | No               |
| `use-office-chat.test.tsx`              | useOfficeChat hook                  | No               |
| `use-tool-invocations-patch.test.tsx`   | Tool invocation argument streaming  | No               |
| `word-tools.test.ts`                    | Tool schema + factory (Word)        | No               |
| `zip-export-service.test.ts`            | ZIP export service                  | No               |
| `zip-import-service.test.ts`            | ZIP import service                  | No               |

### Integration Test Categories

- **Component wiring** — renders real components together (no child mocks)
- **Live Copilot WebSocket** — hits real GitHub Copilot API via proxy (requires `npm run dev`; **fails when server is unavailable — these failures are real and must be flagged**)

### When to Write What

- **New Excel command?** → E2E test in `tests-e2e/`
- **New PowerPoint command?** → E2E test in `tests-e2e-ppt/`
- **New Word command?** → E2E test in `tests-e2e-word/`
- **New task pane interaction flow?** → UI test in `tests-ui/`
- **New React component or hook behavior?** → Integration test in `tests/integration/`
- **New host routing rule?** → Integration test in `tests/integration/`
- **New tool definition?** → Integration test in `tests/integration/`
- **New pure function?** → Integration test — do NOT write a unit test

## Code Conventions

### Imports

- Use `@/` path alias (maps to `src/`)
- Barrel exports: import from `@/services/ai`, `@/tools`, `@/stores`, `@/types`

### State Management

- Single Zustand store: `useSettingsStore` in `src/stores/settingsStore.ts`
- Persisted via `officeStorage` adapter (uses `OfficeRuntime.storage`; throws when unavailable — tests must mock it via `tests/setup.ts`)
- Chat state is ephemeral (lives in `useOfficeChat` hook, not persisted)
- `activeAgentId`, `activeSkillNames` (default: `null` = all ON), and `activeModel` are persisted
- Persist storage key is `office-coding-agent-settings`

### Tool Definitions

- Excel tools are defined across 9 config modules (range, table, chart, sheet, workbook, comment, conditionalFormat, dataValidation, pivotTable)
- Each config module in `src/tools/configs/` defines tool schemas and handlers
- Tool factory in `src/tools/codegen/factory.ts` generates JSON Schema `Tool[]` for the Copilot SDK
- **General-purpose tools** (`src/tools/general.ts` — `web_fetch`) and **management tools** (`src/tools/management.ts` — `manage_skills`, `manage_agents`, `manage_mcp_servers`) are included for all hosts
- Host routing is in `src/tools/index.ts` via `getToolsForHost(host)` → `Tool[]` (host tools + general tools)

### UX Patterns

- **Dynamic thinking indicator** — `ThinkingIndicator` in `thread.tsx` displays dynamic text during tool execution. Text sources: (1) `report_intent` SDK events set the raw intent string (e.g. "Reading the spreadsheet"); (2) every `tool.execution_start` sets the humanized tool name via `humanizeToolName()` (e.g. "Get range values…"). When no text is set, falls back to "Thinking…". Text is provided via `ThinkingContext` (`src/contexts/ThinkingContext.ts`), populated by `useOfficeChat`, and cleared on stream completion.
- **Copilot-style progress indicators** — cycling dot animation + phase labels (auto-derived via `humanizeToolName()`)
- **Choice cards** — `PromptStarterV2` renders ` ```choices ` blocks as clickable cards
- **Tool result summaries** — collapsible progress sections with `toolResultSummary()` one-liners
- **Input toolbar** — AgentPicker + ModelPicker below Composer (GitHub Copilot-style)

### OfficeRuntime in Tests

- `officeStorage.ts` throws if `OfficeRuntime.storage` is unavailable (no localStorage fallback)
- Unit and integration tests rely on the `OfficeRuntime` mock in `tests/setup.ts`
- **Both** `vitest.config.ts` and `vitest.integration.config.ts` must include `setupFiles: ['tests/setup.ts']` and `globals: true`

## Build & Run

```bash
npm install
npm run dev               # Start Copilot proxy + Vite dev server (port 3000)
npm run build:dev         # Development build
npm run build             # Production build
npm run start:desktop     # Sideload into Excel
npm test                  # All Vitest unit tests
npm run test:integration  # Integration tests (429)
npm run test:ui           # Playwright UI tests
npm run test:e2e          # E2E in Excel Desktop (~187 tests)
npm run validate          # Validate manifests/manifest.dev.xml
```

## Key Files

- `src/taskpane/App.tsx` — root component, settings dialog state, theme detection, Office host detection, `ThinkingContext` provider
- `src/hooks/useOfficeChat.ts` — main hook: WebSocket session lifecycle → `useExternalStoreRuntime`, `report_intent` → `thinkingText`
- `src/contexts/ThinkingContext.ts` — React context for dynamic thinking indicator text (populated from `report_intent`)
- `src/lib/websocket-client.ts` — `WebSocketCopilotClient`, `BrowserCopilotSession`, `createWebSocketClient`
- `src/lib/websocket-transport.ts` — JSON-RPC WebSocket transport (browser-compatible)
- `src/server.mjs` — Express HTTPS server (port 3000): Vite dev middleware + Copilot WebSocket proxy
- `src/copilotProxy.mjs` — bridges WebSocket to `@github/copilot-sdk` CopilotClient
- `src/components/ChatHeader.tsx` — header: title, SkillPicker, new convo, Settings
- `src/components/ChatPanel.tsx` — messages, progress, input, AgentPicker, ModelPicker
- `src/components/ChatErrorBoundary.tsx` — error boundary around chat UI
- `src/components/AgentPicker.tsx` — single-select agent dropdown (Radix Popover)
- `src/components/ModelPicker.tsx` — model selection dropdown (dynamic, fetched from Copilot API)
- `src/components/SkillPicker.tsx` — icon-only skill toggle with badge count
- `src/components/SettingsDialog.tsx` — settings/preferences dialog
- `src/services/ai/BASE_PROMPT.md` — universal base system prompt
- `src/services/ai/prompts/` — host-level app prompts (`EXCEL_APP_PROMPT.md`, `POWERPOINT_APP_PROMPT.md`, `WORD_APP_PROMPT.md`)
- `src/services/office/host.ts` — Office host detection (`excel`, `powerpoint`, `word`, `unknown`)
- `src/services/agents/agentService.ts` — parses/filters host-targeted agent frontmatter
- `src/agents/excel/AGENT.md` — default Excel agent definition (host-targeted frontmatter)
- `src/agents/powerpoint/AGENT.md` — default PowerPoint agent definition
- `src/agents/word/AGENT.md` — default Word agent definition
- `src/services/skills/skillService.ts` — parses bundled skill files + host filtering via `buildSkillContext(activeNames?, host?)`
- `src/stores/settingsStore.ts` — Zustand store (activeModel, agent/skill CRUD, reset)
- `src/stores/officeStorage.ts` — OfficeRuntime.storage adapter (throws when unavailable)
- `src/tools/` — 9 tool config modules + codegen factory (`Tool[]` for Copilot SDK)
- `src/tools/general.ts` — `webFetchTool` (general-purpose, all hosts)
- `src/tools/management.ts` — `manage_skills`, `manage_agents`, `manage_mcp_servers` tools
- `src/types/settings.ts` — `CopilotModel`, `inferProvider()`, `UserSettings`
- `src/utils/toolResultSummary.ts` — human-readable one-liner summaries for tool results
- `vite.config.ts` — Vite build config (React plugin, md-raw plugin, static copy, `@/` alias)
- `taskpane.html` — Vite HTML entry point (root level, references `src/taskpane/index.tsx`)
- `vitest.config.ts` — unit test config (jsdom, `@/` alias, setup file, globals)
- `vitest.integration.config.ts` — integration test config (jsdom, setup file, globals, 60s timeout)
- `tests/setup.ts` — `OfficeRuntime.storage` mock + polyfills (ResizeObserver, matchMedia, etc.)
