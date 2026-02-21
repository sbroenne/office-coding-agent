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
- `hosts` (array; supported values: `excel`, `powerpoint`)
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

Bundled skill files in `src/skills/` provide additional context injected into the system prompt. Skills can be toggled on/off via the SkillPicker. Active skills are stored as `activeSkillNames` in the settings store (`null` = all ON).

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
│  • OfficeRuntime.storage (real runtime)              │
└──────────────────────────────────────────────────────┘
```

## UI Layout

The task pane is split into three areas:

- **ChatHeader** — "AI Chat" title + SkillPicker (icon-only with badge) + New Conversation button + Settings gear (SettingsDialog)
- **ChatPanel** — message list (Thread), Copilot-style progress indicators, choice cards, error bar, Composer, and an **input toolbar** below the input box with AgentPicker + ModelPicker (GitHub Copilot-style)
- **App** — owns settings dialog state, detects system theme and Office host; no setup wizard (Copilot CLI handles auth)

## Testing Strategy

### Test Tiers

| Tier            | Runner     | Directory            | Count | What it tests                                           |
| --------------- | ---------- | -------------------- | ----- | ------------------------------------------------------- |
| **Unit**        | Vitest     | `tests/unit/`        | 18    | Pure functions, store logic, JSON Schema tool configs   |
| **Integration** | Vitest     | `tests/integration/` | 12    | Component wiring; live Copilot WebSocket (auto-skipped) |
| **UI**          | Playwright | `tests-ui/`          |       | Browser task pane flows                                 |
| **E2E**         | Mocha      | `tests-e2e/`         | ~187  | Excel commands inside real Excel Desktop                |

### Unit Test Principles

- **DO NOT mock Zustand stores** — test against the real store (jsdom + OfficeRuntime mock from `tests/setup.ts`)
- **DO NOT mock pure functions** — call them directly with test inputs
- **Pure functions get unit tests, not integration tests** — they have no external dependencies
- **Use table-driven tests** (`it.each`) for functions with many input→output mappings
- **Reset store state** in `beforeEach` via `useSettingsStore.getState().reset()`

### Current Unit Test Files (18)

| File                                | What it covers                                                           |
| ----------------------------------- | ------------------------------------------------------------------------ |
| `agentService.test.ts`              | Agent frontmatter parsing, getAgents, getAgent, getAgentInstructions     |
| `buildSkillContext.test.ts`         | `buildSkillContext` and related skill functions with bundled `.md` files |
| `chatErrorBoundary.test.tsx`        | Error boundary fallback rendering and recovery flow                      |
| `chatPanel.test.tsx`                | ChatPanel component logic (mocks assistant-ui components for jsdom)      |
| `chatStore.test.ts`                 | Chat message store: append, clear, tool invocations                      |
| `generalTools.test.ts`              | General-purpose tool definitions (web_fetch, etc.)                       |
| `hostToolsLimit.test.ts`            | Host tool count limits per host                                          |
| `humanizeToolName.test.ts`          | Tool-name → human-readable progress label formatting                     |
| `id.test.ts`                        | `generateId` unique ID generation utility                                |
| `manifest.test.ts`                  | Office manifest / runtime host assumptions                               |
| `mcpService.test.ts`                | MCP server config parsing; HTTP/SSE transport filtering (no stdio)       |
| `officeStorage.test.ts`             | `officeStorage` with `OfficeRuntime` mock (via tests/setup.ts)           |
| `parseFrontmatter.test.ts`          | YAML frontmatter parsing (skill files)                                   |
| `settingsStore.test.ts`             | Zustand store: activeModel, agent/skill CRUD, reset                      |
| `toolSchemas.test.ts`               | JSON Schema validation for all tool definitions (toCopilotTools)         |
| `useOfficeChat.test.tsx`            | useOfficeChat hook: mocked WebSocket session → ThreadMessage[] mapping   |
| `useToolInvocations-patch.test.tsx` | assistant-ui patch for tool invocation argument streaming integrity      |
| `zipImportService.test.ts`          | ZIP import service for custom agents/skills                              |

### Current Integration Test Files (12)

| File                                    | Category                   | Requires server? |
| --------------------------------------- | -------------------------- | ---------------- |
| `agent-picker.test.tsx`                 | Component wiring           | No               |
| `app-error-boundary.test.tsx`           | Component wiring           | No               |
| `app-state.test.tsx`                    | Component wiring           | No               |
| `chat-header-settings-flow.test.tsx`    | Component wiring           | No               |
| `chat-panel.test.tsx`                   | Component wiring           | No               |
| `copilot-websocket.integration.test.ts` | Live Copilot WebSocket E2E | Yes (auto-skip)  |
| `mcp-manager-dialog.test.tsx`           | Component wiring           | No               |
| `model-manager.test.tsx`                | Component wiring           | No               |
| `model-picker-interactions.test.tsx`    | Component wiring           | No               |
| `settings-dialog.test.tsx`              | Component wiring           | No               |
| `skill-picker.test.tsx`                 | Component wiring           | No               |
| `stale-state.test.tsx`                  | Store hydration            | No               |

### Integration Test Categories

- **Component wiring** — renders real components together (no child mocks)
- **Live Copilot WebSocket** — hits real GitHub Copilot API via proxy (requires `npm run dev`; auto-skips when unavailable)

### When to Write What

- **New pure function?** → Unit test in `tests/unit/`
- **New Excel command?** → E2E test in `tests-e2e/`
- **New host routing rule?** → Unit/integration tests (`tests/unit/`, `tests/integration/`)
- **New task pane interaction flow?** → UI test in `tests-ui/`
- **New React component interaction?** → Integration test in `tests/integration/`
- **New tool definition?** → Add schema case to `tests/unit/toolSchemas.test.ts`

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
- Host routing is in `src/tools/index.ts` via `getToolsForHost(host)` → `Tool[]`

### UX Patterns

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
npm test                  # All Vitest unit tests (267)
npm run test:integration  # Integration tests (48)
npm run test:ui           # Playwright UI tests
npm run test:e2e          # E2E in Excel Desktop (~187 tests)
npm run validate          # Validate manifests/manifest.dev.xml
```

## Key Files

- `src/taskpane/App.tsx` — root component, settings dialog state, theme detection, Office host detection
- `src/hooks/useOfficeChat.ts` — main hook: WebSocket session lifecycle → `useExternalStoreRuntime`
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
- `src/services/ai/prompts/` — host-level app prompts (`EXCEL_APP_PROMPT.md`, `POWERPOINT_APP_PROMPT.md`)
- `src/services/office/host.ts` — Office host detection (`excel`, `powerpoint`, `unknown`)
- `src/services/agents/agentService.ts` — parses/filters host-targeted agent frontmatter
- `src/agents/excel/AGENT.md` — default Excel agent definition (host-targeted frontmatter)
- `src/services/skills/skillService.ts` — parses bundled skill files
- `src/stores/settingsStore.ts` — Zustand store (activeModel, agent/skill CRUD, reset)
- `src/stores/officeStorage.ts` — OfficeRuntime.storage adapter (throws when unavailable)
- `src/tools/` — 9 tool config modules + codegen factory (`Tool[]` for Copilot SDK)
- `src/types/settings.ts` — `CopilotModel`, `inferProvider()`, `UserSettings`
- `src/utils/toolResultSummary.ts` — human-readable one-liner summaries for tool results
- `vite.config.ts` — Vite build config (React plugin, md-raw plugin, static copy, `@/` alias)
- `taskpane.html` — Vite HTML entry point (root level, references `src/taskpane/index.tsx`)
- `vitest.config.ts` — unit test config (jsdom, `@/` alias, setup file, globals)
- `vitest.integration.config.ts` — integration test config (jsdom, setup file, globals, 60s timeout)
- `tests/setup.ts` — `OfficeRuntime.storage` mock + polyfills (ResizeObserver, matchMedia, etc.)
