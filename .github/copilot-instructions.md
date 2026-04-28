# Copilot Instructions for office-coding-agent

## ⛔ NEVER PUSH OR MERGE DIRECTLY TO MAIN

**Direct commits, merges, or pushes to `main` are strictly forbidden — no exceptions.**

All changes must go through a pull request on a feature branch:

1. `git checkout -b <branch-name>` — create a feature branch
2. Make and commit changes on that branch
3. `git push -u origin <branch-name>` — push the branch
4. Ask the user to review and merge the PR on GitHub

**Never run:**

- `git push origin main`
- `git merge <branch> ` while on `main`
- `git commit --no-verify` to bypass hooks and then push to `main`

Branch protection is enforced on GitHub (ruleset ID `13260767`). Any attempt to push directly to `main` will be rejected by the server.

**PRs must be merged via squash merge only.** Merge commits and rebase merges are disabled in the repository settings. When asking the user to merge a PR, always tell them to use **"Squash and merge"** on GitHub.

## ⛔ UI GOLDEN RULE: LOOK AND FEEL IDENTICAL TO VS CODE

**The add-in must look and feel identical to VS Code's Copilot Chat — no exceptions.**

Every UI element, interaction, and visual detail must match what a user sees in VS Code's Copilot Chat panel. When in doubt, open VS Code's Copilot Chat and copy what you see.

### Design System
- **Colors:** Use `--vscode-*` CSS custom properties from `src/styles/vscode-theme.css` — NEVER hardcode colors or use non-VS-Code values
- **Icons:** Use codicons (`@vscode/codicons`) for ALL icons — NEVER use lucide-react or other icon libraries for new icons
- **Typography:** VS Code font stack (13px Segoe UI) — NEVER change the base font size or family
- **Spacing & Radii:** Use VS Code corner radii (2px small, 4px medium, 8px large) and compact spacing — NEVER use large padding or oversized elements
- **Focus:** 1px solid `--vscode-focusBorder` outline — NEVER use box-shadow rings

### Chat Components
- **Messages:** Full-width, no chat bubbles. Assistant messages have a 22px Copilot avatar + "Copilot" label header
- **User messages:** Plain text, no border or background box
- **Thinking indicator:** Shimmer gradient animation (`chat-thinking-shimmer`) on text — NEVER use spinners or pulse animations
- **Tool invocations:** Collapsible sections with codicon chevron + shimmer on running tool name
- **Follow-up suggestions:** Blue link-colored text (`textLink-foreground`) with sparkle codicon — NEVER use bordered cards
- **Welcome screen:** Centered Copilot avatar, compact heading, blue link-colored suggestion items
- **Composer input:** VS Code input background/border, 8px border-radius, no toolbar divider line

### When Adding New UI
1. First, check how VS Code's Copilot Chat does it (open VS Code or check `microsoft/vscode` source)
2. Copy the HTML structure and CSS patterns from VS Code
3. Use the same class naming conventions and design tokens
4. If VS Code doesn't have a precedent for the feature, use the closest VS Code UI pattern (panel, quick pick, dialog, etc.)

## Project Overview

**office-coding-agent** is a Microsoft Office add-in with a single task pane UI and host-routed AI runtime behavior. The current implementation is **Excel-first**, but tools and prompts are selected by host (`excel`, `powerpoint`, etc.) to support future hosts without changing the UI.

## Key Technologies

- **React 18** + **Radix UI** + **Tailwind CSS v4** — task pane UI (MessageList, ChatComposer, AssistantMessage, UserMessage, MarkdownContent, ToolProgress, ActionBar); **VS Code design system** via `vscode-theme.css` (`--vscode-*` CSS custom properties) and **codicons** (`@vscode/codicons`) for all icons
- **GitHub Copilot SDK** (`@github/copilot-sdk`) — session management, streaming events, tool registration
- **WebSocket + JSON-RPC** — browser-to-proxy transport (`src/lib/websocket-client.ts`, `src/lib/websocket-transport.ts`)
- **Express + HTTPS** — local proxy server (`src/server.mjs`) that bridges WebSocket to the Copilot CLI
- **Zustand 5** — state management with persistence via `officeStorage` (OfficeRuntime.storage)
- **Vite 7** — bundling, dev server (HMR via middleware mode in Express)
- **TypeScript 5** — type safety
- **Vitest** — unit + integration testing (jsdom env)
- **Mocha** — E2E tests inside real Office Desktop hosts (Excel, PowerPoint, Word, Outlook)
- **Playwright** — browser UI tests for task pane flows

## Architecture

The add-in routes messages through a **local proxy server** — the browser cannot call the Copilot API directly.

```
Browser task pane (React)
         ↓ WebSocket (wss://localhost:3000/api/copilot)
Node.js proxy server  (src/server.mjs + src/copilotProxy.mjs)
         ↓ @github/copilot-sdk (manages CLI lifecycle internally)
GitHub Copilot API
```

The `useOfficeChat` hook creates a `WebSocketCopilotClient`, opens a `BrowserCopilotSession`, and maps incoming `SessionEvent` objects to `ChatMessage[]` state with per-message `thinkingText` fields.

### Prompt and CLI Agent System

The add-in uses a split prompt architecture with host targeting:

- **`src/services/ai/BASE_PROMPT.md`** — universal base prompt (progress narration + presenting choices)
- **`src/services/ai/prompts/*_APP_PROMPT.md`** — host-level app prompt (Excel/PowerPoint/Word/Outlook)
- Instructions = `buildSystemPrompt(host) + memory context`

Agents are owned by the Copilot CLI. The add-in does not ship `src/agents/*/AGENT.md`, parse agent frontmatter, or pass browser-owned `customAgents` into `session.create`. It does render an Agent picker, but that picker lists/selects CLI-owned agents through the SDK session agent RPC and persists only the selected CLI agent name. The required Office agents are installed through the Office Coding Agent CLI plugins (`office-excel`, `office-powerpoint`, `office-word`, `office-outlook`) during startup bootstrap.

### Skills System

Plugin-provided skills are managed by the Copilot CLI. The task pane does not render a skills picker or `/skills` management surface; users can reference installed CLI plugin skills and `.prompt.md` prompt files as `/skill-name` or `/prompt-name`, matching Copilot slash invocation. Installation, updates, and removal stay in the terminal with `copilot plugin`.

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
│  • Zustand store logic (settingsStore)               │
│  • JSON Schema tool configs (toCopilotTools)         │
│  • React component wiring (integration)              │
│  • WebSocket client + session wiring                 │
│  • Skill service parsing                             │
├──────────────────────────────────────────────────────┤
│  Office.run() boundary (all hosts)                   │
├──────────────────────────────────────────────────────┤
│  E2E only (Mocha + real Office Desktop)              │
│  ─────────────────────────────                       │
│  • rangeCommands, tableCommands, sheetCommands       │
│  • chartCommands, workbookCommands, commentCommands  │
│  • conditionalFormatCommands, dataValidationCommands │
│  • pivotTableCommands                                │
│  • PowerPoint / Word / Outlook commands              │
│  • OfficeRuntime.storage (real runtime)              │
└──────────────────────────────────────────────────────┘
```

## UI Layout

The task pane is split into three areas:

- **ChatHeader** — session history, Copilot CLI plugin help link, permissions, compact conversation, and new conversation controls
- **ChatPanel** — message list (MessageList), Copilot-style progress indicators, choice cards, error bar, ChatComposer, and an input toolbar below the input box with ModelPicker + MCP picker
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

| Tier              | Runner     | Directory            | Count | What it tests                                                         |
| ----------------- | ---------- | -------------------- | ----- | --------------------------------------------------------------------- |
| **Integration**   | Vitest     | `tests/integration/` | 36    | Component wiring; tool schemas; stores; hooks; live Copilot WebSocket |
| **UI**            | Playwright | `tests-ui/`          |       | Browser task pane flows (real Copilot API, NO mocking)                |
| **E2E (Excel)**   | Mocha      | `tests-e2e/`         | ~187  | Excel commands inside real Excel Desktop                              |
| **E2E (PPT)**     | Mocha      | `tests-e2e-ppt/`     | ~13   | PowerPoint commands inside real PowerPoint Desktop                    |
| **E2E (Word)**    | Mocha      | `tests-e2e-word/`    | ~12   | Word commands inside real Word Desktop                                |
| **E2E (Outlook)** | Mocha      | `tests-e2e-outlook/` | ~9    | Outlook commands (requires Exchange sideloading approval)             |
| ~~Unit~~          | ~~Vitest~~ | ~~`tests/unit/`~~    |       | ~~DO NOT ADD NEW UNIT TESTS~~                                         |

### Required Test Execution After Any Code Change

**ALWAYS run integration and E2E tests after making code changes — these are the only tests that matter:**

2. `npm run test:integration` — integration tests via `--project integration` (**ALL must pass — 0 failures is the only acceptable result**)
3. `npm run test:e2e` — E2E tests inside real Excel Desktop (requires `npm run start:desktop` first; **must pass before marking work complete**)
4. `npm run test:e2e:ppt` — E2E tests inside real PowerPoint Desktop (requires PPT open; **must pass before marking PPT work complete**)
5. `npm run test:e2e:word` — E2E tests inside real Word Desktop (requires Word open; **must pass before marking Word work complete**)
6. `npm run test:e2e:outlook` — E2E tests inside real Outlook Desktop (requires Exchange sideloading approval; blocked on tenants with policy restrictions — flag as blocker if unavailable)
7. `npm run test:ui` — Playwright UI tests when task pane flows are changed

**Never consider work done until integration and E2E tests pass for the affected host(s).** If live Copilot WebSocket tests fail because the dev server is not running, start `npm run dev` as a background process and re-run — do not skip or report as blocked. If E2E tests cannot be run (Office app not open), explicitly flag this as a blocker to the user — do not silently skip them.

> ### ⛔ ZERO FAILURES POLICY
>
> **0 test failures is the only acceptable result.** Any failure is a blocker — there are no acceptable failures, no "expected" failures, no "needs server" exceptions.
>
> **If tests fail because the dev server is not running, START IT — do not just report it.**
> Run `npm run dev` as a background process, wait for `https://localhost:3000` to be ready, then re-run the failing tests. Only after the server is up and the tests still fail should you escalate to the user.
>
> Dev server start sequence:
>
> 1. Start `npm run dev` in the background (it binds to `https://localhost:3000`)
> 2. Poll or wait ~10 s for `Copilot Office Add-in server running on https://localhost:3000`
> 3. Re-run the affected tests
> 4. If tests pass → work is complete. If tests still fail → report the actual failure to the user.

> ### ⛔ NO MOCKING IN PLAYWRIGHT TESTS
>
> Playwright UI tests (`tests-ui/`) exist to verify end-to-end flows through the **real** Copilot API and proxy server. **Never mock, stub, intercept, or fake** any network request, WebSocket connection, API response, or server behavior in Playwright tests. No `page.route()`, no mock WebSocket servers, no fake session events. If the dev server must be running for a test to work, that is a prerequisite — not a reason to mock. Mocking in Playwright defeats the entire purpose of these tests.

> `tests/unit/` is **empty** — all logic has been migrated to `tests/integration/`. There are no unit tests in this codebase.

### Current Integration Test Files (40)

| File                                       | Category                            | Requires server?           |
| ------------------------------------------ | ----------------------------------- | -------------------------- |
| `agent-manager-dialog.test.tsx`            | Component wiring                    | No                         |
| `agent-picker.test.tsx`                    | Component wiring                    | No                         |
| `agent-service.test.ts`                    | Agent service + frontmatter parsing | No                         |
| `app-error-boundary.test.tsx`              | Component wiring                    | No                         |
| `app-session-error.test.tsx`               | Component wiring                    | No                         |
| `app-state.test.tsx`                       | Component wiring                    | No                         |
| `chat-error-boundary.test.tsx`             | Component wiring                    | No                         |
| `chat-header-settings-flow.test.tsx`       | Component wiring                    | No                         |
| `chat-panel.test.tsx`                      | Component wiring                    | No                         |
| `chat-store.test.ts`                       | Chat message store                  | No                         |
| `copilot-custom-agent.integration.test.ts` | Live Copilot custom agent + skills  | Yes (fails without server) |
| `copilot-websocket.integration.test.ts`    | Live Copilot WebSocket E2E          | Yes (fails without server) |
| `excel-tools.test.ts`                      | Tool schema + factory (Excel)       | No                         |
| `host-tools-limit.test.ts`                 | Host tool count limits              | No                         |
| `humanize-tool-name.test.ts`               | Tool-name → human-readable labels   | No                         |
| `id.test.ts`                               | `generateId` utility                | No                         |
| `management-tools.test.ts`                 | Management tool schemas + handlers  | No                         |
| `manifest.test.ts`                         | Office manifest / host assumptions  | No                         |
| `mcp-manager-dialog.test.tsx`              | Component wiring                    | No                         |
| `mcp-service.test.ts`                      | MCP server config parsing           | No                         |
| `model-manager.test.tsx`                   | Component wiring                    | No                         |
| `model-picker-interactions.test.tsx`       | Component wiring                    | No                         |
| `office-storage.test.ts`                   | `officeStorage` with OfficeRuntime  | No                         |
| `powerpoint-tools.test.ts`                 | Tool schema + factory (PPT)         | No                         |
| `settings-dialog.test.tsx`                 | Component wiring                    | No                         |
| `settings-store.test.ts`                   | Zustand store (model/agent/skills)  | No                         |
| `skill-manager-dialog.test.tsx`            | Component wiring                    | No                         |
| `skill-service.test.ts`                    | Skill service + context building    | No                         |
| `stale-state.test.tsx`                     | Store hydration                     | No                         |
| `thread-message-rendering.test.tsx`        | Component wiring                    | No                         |
| `use-office-chat.test.tsx`                 | useOfficeChat hook                  | No                         |
| `use-tool-invocations-patch.test.tsx`      | Tool invocation argument streaming  | No                         |
| `word-tools.test.ts`                       | Tool schema + factory (Word)        | No                         |
| `zip-export-service.test.ts`               | ZIP export service                  | No                         |
| `zip-import-service.test.ts`               | ZIP import service                  | No                         |

### Integration Test Categories

- **Component wiring** — renders real components together (no child mocks)
- **Live Copilot WebSocket** — hits real GitHub Copilot API via proxy (requires `npm run dev`; **fails when server is unavailable — these failures are real and must be flagged**)

### When to Write What

- **New Excel command?** → E2E test in `tests-e2e/`
- **New PowerPoint command?** → E2E test in `tests-e2e-ppt/`
- **New Word command?** → E2E test in `tests-e2e-word/`
- **New Outlook command?** → E2E test in `tests-e2e-outlook/`
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
- `activeModel`, `activeAgentName`, and `disabledMcpServerNames` are persisted
- Persist storage key is `office-coding-agent-settings`

### Tool Definitions

- Excel tools are defined across 9 config modules (range, table, chart, sheet, workbook, comment, conditionalFormat, dataValidation, pivotTable)
- Each config module in `src/tools/configs/` defines tool schemas and handlers
- Tool factory in `src/tools/codegen/factory.ts` generates JSON Schema `Tool[]` for the Copilot SDK
- **Management tools** (`src/tools/management.ts` — `manage_skills`, `manage_agents`, `manage_mcp_servers`) are included for all hosts
- Host routing is in `src/tools/index.ts` via `getToolsForHost(host)` → `Tool[]` (host tools + general tools)

### UX Patterns

- **Dynamic thinking indicator** — `ThinkingIndicator` component displays dynamic text during tool execution. Text sources: (1) `report_intent` SDK events set the raw intent string (e.g. "Reading the spreadsheet"); (2) every `tool.execution_start` sets the humanized tool name via `humanizeToolName()` (e.g. "Get range values…"). When no text is set, falls back to "Thinking…". Text is stored per-message via the `thinkingText` field on `ChatMessage`, and cleared on stream completion.
- **Copilot-style progress indicators** — VS Code shimmer effect (`chat-thinking-shimmer` CSS animation using `background-clip: text` gradient animation matching VS Code Copilot Chat) + phase labels (auto-derived via `humanizeToolName()`)
- **Choice cards** — `PromptStarterV2` renders ` ```choices ` blocks as clickable cards
- **Tool result summaries** — collapsible progress sections with `toolResultSummary()` one-liners
- **Input toolbar** — ModelPicker below ChatComposer plus MCP picker on the right

### Styling

- Colors use VS Code design tokens via `--vscode-*` CSS custom properties defined in `vscode-theme.css`
- Layout uses Tailwind utility classes; colors reference VS Code tokens
- Focus styles: 1px outline with `--vscode-focusBorder`, no box-shadow rings
- Font: VS Code font stack (Segoe UI, system fonts), 13px base size

### OfficeRuntime in Tests

- `officeStorage.ts` throws if `OfficeRuntime.storage` is unavailable (no localStorage fallback)
- Unit and integration tests rely on the `OfficeRuntime` mock in `tests/setup.ts`
- Both the `unit` and `integration` projects in `vitest.config.ts` must include `setupFiles: ['tests/setup.ts']` and `globals: true`

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

- `src/taskpane/App.tsx` — root component, settings dialog state, theme detection, Office host detection
- `src/hooks/useOfficeChat.ts` — main hook: WebSocket session lifecycle → `ChatMessage[]` state, `report_intent` → per-message `thinkingText`
- `src/lib/websocket-client.ts` — `WebSocketCopilotClient`, `BrowserCopilotSession`, `createWebSocketClient`
- `src/lib/websocket-transport.ts` — JSON-RPC WebSocket transport (browser-compatible)
- `src/server.mjs` — Express HTTPS server (port 3000): Vite dev middleware + Copilot WebSocket proxy
- `src/copilotProxy.mjs` — bridges WebSocket to `@github/copilot-sdk` CopilotClient
- `src/components/ChatHeader.tsx` — header: session history, plugin help, permissions, compact conversation, new conversation
- `src/components/ChatPanel.tsx` — messages (MessageList), progress, input (ChatComposer), ModelPicker, MCP picker
- `src/components/ChatErrorBoundary.tsx` — error boundary around chat UI
- `src/components/chat/MessageList.tsx` — renders `ChatMessage[]` as AssistantMessage, UserMessage, or ToolProgress components
- `src/components/chat/ChatComposer.tsx` — text input for user messages
- `src/components/chat/AssistantMessage.tsx` — assistant message with MarkdownContent + ToolGroup
- `src/components/chat/UserMessage.tsx` — user message display
- `src/components/chat/MarkdownContent.tsx` — renders Markdown with syntax highlighting
- `src/components/chat/ToolProgress.tsx` — tool execution progress and results
- `src/components/chat/ActionBar.tsx` — action buttons (copy, retry, etc.)
- `src/components/ModelPicker.tsx` — model selection dropdown (dynamic, fetched from Copilot API)
- `src/components/SettingsDialog.tsx` — settings/preferences dialog
- `src/services/ai/BASE_PROMPT.md` — universal base system prompt
- `src/services/ai/prompts/` — host-level app prompts (`EXCEL_APP_PROMPT.md`, `POWERPOINT_APP_PROMPT.md`, `WORD_APP_PROMPT.md`)
- `src/services/office/host.ts` — Office host detection (`excel`, `powerpoint`, `word`, `unknown`)
- `src/services/skills/skillService.ts` — parses bundled skill files + host filtering via `buildSkillContext(activeNames?, host?)`
- `src/stores/settingsStore.ts` — Zustand store (activeModel, skill/MCP enablement, reset)
- `src/stores/officeStorage.ts` — OfficeRuntime.storage adapter (throws when unavailable)
- `src/tools/` — 9 tool config modules + codegen factory (`Tool[]` for Copilot SDK)
- `src/tools/management.ts` — `manage_skills`, `manage_agents`, `manage_mcp_servers` tools
- `src/types/settings.ts` — `CopilotModel`, `inferProvider()`, `UserSettings`, `ChatMessage` (stores `thinkingText` per message)
- `src/utils/toolResultSummary.ts` — human-readable one-liner summaries for tool results
- `vite.config.ts` — Vite build config (React plugin, md-raw plugin, static copy, `@/` alias)
- `src/styles/vscode-theme.css` — VS Code design tokens (`--vscode-*` CSS custom properties for dark/light themes)
- `taskpane.html` — Vite HTML entry point (root level, references `src/taskpane/index.tsx`)
- `vitest.config.ts` — unified vitest config with two named projects: `unit` (30s) and `integration` (60s, live Copilot tests)
- `tests/setup.ts` — `OfficeRuntime.storage` mock + polyfills (ResizeObserver, matchMedia, etc.)
