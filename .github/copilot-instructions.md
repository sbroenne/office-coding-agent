# Copilot Instructions for office-coding-agent

## Project Overview

**office-coding-agent** is a Microsoft Office add-in with a single task pane UI and host-routed AI runtime behavior. The current implementation is **Excel-first**, but tools and prompts are selected by host (`excel`, `powerpoint`, etc.) to support future hosts without changing the UI.

## Key Technologies

- **React 18** + **assistant-ui** + **Radix UI** + **Tailwind CSS v4** — task pane UI (Thread, ToolFallback, Popovers)
- **Vercel AI SDK** — `ai` (ToolLoopAgent, DirectChatTransport), `@ai-sdk/react` (useChat), `@ai-sdk/azure`
- **Zustand 5** — state management with persistence via `officeStorage` (OfficeRuntime.storage with localStorage fallback)
- **Webpack 5** — bundling (ts-loader, full type-checking during builds)
- **TypeScript 5** — type safety
- **Vitest** — unit + integration testing (jsdom env)
- **Mocha** — E2E tests inside Excel Desktop (current host runtime E2E)
- **Playwright** — browser UI tests for task pane flows

## Architecture

The add-in runs a **fully client-side AI agent** — no backend API route. The `useOfficeChat` hook creates a `ToolLoopAgent` with `DirectChatTransport`, passing it to Vercel AI SDK's `useChat`.

```
useOfficeChat + detectOfficeHost
             ↓
buildSystemPrompt(host) + getToolsForHost(host)
             ↓
ToolLoopAgent → DirectChatTransport → useChat
             ↓
Azure AI Foundry (streaming LLM)
```

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

Current state:

- Excel commands call `Excel.run()` and require Excel E2E tests.
- Host routing, prompt composition, agent parsing/filtering, and UI logic are testable in Vitest/Playwright.

```
┌──────────────────────────────────────────────────────┐
│  Testable with Vitest/Playwright (no Excel host)     │
│  ─────────────────────────────                       │
│  • Pure functions (normalizeEndpoint,                │
│    inferProvider, formatModelName,                    │
│    isEmbeddingOrUtilityModel, parseFrontmatter,      │
│    buildSkillContext, toolResultSummary, generateId)  │
│  • Host routing (detectOfficeHost,                   │
│    getToolsForHost, buildSystemPrompt)               │
│  • Agent targeting and default resolution            │
│  • Zustand store logic (settingsStore)               │
│  • Zod tool schemas (inputSchema validation)         │
│  • React component wiring (integration)              │
│  • AI client factory (with live API creds)           │
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
- **App** — owns settings dialog state, routes between SetupWizard and chat UI, detects system theme and Office host

Settings dialog state is **lifted to App** so both ChatHeader (gear button) and ChatPanel (ModelPicker's "Configure models in Settings") can open it.

## Testing Strategy

### Test Tiers

| Tier            | Runner     | Directory            | Count | What it tests                                               |
| --------------- | ---------- | -------------------- | ----- | ----------------------------------------------------------- |
| **Unit**        | Vitest     | `tests/unit/`        |       | Pure functions, store logic, Zod schemas, host routing      |
| **Integration** | Vitest     | `tests/integration/` |       | Component wiring, live API round-trips, host-agent behavior |
| **UI**          | Playwright | `tests-ui/`          |       | Browser task pane flows (wizard/chat/settings)              |
| **E2E**         | Mocha      | `tests-e2e/`         |       | Excel commands inside real Excel Desktop                    |

### Unit Test Principles

- **DO NOT mock Zustand stores** — test against the real store with localStorage (works in jsdom)
- **DO NOT mock pure functions** — call them directly with test inputs
- **Pure functions get unit tests, not integration tests** — they have no external dependencies
- **Use table-driven tests** (`it.each`) for functions with many input→output mappings
- **Reset store state** in `beforeEach` via `useSettingsStore.getState().reset()`

### Current Unit Test Files (12)

| File                                        | What it covers                                                                 |
| ------------------------------------------- | ------------------------------------------------------------------------------ |
| `tests/unit/agentService.test.ts`           | Agent frontmatter parsing, getAgents, getAgent, getAgentInstructions           |
| `tests/unit/buildSkillContext.test.ts`      | `buildSkillContext` and related skill functions with bundled `.md` files       |
| `tests/unit/chatPanel.test.tsx`             | ChatPanel component logic (mocks assistant-ui components for jsdom)            |
| `tests/unit/id.test.ts`                     | `generateId` unique ID generation utility                                      |
| `tests/unit/messagesToCoreMessages.test.ts` | `messagesToCoreMessages` conversion from ChatMessage[] to core messages        |
| `tests/unit/modelDiscoveryHelpers.test.ts`  | `inferProvider`, `isEmbeddingOrUtilityModel`, `formatModelName`                |
| `tests/unit/normalizeEndpoint.test.ts`      | Endpoint URL normalization (trailing slashes, /openai suffixes, Foundry paths) |
| `tests/unit/officeStorage.test.ts`          | `officeStorage` localStorage fallback (OfficeRuntime undefined in jsdom)       |
| `tests/unit/parseFrontmatter.test.ts`       | YAML frontmatter parsing (skill files)                                         |
| `tests/unit/settingsStore.test.ts`          | Endpoint CRUD, model CRUD, cascade delete, auto-selection, dedup               |
| `tests/unit/toolResultSummary.test.ts`      | `toolResultSummary` — short human-readable summary from tool-result JSON       |
| `tests/unit/toolSchemas.test.ts`            | Zod inputSchema validation for all Excel tool definitions                      |

### Current Integration Test Files (12)

| File                                                     | Category           | Requires API? |
| -------------------------------------------------------- | ------------------ | ------------- |
| `tests/integration/agent-picker.test.tsx`                | Component wiring   | No            |
| `tests/integration/app-state.test.tsx`                   | Component wiring   | No            |
| `tests/integration/chat-header-settings-flow.test.tsx`   | Component wiring   | No            |
| `tests/integration/chat-panel.test.tsx`                  | Component wiring   | No            |
| `tests/integration/chat-pipeline.integration.test.ts`    | Live API           | Yes           |
| `tests/integration/foundry.integration.test.ts`          | Live API           | Yes           |
| `tests/integration/llm-tool-calling.integration.test.ts` | LLM tool selection | Yes           |
| `tests/integration/multi-turn.integration.test.ts`       | Live API           | Yes           |
| `tests/integration/settings-dialog.test.tsx`             | Component wiring   | No            |
| `tests/integration/skill-picker.test.tsx`                | Component wiring   | No            |
| `tests/integration/stale-state.test.tsx`                 | Store hydration    | No            |
| `tests/integration/wizard-to-chat.test.tsx`              | Component wiring   | No            |

### Integration Test Categories

- **Component wiring** — renders real components together (no child mocks)
- **Live API** — hits a real Azure AI Foundry endpoint (requires `.env` credentials)
- **LLM tool calling** — full ToolLoopAgent pipeline with live LLM

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
- Persisted via `officeStorage` adapter (auto-detects OfficeRuntime vs localStorage)
- Chat state lives in `useChat` (ephemeral, not persisted)
- `activeAgentId` and `activeSkillNames` (default: `null` = all ON) are persisted
- Persist storage key is `office-coding-agent-settings` (clean break from old Excel key)

### Tool Definitions

- Excel tools are defined across 9 config modules (range, table, chart, sheet, workbook, comment, conditionalFormat, dataValidation, pivotTable)
- Each config module in `src/tools/configs/` defines tool schemas and handlers
- Tool factory in `src/tools/codegen/factory.ts` generates tools from configs
- Host routing is in `src/tools/index.ts` via `getToolsForHost(host)`
- `src/tools/tools-manifest.json` — auto-generated manifest with all 83 tools

### UX Patterns

- **Copilot-style progress indicators** — cycling dot animation + phase labels (auto-derived via `humanizeToolName()`)
- **Choice cards** — `PromptStarterV2` renders `\`\`\`choices` blocks as clickable cards
- **Tool result summaries** — collapsible progress sections with `toolResultSummary()` one-liners
- **Input toolbar** — AgentPicker + ModelPicker below Composer (GitHub Copilot-style)

### Error Handling (OfficeRuntime)

- **NEVER use `declare const OfficeRuntime`** — it satisfies TypeScript but throws `ReferenceError` at runtime in vitest
- Use `typeof OfficeRuntime !== 'undefined'` guard instead (see `officeStorage.ts`)

## Build & Run

```bash
npm install
npm run dev            # Dev server with HMR
npm run build:dev      # Development build
npm run build          # Production build
npm run start:desktop  # Sideload into Excel
npm test               # All Vitest tests
npm run test:ui        # Playwright UI tests
npm run test:e2e       # E2E in Excel Desktop (~187 tests)
npm run validate       # Typecheck + lint + test in one command
```

## Key Files

- `src/taskpane/App.tsx` — root component, settings state, theme detection, routing, error boundary
- `src/hooks/useOfficeChat.ts` — main hook wiring host-routed AI agent
- `src/components/ChatHeader.tsx` — header: title, SkillPicker, new convo, Settings
- `src/components/ChatPanel.tsx` — messages, progress, input, AgentPicker, ModelPicker
- `src/components/ChatErrorBoundary.tsx` — error boundary around chat UI (keeps header functional on render errors)
- `src/components/AgentPicker.tsx` — single-select agent dropdown (Radix Popover)
- `src/components/ModelPicker.tsx` — model selection dropdown grouped by provider
- `src/components/SkillPicker.tsx` — icon-only skill toggle with badge count
- `src/components/SettingsDialog.tsx` — endpoint management dialog (controlled + uncontrolled)
- `src/components/SetupWizard.tsx` — first-time onboarding wizard
- `src/services/ai/aiClientFactory.ts` — creates Azure AI provider, normalizes endpoints
- `src/services/ai/chatService.ts` — alternative chat pipeline using streamText (used by multi-turn integration tests)
- `src/services/ai/modelDiscoveryService.ts` — discovers models, infers providers
- `src/services/ai/BASE_PROMPT.md` — universal base system prompt
- `src/services/ai/prompts/` — host-level app prompts (`EXCEL_APP_PROMPT.md`, `POWERPOINT_APP_PROMPT.md`)
- `src/services/office/host.ts` — Office host detection (`excel`, `powerpoint`, `unknown`)
- `src/services/agents/agentService.ts` — parses/filters host-targeted agent frontmatter and resolves host defaults
- `src/agents/excel/AGENT.md` — default Excel agent definition (host-targeted frontmatter)
- `src/services/skills/skillService.ts` — parses bundled skill files
- `src/stores/settingsStore.ts` — Zustand store with all CRUD operations
- `src/stores/officeStorage.ts` — OfficeRuntime.storage adapter with localStorage fallback
- `src/tools/` — 9 tool config modules + codegen factory (83 tools total)
- `src/utils/toolResultSummary.ts` — human-readable one-liner summaries for tool results
- `vitest.config.ts` — test config (jsdom, `@/` alias, setup file)
- `tests/setup.ts` — polyfills (ResizeObserver, IntersectionObserver, matchMedia)
