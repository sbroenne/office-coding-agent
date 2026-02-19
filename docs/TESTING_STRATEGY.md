# Testing Strategy

> Detailed testing strategy for the office-coding-agent project.

## The Excel.run() Boundary

The single most important architectural concept for testing this project is the **Excel.run() boundary**. Code that calls `Excel.run()` can only execute inside a real Excel instance. Everything else runs fine in Vitest with jsdom.

```
┌──────────────────────────────────────────────────────┐
│  Tier 1 — Testable with Vitest (no Excel needed)     │
│  ────────────────────────────────────────────         │
│  Pure functions:                                     │
│    parseFrontmatter, buildSkillContext,               │
│    toolResultSummary, generateId,                    │
│    humanizeToolName, zipImportService                │
│  Zustand store logic:                                │
│    settingsStore (activeModel, agents, skills)       │
│  JSON Schema tool configs:                           │
│    toCopilotTools() for all 9 tool modules           │
│  React components:                                   │
│    Full component wiring (integration tests)         │
│  useOfficeChat hook (mocked WebSocket session)       │
│  Agent/skill service parsing                         │
├──────────────────────────────────────────────────────┤
│  Excel.run() boundary                                │
├──────────────────────────────────────────────────────┤
│  Tier 2 — E2E only (Mocha + real Excel Desktop)      │
│  ────────────────────────────────────────────         │
│  services/excel/*Commands.ts                         │
│    rangeCommands, tableCommands, sheetCommands,      │
│    chartCommands, workbookCommands, commentCommands,  │
│    conditionalFormatCommands, dataValidationCommands, │
│    pivotTableCommands                                │
│  OfficeRuntime.storage (real Office runtime)         │
│  Full agent loop inside Excel                        │
└──────────────────────────────────────────────────────┘
```

## Test Tiers

### Tier 1: Unit Tests (`tests/unit/`) — 18 files

**Runner:** Vitest with jsdom  
**No mocks, no Excel, no network.**

| File                                  | What it covers                                                                 | Pattern                           |
| ------------------------------------- | ------------------------------------------------------------------------------ | --------------------------------- |
| `agentService.test.ts`                | Agent frontmatter parsing, getAgents, getAgent, getAgentInstructions           | Direct function calls             |
| `buildSkillContext.test.ts`           | `buildSkillContext` and related skill functions with bundled `.md` files       | Direct function calls             |
| `chatErrorBoundary.test.tsx`          | Error boundary fallback rendering and recovery flow                            | Component render + assertions     |
| `chatPanel.test.tsx`                  | ChatPanel component logic (mocks assistant-ui components for jsdom)            | Component render + assertions     |
| `chatStore.test.ts`                   | Chat message store: append, clear, tool invocations                            | Direct store calls                |
| `generalTools.test.ts`               | General-purpose tool definitions (web_fetch, etc.)                             | Schema + handler assertions       |
| `hostToolsLimit.test.ts`             | Host tool count limits per host                                                | Direct function calls             |
| `humanizeToolName.test.ts`           | Tool-name → human-readable progress label formatting                          | Table-driven (`it.each`)          |
| `id.test.ts`                          | `generateId` unique ID generation utility                                      | Direct function calls             |
| `manifest.test.ts`                    | Office manifest / runtime host assumptions used by tests                       | Assertions                        |
| `mcpService.test.ts`                  | MCP server config parsing; HTTP/SSE transport filtering (no stdio)             | Direct function calls             |
| `officeStorage.test.ts`               | `officeStorage` with `OfficeRuntime` mock (via tests/setup.ts)                 | Store assertions                  |
| `parseFrontmatter.test.ts`            | YAML frontmatter parsing (delimiters, multiline, tag arrays)                   | Direct function calls             |
| `settingsStore.test.ts`               | Zustand store: activeModel, agent/skill CRUD, reset                            | Real Zustand store + OfficeRuntime mock |
| `toolSchemas.test.ts`                 | JSON Schema validation for all tool definitions (toCopilotTools)               | Schema validation only            |
| `useOfficeChat.test.tsx`              | useOfficeChat hook: mocked WebSocket session → ThreadMessage[] mapping         | renderHook + mocked client        |
| `useToolInvocations-patch.test.tsx`   | assistant-ui patch for tool invocation argument streaming integrity            | Component assertions              |
| `zipImportService.test.ts`            | ZIP import service for custom agents/skills                                    | Direct function calls             |

**Key rules:**

- **DO NOT mock Zustand stores.** Test against the real store (jsdom with OfficeRuntime mock from `tests/setup.ts`).
- **DO NOT mock pure functions.** Call them directly with test inputs.
- **Reset store state** in `beforeEach` via `useSettingsStore.getState().reset()`.
- **Use table-driven tests** (`it.each`) for functions with many input→output mappings.

### Tier 1b: Integration Tests (`tests/integration/`) — 12 files

**Runner:** Vitest with jsdom  
**Real components wired together; live Copilot WebSocket test auto-skips without a server.**

| File                                           | Category                   | Requires server? |
| ---------------------------------------------- | -------------------------- | ---------------- |
| `agent-picker.test.tsx`                        | Component wiring           | No               |
| `app-error-boundary.test.tsx`                  | Component wiring           | No               |
| `app-state.test.tsx`                           | Component wiring           | No               |
| `chat-header-settings-flow.test.tsx`           | Component wiring           | No               |
| `chat-panel.test.tsx`                          | Component wiring           | No               |
| `copilot-websocket.integration.test.ts`        | Live Copilot WebSocket E2E | Yes (auto-skip)  |
| `mcp-manager-dialog.test.tsx`                  | Component wiring           | No               |
| `model-manager.test.tsx`                       | Component wiring           | No               |
| `model-picker-interactions.test.tsx`           | Component wiring           | No               |
| `settings-dialog.test.tsx`                     | Component wiring           | No               |
| `skill-picker.test.tsx`                        | Component wiring           | No               |
| `stale-state.test.tsx`                         | Store hydration            | No               |

**Key rules:**

- **No child mocks.** Render real components together to test cross-component interactions.
- The live Copilot WebSocket test (`copilot-websocket.integration.test.ts`) auto-skips when `npm run server` is not running.
- Both `vitest.config.ts` and `vitest.integration.config.ts` must include `setupFiles: ['tests/setup.ts']` and `globals: true`.

### Tier 2: E2E Tests (`tests-e2e/`) — ~187 tests

**Runner:** Mocha inside Excel Desktop  
**Real Excel.run(), real Office.js APIs.**

Tests in this tier exercise the actual Excel commands across 9 tool categories:

| Category             | Tests |
| -------------------- | ----- |
| Range Tools          | 52    |
| Table Tools          | 15    |
| Chart Tools          | 14    |
| Sheet Tools          | 22    |
| Workbook Tools       | 8     |
| Comment Tools        | 8     |
| Conditional Format   | 27    |
| Data Validation      | 21    |
| Pivot Table          | 10    |
| Settings Persistence | 4     |
| AI Round-Trip        | 5     |
| Summary              | 1     |

See [the E2E section in README.md](../README.md#e2e-testing) for setup.

## When to Write What

| Scenario                           | Test type        | Location                              |
| ---------------------------------- | ---------------- | ------------------------------------- |
| New pure function                  | Unit test        | `tests/unit/`                         |
| New Zustand store action           | Unit test        | `tests/unit/settingsStore.test.ts`    |
| New tool definition                | Schema test      | `tests/unit/toolSchemas.test.ts`      |
| New hook behavior                  | Unit test        | `tests/unit/` (mock the WebSocket client) |
| New React component interaction    | Integration test | `tests/integration/`                  |
| New Excel command (`Excel.run`)    | E2E test         | `tests-e2e/`                          |

## Adding a New Unit Test

1. Create a file in `tests/unit/` named `<module>.test.ts`.
2. Import the function under test from the barrel export (e.g., `@/stores`, `@/tools`).
3. Use `describe` / `it` / `expect` directly (vitest globals are enabled).
4. For functions with many input→output mappings, use `it.each`.
5. Run `npx vitest run tests/unit/<module>.test.ts` to verify.

## Running Tests

```bash
# All Vitest unit tests (265 tests, 30 files including integration)
npm test

# Only integration tests (47 tests, 12 files)
npm run test:integration

# Watch mode
npm run test:watch

# E2E (requires Excel Desktop, ~187 tests)
npm run test:e2e

# Validate manifest
npm run validate
```