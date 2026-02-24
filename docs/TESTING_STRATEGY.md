# Testing Strategy

> Detailed testing strategy for the office-coding-agent project.

## The Host Runtime Boundary

The single most important architectural concept for testing this project is the **host runtime boundary**. Code that calls `Excel.run()`, PowerPoint, or Word APIs can only execute inside a real Office host instance. Everything else runs fine in Vitest with jsdom.

```
┌──────────────────────────────────────────────────────┐
│  Testable with Vitest/Playwright (no Office host)    │
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
│  • WebSocket client + session (mocked in tests)      │
│  • Agent/skill service parsing                       │
├──────────────────────────────────────────────────────┤
│  Excel.run() / PowerPoint / Word boundary            │
├──────────────────────────────────────────────────────┤
│  E2E only (Mocha + real Office Desktop)              │
│  ─────────────────────────────                       │
│  • rangeCommands, tableCommands, sheetCommands       │
│  • chartCommands, workbookCommands, commentCommands  │
│  • conditionalFormatCommands, dataValidationCommands │
│  • pivotTableCommands                                │
│  • PowerPoint / Word commands                        │
│  • OfficeRuntime.storage (real runtime)              │
└──────────────────────────────────────────────────────┘
```

## ⛔ CRITICAL RULE: DO NOT WRITE UNIT TESTS

Unit tests that mock Office APIs or fabricate fake contexts provide zero confidence that code works in a real host. They test the mock, not the code.
**Integration tests and E2E tests are the ONLY acceptable test forms for new functionality.**

- Writing a unit test when an integration or E2E test is possible is forbidden.
- If you are tempted to write a unit test, write an integration test instead.
- If the feature touches Office APIs, write an E2E test.

> `tests/unit/` is **empty** — all logic has been migrated to `tests/integration/`. There are no unit tests in this codebase.

## Test Tiers

### Integration Tests (`tests/integration/`)

**Runner:** Vitest with jsdom  
**Real components wired together; real store operations; live Copilot tests require a running dev server.**

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
| `copilot-custom-agent.integration.test.ts` | Live Copilot custom agent + skills | Yes              |
| `copilot-websocket.integration.test.ts` | Live Copilot WebSocket E2E          | Yes              |
| `excel-tools.test.ts`                   | Tool schema + factory (Excel)       | No               |
| `host-tools-limit.test.ts`             | Host tool count limits              | No               |
| `humanize-tool-name.test.ts`           | Tool-name → human-readable labels   | No               |
| `id.test.ts`                            | `generateId` utility                | No               |
| `management-tools.test.ts`             | Management tool schemas + handlers  | No               |
| `manifest.test.ts`                      | Office manifest / host assumptions  | No               |
| `mcp-manager-dialog.test.tsx`           | Component wiring                    | No               |
| `mcp-service.test.ts`                   | MCP server config parsing           | No               |
| `model-manager.test.tsx`                | Component wiring                    | No               |
| `model-picker-interactions.test.tsx`    | Component wiring                    | No               |
| `office-storage.test.ts`               | `officeStorage` with OfficeRuntime  | No               |
| `powerpoint-tools.test.ts`             | Tool schema + factory (PPT)         | No               |
| `settings-dialog.test.tsx`              | Component wiring                    | No               |
| `settings-store.test.ts`               | Zustand store (model/agent/skills)  | No               |
| `skill-manager-dialog.test.tsx`         | Component wiring                    | No               |
| `skill-picker.test.tsx`                 | Component wiring                    | No               |
| `skill-service.test.ts`                | Skill service + context building    | No               |
| `stale-state.test.tsx`                  | Store hydration                     | No               |
| `use-office-chat.test.tsx`             | useOfficeChat hook                  | No               |
| `use-tool-invocations-patch.test.tsx`  | Tool invocation argument streaming  | No               |
| `word-tools.test.ts`                    | Tool schema + factory (Word)        | No               |
| `zip-export-service.test.ts`            | ZIP export service                  | No               |
| `zip-import-service.test.ts`            | ZIP import service                  | No               |

**Key rules:**

- **DO NOT mock Zustand stores.** Test against the real store (jsdom with OfficeRuntime mock from `tests/setup.ts`).
- **DO NOT mock pure functions.** Call them directly with test inputs.
- **No child mocks.** Render real components together to test cross-component interactions.
- **Reset store state** in `beforeEach` via `useSettingsStore.getState().reset()`.
- **Use table-driven tests** (`it.each`) for functions with many input→output mappings.
- Live Copilot tests must run against a live server (`npm run dev`); do not add auto-skip behavior.
- Both `vitest.config.ts` and `vitest.integration.config.ts` must include `setupFiles: ['tests/setup.ts']` and `globals: true`.

### UI Tests (`tests-ui/`) — Playwright

**Runner:** Playwright  
**Browser task pane flows against the running dev server.**

### E2E Tests — Mocha inside real Office hosts

| Host       | Directory            | Tests  |
| ---------- | -------------------- | ------ |
| Excel      | `tests-e2e/`         | ~187   |
| PowerPoint | `tests-e2e-ppt/`     | ~13    |
| Word       | `tests-e2e-word/`    | ~12    |

**Real Office.js APIs, real host runtime.**

## When to Write What

| Scenario                                | Test type        | Location             |
| --------------------------------------- | ---------------- | -------------------- |
| New Excel command (`Excel.run`)         | E2E test         | `tests-e2e/`         |
| New PowerPoint command                  | E2E test         | `tests-e2e-ppt/`     |
| New Word command                        | E2E test         | `tests-e2e-word/`    |
| New task pane interaction flow          | UI test          | `tests-ui/`          |
| New React component or hook behavior   | Integration test | `tests/integration/` |
| New host routing rule                   | Integration test | `tests/integration/` |
| New tool definition                     | Integration test | `tests/integration/` |
| New pure function                       | Integration test | `tests/integration/` |

## Running Tests

```bash
# Integration tests
npm run test:integration

# Playwright UI tests
npm run test:ui

# E2E — requires Office host to be open
npm run test:e2e          # Excel Desktop (~187 tests)
npm run test:e2e:ppt      # PowerPoint Desktop (~13 tests)
npm run test:e2e:word     # Word Desktop (~12 tests)

# Validate manifest
npm run validate
```
