# Testing Strategy

> Detailed testing strategy for the excel-ai-addin project.

## The Excel.run() Boundary

The single most important architectural concept for testing this project is the **Excel.run() boundary**. Code that calls `Excel.run()` can only execute inside a real Excel instance. Everything else runs fine in Vitest with jsdom.

```
┌──────────────────────────────────────────────────────┐
│  Tier 1 — Testable with Vitest (no Excel needed)     │
│  ────────────────────────────────────────────         │
│  Pure functions:                                     │
│    normalizeEndpoint, inferProvider, formatModelName, │
│    isEmbeddingOrUtilityModel, parseFrontmatter,      │
│    buildSkillContext, toolResultSummary, generateId,  │
│    messagesToCoreMessages                            │
│  Zustand store logic:                                │
│    settingsStore (addEndpoint, removeEndpoint, etc.) │
│  Zod tool schemas:                                   │
│    inputSchema.safeParse() for all 9 tool modules    │
│  React components:                                   │
│    Full component wiring (integration tests)         │
│  AI client factory (with live API creds)             │
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

### Tier 1: Unit Tests (`tests/unit/`) — 12 files

**Runner:** Vitest with jsdom  
**No mocks, no Excel, no network.**

| File                             | What it covers                                                                                     | Pattern                           |
| -------------------------------- | -------------------------------------------------------------------------------------------------- | --------------------------------- |
| `agentService.test.ts`           | Agent frontmatter parsing, getAgents, getAgent, getAgentInstructions                               | Direct function calls             |
| `buildSkillContext.test.ts`      | `buildSkillContext` and related skill functions with bundled `.md` files                           | Direct function calls             |
| `chatPanel.test.tsx`             | ChatPanel component logic (mocks Fluent UI Copilot components for jsdom)                           | Component render + assertions     |
| `id.test.ts`                     | `generateId` unique ID generation utility                                                          | Direct function calls             |
| `messagesToCoreMessages.test.ts` | `messagesToCoreMessages` conversion from ChatMessage[] to core messages                            | Direct function calls             |
| `modelDiscoveryHelpers.test.ts`  | `inferProvider` → ModelProvider, `isEmbeddingOrUtilityModel` → boolean, `formatModelName` → string | Table-driven (`it.each`)          |
| `normalizeEndpoint.test.ts`      | URL normalization (slashes, `/openai` suffixes, Foundry paths)                                     | Direct function calls             |
| `officeStorage.test.ts`          | `officeStorage` localStorage fallback (OfficeRuntime undefined in jsdom)                           | Direct function calls             |
| `parseFrontmatter.test.ts`       | YAML frontmatter parsing (delimiters, multiline `>`, tag arrays, missing fields)                   | Direct function calls             |
| `settingsStore.test.ts`          | Endpoint CRUD, dedup by URL, cascade delete, model auto-selection, reset                           | Real Zustand store + localStorage |
| `toolResultSummary.test.ts`      | `toolResultSummary` — short human-readable summary from tool-result JSON                           | Direct function calls             |
| `toolSchemas.test.ts`            | Zod `inputSchema.safeParse()` for valid/invalid inputs on all 83 tool definitions                  | Schema validation only            |

**Key rules:**

- **DO NOT mock Zustand stores.** Test against the real store with localStorage (works in jsdom).
- **DO NOT mock pure functions.** Call them directly with test inputs.
- **Reset store state** in `beforeEach` via `useSettingsStore.getState().reset()`.
- **Use table-driven tests** (`it.each`) for functions with many input→output mappings.

### Tier 1b: Integration Tests (`tests/integration/`) — 12 files

**Runner:** Vitest with jsdom  
**Real components wired together, optionally real API calls.**

| File                                   | Category            | Requires API? |
| -------------------------------------- | ------------------- | ------------- |
| `agent-picker.test.tsx`                | Component wiring    | No            |
| `app-state.test.tsx`                   | Component wiring    | No            |
| `chat-header-settings-flow.test.tsx`   | Component wiring    | No            |
| `chat-panel.test.tsx`                  | Component wiring    | No            |
| `chat-pipeline.integration.test.ts`    | Live API round-trip | Yes           |
| `foundry.integration.test.ts`          | Live API round-trip | Yes           |
| `llm-tool-calling.integration.test.ts` | LLM tool selection  | Yes           |
| `multi-turn.integration.test.ts`       | Live API round-trip | Yes           |
| `settings-dialog.test.tsx`             | Component wiring    | No            |
| `skill-picker.test.tsx`                | Component wiring    | No            |
| `stale-state.test.tsx`                 | Store hydration     | No            |
| `wizard-to-chat.test.tsx`              | Component wiring    | No            |

**Key rules:**

- **No child mocks.** Render real components together to test cross-component interactions.
- Live API tests auto-skip when `FOUNDRY_ENDPOINT` / `FOUNDRY_API_KEY` are not set.
- Use longer timeouts for LLM tests (30s default in vitest config).

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

| Scenario                        | Test type        | Location                           |
| ------------------------------- | ---------------- | ---------------------------------- |
| New pure function               | Unit test        | `tests/unit/`                      |
| New Zustand store action        | Unit test        | `tests/unit/settingsStore.test.ts` |
| New Excel tool definition       | Schema test      | `tests/unit/toolSchemas.test.ts`   |
| New React component interaction | Integration test | `tests/integration/`               |
| New Excel command (`Excel.run`) | E2E test         | `tests-e2e/`                       |

## Adding a New Unit Test

1. Create a file in `tests/unit/` named `<module>.test.ts`.
2. Import the function under test from the barrel export (e.g., `@/services/ai`).
3. Use `describe` / `it` / `expect` directly (vitest globals are enabled).
4. For functions with many input→output mappings, use `it.each`.
5. Run `npx vitest run tests/unit/<module>.test.ts` to verify.

## Running Tests

```bash
# All Vitest tests (305 tests, 21 files)
npm test

# Only unit tests
npx vitest run tests/unit/

# Only integration tests
npx vitest run tests/integration/

# Watch mode
npm run test:watch

# E2E (requires Excel Desktop, ~187 tests)
npm run test:e2e

# Full validation (typecheck + lint + test)
npm run validate
```
