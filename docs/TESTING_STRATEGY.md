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

| File                                       | Category                            | Requires server? |
| ------------------------------------------ | ----------------------------------- | ---------------- |
| `agent-manager-dialog.test.tsx`            | Component wiring                    | No               |
| `agent-picker.test.tsx`                    | Component wiring                    | No               |
| `agent-service.test.ts`                    | Agent service + frontmatter parsing | No               |
| `app-error-boundary.test.tsx`              | Component wiring                    | No               |
| `app-session-error.test.tsx`               | Component wiring                    | No               |
| `app-state.test.tsx`                       | Component wiring                    | No               |
| `chat-error-boundary.test.tsx`             | Component wiring                    | No               |
| `chat-header-settings-flow.test.tsx`       | Component wiring                    | No               |
| `chat-panel.test.tsx`                      | Component wiring                    | No               |
| `cli-plugin-bootstrap.test.ts`             | CLI plugin startup bootstrap        | No               |
| `chat-store.test.ts`                       | Chat message store                  | No               |
| `copilot-custom-agent.integration.test.ts` | Live Copilot custom agent + skills  | Yes              |
| `copilot-websocket.integration.test.ts`    | Live Copilot WebSocket E2E          | Yes              |
| `excel-tools.test.ts`                      | Tool schema + factory (Excel)       | No               |
| `host-tools-limit.test.ts`                 | Host tool count limits              | No               |
| `humanize-tool-name.test.ts`               | Tool-name → human-readable labels   | No               |
| `id.test.ts`                               | `generateId` utility                | No               |
| `management-tools.test.ts`                 | Management tool schemas + handlers  | No               |
| `manifest.test.ts`                         | Office manifest / host assumptions  | No               |
| `model-manager.test.tsx`                   | Component wiring                    | No               |
| `model-picker-interactions.test.tsx`       | Component wiring                    | No               |
| `office-storage.test.ts`                   | `officeStorage` with OfficeRuntime  | No               |
| `powerpoint-tools.test.ts`                 | Tool schema + factory (PPT)         | No               |
| `settings-dialog.test.tsx`                 | Component wiring                    | No               |
| `settings-store.test.ts`                   | Zustand store (model/agent/skills)  | No               |
| `skill-service.test.ts`                    | Skill service + context building    | No               |
| `stale-state.test.tsx`                     | Store hydration                     | No               |
| `use-office-chat.test.tsx`                 | useOfficeChat hook                  | No               |
| `use-tool-invocations-patch.test.tsx`      | Tool invocation argument streaming  | No               |
| `word-tools.test.ts`                       | Tool schema + factory (Word)        | No               |
| `zip-export-service.test.ts`               | ZIP export service                  | No               |
| `zip-import-service.test.ts`               | ZIP import service                  | No               |

**Key rules:**

- **DO NOT mock Zustand stores.** Test against the real store (jsdom with OfficeRuntime mock from `tests/setup.ts`).
- **DO NOT mock pure functions.** Call them directly with test inputs.
- **No child mocks.** Render real components together to test cross-component interactions.
- **Reset store state** in `beforeEach` via `useSettingsStore.getState().reset()`.
- **Use table-driven tests** (`it.each`) for functions with many input→output mappings.
- Live Copilot tests must run against a live server (`npm run dev`); do not add auto-skip behavior.
- Both the `unit` and `integration` projects in `vitest.config.ts` must include `setupFiles: ['tests/setup.ts']` and `globals: true`.

### AI / Skill-Engineering Tests (`tests-aitest/`)

**Runner:** Python pytest via `uv`  
**Purpose:** Validate that an LLM can correctly understand and call the tool schemas — i.e. that the schemas are *intelligible* to a model, not just structurally valid.

> ⚠️ These tests make **real LLM API calls** and cost money. Use `--lf` to re-run failures only.

#### Architecture

```
Real LLM (Azure OpenAI / gpt-5-mini by default)
    ↓ receives production system prompt + tool schemas (via manifest JSON)
    ↓ given a natural-language task ("write values to A1:B2, then read them back")
    ↓ calls tools via MCP
    ↓
In-memory simulator (ExcelSimulator / PowerPointSimulator / etc.)
    ↑ executes tool calls, returns results — no real Office host required
```

| File | What it does |
|------|-------------|
| `conftest.py` | Fixtures, model defaults (`gpt-5-mini`), manifest paths, production system prompt loader |
| `excel_sim.py` / `powerpoint_sim.py` / `word_sim.py` / `outlook_sim.py` | In-memory simulators — minimal spreadsheet/presentation/document/mailbox engines |
| `excel_mcp.py` / `powerpoint_mcp.py` / `word_mcp.py` / `outlook_mcp.py` | MCP servers wired to each simulator |
| `test_excel_tools.py` | Excel tool schema eval — range read/write, tables, charts, formatting |
| `test_powerpoint_tools.py` | PowerPoint tool schema eval |
| `test_word_tools.py` | Word tool schema eval |
| `test_outlook_tools.py` | Outlook tool schema eval |
| `test_excel_adversarial.py` | Adversarial cases that probe edge cases and ambiguous instructions |
| `test_token_efficiency.py` | Validates tool schemas don't bloat the context window |
| `manifests/` | Static JSON manifests (generated by `npm run manifest`) listing all tools per host |

#### Prerequisites

- Python environment via `uv` (`uv sync`)
- `AZURE_OPENAI_ENDPOINT` and `AZURE_OPENAI_API_KEY` env vars set (or equivalent LiteLLM-compatible provider config)
- Manifests generated: `npm run manifest`

#### Running

```bash
# All AI tests
uv run pytest tests-aitest/ -v

# Re-run only last failures (saves cost)
uv run pytest tests-aitest/ -v --lf

# Specific host
uv run pytest tests-aitest/ -v -m excel
uv run pytest tests-aitest/ -v -m powerpoint
uv run pytest tests-aitest/ -v -m adversarial
```

#### Pytest markers

| Marker | Tests |
|--------|-------|
| `integration` | All live LLM tests |
| `excel` | Excel host tests |
| `powerpoint` | PowerPoint host tests |
| `word` | Word host tests |
| `outlook` | Outlook host tests |
| `adversarial` | Adversarial edge-case evals |
| `token_efficiency` | Token budget experiments |

### UI Tests (`tests-ui/`) — Playwright

**Runner:** Playwright  
**Browser task pane flows against the running dev server.**

### E2E Tests — Mocha inside real Office hosts

| Host       | Directory            | Tests |
| ---------- | -------------------- | ----- |
| Excel      | `tests-e2e/`         | ~233  |
| PowerPoint | `tests-e2e-ppt/`     | ~15   |
| Word       | `tests-e2e-word/`    | ~14   |
| Outlook    | `tests-e2e-outlook/` | ~8    |

**Real Office.js APIs, real host runtime.**

## When to Write What

| Scenario                             | Test type        | Location             |
| ------------------------------------ | ---------------- | -------------------- |
| New Excel command (`Excel.run`)      | E2E test         | `tests-e2e/`         |
| New PowerPoint command               | E2E test         | `tests-e2e-ppt/`     |
| New Word command                     | E2E test         | `tests-e2e-word/`    |
| New Outlook command                  | E2E test         | `tests-e2e-outlook/` |
| New task pane interaction flow       | UI test          | `tests-ui/`          |
| New React component or hook behavior | Integration test | `tests/integration/` |
| New host routing rule                | Integration test | `tests/integration/` |
| New tool definition                  | Integration test | `tests/integration/` |
| New pure function                    | Integration test | `tests/integration/` |
| New or changed tool schema           | AI eval test     | `tests-aitest/`      |
| Tool schema may confuse the LLM      | Adversarial test | `tests-aitest/`      |

## Running Tests

```bash
# AI / skill-engineering tests (requires Azure OpenAI credentials + manifests)
uv run pytest tests-aitest/ -v
uv run pytest tests-aitest/ -v --lf   # re-run failures only (saves cost)

# Integration tests
npm run test:integration

# Playwright UI tests
npm run test:ui

# E2E — requires Office host to be open
npm run test:e2e          # Excel Desktop (~233 tests)
npm run test:e2e:ppt      # PowerPoint Desktop (~15 tests)
npm run test:e2e:word     # Word Desktop (~14 tests)
npm run test:e2e:outlook  # Outlook Desktop (~8 tests; requires Exchange sideloading approval)
npm run test:e2e:all      # All four suites in sequence

# Validate manifest
npm run validate
```

