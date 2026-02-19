# Office Coding Agent

An Office add-in that embeds GitHub Copilot as an AI assistant in Excel (and other Office hosts). Built with React, [assistant-ui](https://github.com/assistant-ui/assistant-ui), Tailwind CSS, and the [GitHub Copilot SDK](https://github.com/patniko/github-copilot-office). Requires an active GitHub Copilot subscription — no API keys or endpoint configuration needed.

> **Research Project Disclaimer**
>
> This repository is an independent **research project**. It is **not** affiliated with, endorsed by, sponsored by, or otherwise officially related to Microsoft or GitHub.

## How It Works

```
Excel Task Pane (React + assistant-ui)
        ↓ WebSocket (wss://localhost:3000/api/copilot)
Node.js proxy server  (src/server.js)
        ↓ stdio LSP
@github/copilot CLI  (authenticates via your GitHub account)
        ↓ HTTPS
GitHub Copilot API
```

The proxy server spawns the `@github/copilot` CLI process and bridges it to the browser task pane via WebSocket + JSON-RPC. Tool calls (Excel commands) flow back from the server to the browser.

## Features

- **GitHub Copilot authentication** — sign in once with your GitHub account; no API keys or endpoint config
- **10 Excel tool groups** — range, table, chart, sheet, workbook, comment, conditional format, data validation, pivot table, range format — covering ~83 actions the AI can perform
- **Agent system** — host-targeted agents with YAML frontmatter (`hosts`, `defaultForHosts`)
- **Skills system** — bundled skill files inject context into the system prompt, toggleable via SkillPicker
- **Custom agents & skills** — import local ZIP files for custom agents and skills
- **Model picker** — switch between supported Copilot models (Claude Sonnet, GPT-4.1, Gemini, etc.)
- **Streaming responses** — real-time token streaming with Copilot-style progress indicators
- **Web fetch tool** — proxied through the local server to avoid CORS restrictions

## Agent Skills Format

A skill is a folder containing `SKILL.md`. Optional supporting docs live under `references/` inside that skill folder.

## Prerequisites

- [Node.js](https://nodejs.org/) >= 20
- Microsoft Excel (desktop or Microsoft 365 web)
- An active **GitHub Copilot** subscription (individual, business, or enterprise)
- The `@github/copilot` CLI authenticated (`gh auth login` or equivalent)

## Getting Started

```bash
# Install dependencies
npm install

# Terminal 1: start the Copilot proxy server
npm run server

# Terminal 2: sideload into Excel Desktop
npm run start:desktop
```

The proxy server runs on `https://localhost:3000` and handles both the webpack-dev-server UI and the WebSocket Copilot proxy.

For local shared-folder sideloading and staging manifest workflows, see [docs/SIDELOADING.md](./docs/SIDELOADING.md).

## Available Scripts

## Available Scripts

| Script                           | Description                                           |
| -------------------------------- | ----------------------------------------------------- |
| `npm run server`                 | Start Copilot proxy + webpack dev server (port 3000)  |
| `npm run dev`                    | Start webpack-dev-server only (UI, no Copilot proxy)  |
| `npm run build`                  | Production build to `dist/`                           |
| `npm run build:dev`              | Development build to `dist/`                          |
| `npm run start:desktop`          | Sideload into Excel Desktop                           |
| `npm run stop`                   | Stop debugging / unload the add-in                    |
| `npm run extensions:samples`     | Generate sample `agents` and `skills` ZIP files       |
| `npm run sideload:share:setup`   | Create local shared-folder catalog on Windows         |
| `npm run sideload:share:trust`   | Register local share as trusted Office catalog        |
| `npm run sideload:share:publish` | Copy staging manifest into local shared folder        |
| `npm run sideload:share:cleanup` | Remove local share and trusted-catalog setup          |
| `npm run lint`                   | Run ESLint                                            |
| `npm run lint:fix`               | Auto-fix ESLint issues                                |
| `npm run format`                 | Format code with Prettier                             |
| `npm run typecheck`              | Type-check without emitting                           |
| `npm test`                       | Run all Vitest tests                                  |
| `npm run test:watch`             | Run tests in watch mode                               |
| `npm run test:coverage`          | Run tests with coverage                               |
| `npm run test:e2e`               | Run E2E tests in Excel Desktop (~187)                 |
| `npm run validate`               | Validate `manifests/manifest.dev.xml`                 |

## Testing

The project has four layers of tests:

| Layer           | Tool       | Count    | What it covers                                                          |
| --------------- | ---------- | -------- | ----------------------------------------------------------------------- |
| **Unit**        | Vitest     | ~12 files | Pure functions, Zustand store, JSON Schema tool configs, host/agent parsing |
| **Integration** | Vitest     | ~10 files | Component wiring (no live API needed)                                   |
| **UI**          | Playwright | ~14 tests | Browser taskpane flows                                                  |
| **E2E**         | Mocha      | ~187     | Excel commands inside real Excel Desktop                                |

### Running Tests

```bash
# All Vitest tests
npm test

# Watch mode
npm run test:watch

# With coverage
npm run test:coverage

# Browser UI tests
npm run test:ui

# E2E tests (requires Excel Desktop)
npm run test:e2e

# Validate the Office add-in manifest
npm run validate
```

### Unit Tests

Unit tests in `tests/unit/` cover pure functions and store logic with **no** `Excel.run()` dependency:

| File                                | What it tests                                                              |
| ----------------------------------- | -------------------------------------------------------------------------- |
| `agentService.test.ts`              | Agent frontmatter parsing, getAgents, getAgent, getAgentInstructions       |
| `buildSkillContext.test.ts`         | `buildSkillContext` and related skill functions with bundled `.md` files   |
| `chatErrorBoundary.test.tsx`        | Error boundary fallback rendering and recovery flow                        |
| `chatPanel.test.tsx`                | ChatPanel component logic (mocks assistant-ui components for jsdom)        |
| `humanizeToolName.test.ts`          | Tool-name formatting for user-facing progress labels                       |
| `id.test.ts`                        | `generateId` unique ID generation utility                                  |
| `manifest.test.ts`                  | Manifest and runtime host assumptions used by tests                        |
| `officeStorage.test.ts`             | `officeStorage` localStorage fallback (OfficeRuntime undefined in jsdom)   |
| `officeStorageRuntime.test.ts`      | `officeStorage` behavior when OfficeRuntime is present and throws          |
| `parseFrontmatter.test.ts`          | YAML frontmatter parsing for skill files (delimiters, arrays)              |
| `settingsStore.test.ts`             | Zustand store: activeModel, agent/skill management                         |
| `toolSchemas.test.ts`               | JSON Schema validation for all tool definitions                            |
| `useToolInvocations-patch.test.tsx` | assistant-ui patch for tool invocation argument streaming integrity        |

**Key principle:** unit tests run against the **real** Zustand store with localStorage (jsdom). No mocking.

### Integration Tests

Integration tests in `tests/integration/` exercise **component wiring only** (no live API needed):

- `agent-picker.test.tsx`, `skill-picker.test.tsx`, `chat-panel.test.tsx`, `chat-header-settings-flow.test.tsx`
- `app-state.test.tsx`, `app-error-boundary.test.tsx`, `model-picker-interactions.test.tsx`
- `stale-state.test.tsx` — store hydration recovery from stale localStorage

Integration tests run as part of the default `npm test` suite.

## E2E Testing

The project includes ~187 end-to-end tests that validate all 83 Excel tools plus settings persistence and AI round-trips inside a real Excel Desktop instance.

### How It Works

1. **Mocha runner** (`tests-e2e/runner.test.ts`) starts a local test server on port 4201.
2. A separate **test add-in** is built by webpack and served on `https://localhost:3001`.
3. The test add-in is **sideloaded into Excel Desktop** using `office-addin-debugging`.
4. Inside Excel, `test-taskpane.ts` runs the Excel command tests and **sends results back** to the test server.
5. The Mocha runner **receives the results** and asserts on them.

### Tool Coverage

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

### Running E2E Tests

```bash
npm run test:e2e
```

This command:

- Uses `ts-node` with the `tests-e2e/tsconfig.json` project
- Starts the test server, builds the test add-in, sideloads into Excel
- Waits for all test results, then tears down (closes Excel, stops server)

> **Note:** E2E tests require Excel Desktop installed on the machine. They use a separate manifest (`tests-e2e/test-manifest.xml`) with its own GUID so it can coexist with the dev add-in.

### Architecture

```
┌─────────────────────┐    results     ┌─────────────────────┐
│  Mocha Runner       │◄───(port 4201)──│  Test Taskpane      │
│  (Node.js)          │                 │  (inside Excel)     │
│                     │  sideload       │                     │
│  - starts server    │────────────────►│  - writes ranges    │
│  - builds add-in    │                 │  - creates sheets   │
│  - asserts results  │                 │  - manages tables   │
│                     │                 │  - creates charts   │
└─────────────────────┘                 └─────────────────────┘
        port 4201                              port 3001
     (test server)                         (webpack dev server)
```

### Adding a New E2E Test

1. Add test logic in `tests-e2e/src/test-taskpane.ts` using the `pass()`/`fail()`/`assert()` helpers.
2. Add a corresponding Mocha `it()` block in `tests-e2e/runner.test.ts` that reads the result via `e2eContext.getResult('your_test_name')`.
3. Run `npm run test:e2e` to verify.

## Chat Architecture

The add-in routes messages through a **local proxy server** — the browser task pane cannot call the GitHub Copilot API directly due to browser security restrictions.

```
useOfficeChat(host)
      ↓ createWebSocketClient(wss://localhost:3000/api/copilot)
BrowserCopilotSession.query({ prompt, tools })
      ↓ SessionEvent stream
assistant.message_delta / tool.* / session.idle
      ↓
ThreadMessage[] → useExternalStoreRuntime
      ↓ wss://localhost:3000/api/copilot
src/server.js (Express HTTPS, port 3000)
src/copilotProxy.js → @github/copilot CLI → GitHub Copilot API
```

### Agent System

The AI agent uses a **split system prompt** architecture:

- **`src/services/ai/BASE_PROMPT.md`** — universal base prompt (progress narration, presenting choices)
- **`src/services/ai/prompts/*_APP_PROMPT.md`** — host-level app prompts
- **`src/agents/*/AGENT.md`** — agent-specific instructions with YAML frontmatter
- Instructions = `buildSystemPrompt(host) + resolvedAgent.instructions + skillContext`

`agentService` parses and filters agents by host. Agents are targeted via frontmatter `hosts` and can declare host defaults via `defaultForHosts`.

### Skills and Agents

- Bundled skills/agents are a separate immutable category (read-only in UI).
- Custom skills/agents are imported locally from ZIP files via picker management dialogs.
- Imported items are persisted in settings storage and can be removed from the same management dialogs.

#### Skills ZIP format (folder inference)

Use a ZIP containing markdown files under `skills/`:

```text
my-skills.zip
└── skills/
    ├── security-review.md
    └── finance-helper.md
```

Each skill markdown file must include frontmatter:

```markdown
---
name: Security Review
description: Threat-model and security-review guidance.
version: 1.0.0
tags:
  - security
  - review
---

Your skill instructions go here.
```

#### Agents ZIP format (folder inference)

Use a ZIP containing markdown files under `agents/`:

```text
my-agents.zip
└── agents/
    ├── excel-data-analyst.md
    └── powerpoint-storyteller.md
```

Each agent markdown file must include frontmatter with supported hosts:

```markdown
---
name: Excel Data Analyst
description: Focused Excel analysis assistant.
version: 1.0.0
hosts: [excel]
defaultForHosts: []
---

Your agent instructions go here.
```

Notes:

- Skills ZIP and Agents ZIP are imported separately.
- If an imported item name collides with an existing one, the add-in keeps both and auto-suffixes the imported name.
- Use `npm run extensions:samples` to generate starter ZIP files under `samples/extensions/`.

#### Import and Management UX

In picker management dialogs:

- Open **Skill picker → Manage skills…** for skill import/removal
- Open **Agent picker → Manage agents…** for agent import/removal
- Import custom entries via **Import Skills ZIP** / **Import Agents ZIP**
- Remove imported entries directly from Imported lists
- View bundled entries separately in read-only sections

In chat pickers:

- Agent and skill lists are grouped by source (**Bundled** vs **Imported**)
- Bundled category remains immutable; only imported entries are removable via Settings

### Key Hooks and Components

- **`useOfficeChat`** — creates a `WebSocketCopilotClient`, opens a `BrowserCopilotSession`, maps `SessionEvent` stream to `ThreadMessage[]` for `useExternalStoreRuntime`
- **`BrowserCopilotSession.query()`** — async generator yielding `SessionEvent` objects (assistant.message_delta, tool.execution_start, session.idle, etc.)
- **`getToolsForHost(host)`** — returns `Tool[]` (Copilot SDK format) for the current Office host

State is minimal: `useSettingsStore` (Zustand) persists model/agent/skill configuration; chat state is ephemeral.

## UI Layout

The task pane is organized into three areas:

- **ChatHeader** — "AI Chat" title + SkillPicker (icon-only with badge) + New Conversation button + Settings gear
- **ChatPanel** — CopilotChat messages, Copilot-style progress indicators (cycling dots + phase labels), choice cards, error bar, ChatInput, and an **input toolbar** below the text box with AgentPicker + ModelPicker (GitHub Copilot-style)
- **App** — root component that owns settings dialog state, detects system theme and Office host

## Authentication

Authentication is handled entirely by the **GitHub Copilot CLI** (`@github/copilot` package). Run `gh auth login` once and the CLI handles OAuth token management. No API keys or Azure AD configuration is needed.

## Tech Stack

- **React 18** — UI framework
- **assistant-ui + Radix UI + Tailwind CSS v4** — task pane UI components and styling
- **GitHub Copilot SDK** (`@github/copilot-sdk`) — session management, streaming events, tool registration
- **WebSocket + JSON-RPC** (`vscode-jsonrpc`, `ws`) — browser-to-proxy transport
- **Express + HTTPS** — local proxy server with webpack-dev-middleware
- **Zustand 5** — lightweight state management with `OfficeRuntime.storage` persistence
- **Webpack 5** — bundling with HMR
- **TypeScript 5** — type safety
- **Vitest** — unit, component, and integration testing
- **Playwright** — browser UI testing for task pane flows
- **Mocha** — E2E testing inside Excel Desktop (~187 tests)
- **Testing Library** — React component testing (`@testing-library/react`, `user-event`)
- **ESLint + Prettier** — code quality

## Community & Security

- [Code of Conduct](./CODE_OF_CONDUCT.md)
- [Security Policy](./SECURITY.md)

## License

MIT
