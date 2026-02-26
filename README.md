# Office Coding Agent

An Office add-in that embeds GitHub Copilot as an AI assistant in Excel, PowerPoint, Word, and Outlook. Built with React, [assistant-ui](https://github.com/assistant-ui/assistant-ui), Tailwind CSS, and the [GitHub Copilot SDK](https://www.npmjs.com/package/@github/copilot-sdk). The Copilot SDK integration architecture is based on [patniko/github-copilot-office](https://github.com/patniko/github-copilot-office). Requires an active GitHub Copilot subscription — no API keys or endpoint configuration needed.

> **Research Project Disclaimer**
>
> This repository is an independent **research project**. It is **not** affiliated with, endorsed by, sponsored by, or otherwise officially related to Microsoft or GitHub.

## How It Works

```
Office Task Pane (React + assistant-ui)
      ↓ WebSocket (wss://localhost:3000/api/copilot)
Node.js proxy server  (src/server.mjs)
      ↓ @github/copilot-sdk (manages CLI lifecycle internally)
GitHub Copilot API
```

The proxy server uses the `@github/copilot-sdk` to manage the Copilot CLI lifecycle and bridges it to the browser task pane via WebSocket + JSON-RPC. Tool calls flow back from the server to the browser, where host-specific handlers execute them (e.g., `Excel.run()`, `PowerPoint.run()`, `Word.run()`, or Outlook REST APIs).

## Features

- **GitHub Copilot authentication** — sign in once with your GitHub account; no API keys or endpoint config
- **Host-routed tools** — Excel, PowerPoint, Word, and Outlook toolsets selected by current Office host
- **10 Excel tool groups** — range, table, chart, sheet, workbook, comment, conditional format, data validation, pivot table, range format — covering ~83 actions
- **24 PowerPoint tools** — slides, shapes, text, images, tables, charts, notes, layouts; includes visual QA with `get_slide_image` region cropping for overflow detection
- **35 Word tools** — documents, paragraphs, tables, images, headers/footers, styles, comments, sections, fields, content controls
- **22 Outlook tools** — emails, calendar, contacts, folders, attachments, categories, search, flags, drafts
- **Agent system** — host-targeted agents with YAML frontmatter (`hosts`, `defaultForHosts`)
- **Skills system** — bundled skill files inject context into the system prompt, toggleable via SkillPicker
- **Custom agents & skills** — import local ZIP files for custom agents and skills
- **Model picker** — switch between supported Copilot models (Claude Sonnet, GPT-4.1, Gemini, etc.)
- **Streaming responses** — real-time token streaming with Copilot-style progress indicators
- **Auto-scroll chat** — thread stays pinned to newest content so follow-up output remains visible
- **Web fetch tool** — proxied through the local server to avoid CORS restrictions

## Agent Skills Format

A skill is a folder containing `SKILL.md`. Optional supporting docs live under `references/` inside that skill folder.

## Prerequisites

- [Node.js](https://nodejs.org/) >= 20
- Microsoft Office (Excel, PowerPoint, Word, or Outlook — desktop or Microsoft 365 web)
- An active **GitHub Copilot** subscription (individual, business, or enterprise)
- The `@github/copilot` CLI authenticated (`gh auth login` or equivalent)

## Getting Started

**👉 See [GETTING_STARTED.md](./GETTING_STARTED.md) for full setup instructions** — including authentication, starting the proxy server, registering the add-in, and sideloading into Office.

**Quick start** (requires [Node.js 20+](https://nodejs.org/), [GitHub CLI](https://cli.github.com/), and an active [GitHub Copilot](https://github.com/features/copilot) subscription):

```bash
# 1. Install dependencies
npm install

# 2. Authenticate with GitHub Copilot (once)
gh auth login

# 3. Register the add-in manifest + trust the SSL cert
npm run register:win    # Windows
npm run register:mac    # macOS

# 4. Terminal 1 — start the proxy server (keep this running)
npm run dev

# 5. Terminal 2 — sideload into Office
npm run start:desktop:excel   # or :ppt / :word
```

The proxy server runs on `https://localhost:3000` and handles both the Vite dev server UI and the Copilot WebSocket proxy. It must be running whenever you use the add-in.

For local shared-folder sideloading and staging manifest workflows, see [docs/SIDELOADING.md](./docs/SIDELOADING.md).

## Available Scripts

## Available Scripts

| Script                           | Description                                                           |
| -------------------------------- | --------------------------------------------------------------------- |
| `npm run dev`                    | Start Copilot proxy + Vite dev server (port 3000)                     |
| `npm run start:prod-server`      | Start production HTTPS server from `dist/`                            |
| `npm run start:tray`             | Build + run Electron system tray app                                  |
| `npm run start:tray:desktop`     | Start tray app (if needed) then sideload Excel desktop (legacy alias) |
| `npm run start:tray:excel`       | Start tray app (if needed) then sideload Excel desktop                |
| `npm run start:tray:ppt`         | Start tray app (if needed) then sideload PowerPoint desktop           |
| `npm run start:tray:word`        | Start tray app (if needed) then sideload Word desktop                 |
| `npm run stop:tray:desktop`      | Stop desktop sideload/debug session and server port 3000              |
| `npm run build:installer`        | Build desktop installer artifacts via electron-builder                |
| `npm run build:installer:win`    | Build Windows installer (NSIS)                                        |
| `npm run build:installer:dir`    | Build unpacked desktop app directory                                  |
| `npm run build`                  | Production build to `dist/`                                           |
| `npm run build:dev`              | Development build to `dist/`                                          |
| `npm run start:desktop`          | Sideload into Excel Desktop (legacy alias)                            |
| `npm run start:desktop:excel`    | Sideload into Excel Desktop                                           |
| `npm run start:desktop:ppt`      | Sideload into PowerPoint Desktop                                      |
| `npm run start:desktop:word`     | Sideload into Word Desktop                                            |
| `npm run stop`                   | Stop debugging / unload the add-in                                    |
| `npm run extensions:samples`     | Generate sample `agents` and `skills` ZIP files                       |
| `npm run sideload:share:setup`   | Create local shared-folder catalog on Windows                         |
| `npm run sideload:share:trust`   | Register local share as trusted Office catalog                        |
| `npm run sideload:share:publish` | Copy staging manifest into local shared folder                        |
| `npm run sideload:share:cleanup` | Remove local share and trusted-catalog setup                          |
| `npm run register:win`           | Trust cert and register manifest for Word/PPT/Excel (Windows)         |
| `npm run unregister:win`         | Remove registered manifest entry (Windows)                            |
| `npm run register:mac`           | Trust cert and register manifest for Word/PPT/Excel (macOS)           |
| `npm run unregister:mac`         | Remove manifest from Word/PPT/Excel WEF folders (macOS)               |
| `npm run lint`                   | Run ESLint                                                            |
| `npm run lint:fix`               | Auto-fix ESLint issues                                                |
| `npm run format`                 | Format code with Prettier                                             |
| `npm run typecheck`              | Type-check without emitting                                           |
| `npm test`                       | Run all Vitest suites                                                 |
| `npm run test:integration`       | Run integration test suite                                            |
| `npm run test:ui`                | Run Playwright UI tests                                               |
| `npm run test:watch`             | Run tests in watch mode                                               |
| `npm run test:coverage`          | Run tests with coverage                                               |
| `npm run test:e2e`               | Run E2E tests in Excel Desktop                                        |
| `npm run test:e2e:ppt`           | Run E2E tests in PowerPoint Desktop                                   |
| `npm run test:e2e:word`          | Run E2E tests in Word Desktop                                         |
| `npm run test:e2e:outlook`       | Run E2E tests in Outlook Desktop                                      |
| `npm run test:e2e:all`           | Run all four E2E suites in sequence                                   |
| `npm run validate`               | Validate `manifests/manifest.dev.xml`                                 |
| `npm run validate:outlook`       | Validate `manifests/manifest.outlook.dev.xml`                         |

## Testing

This project uses three active test layers:

- **Integration** (`tests/integration/`, Vitest) — component wiring, stores, host/tool routing, and live Copilot websocket flows
- **UI** (`tests-ui/`, Playwright) — browser taskpane behavior and regression coverage
- **E2E** (`tests-e2e*`, Mocha) — real Office host validation in Excel, PowerPoint, Word, and Outlook desktop

Unit tests are intentionally not used for new work in this repository.

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

# E2E tests (require Office desktop app)
npm run test:e2e
npm run test:e2e:ppt
npm run test:e2e:word
npm run test:e2e:outlook

# Validate the Office add-in manifest
npm run validate
```

Integration tests run as part of the default `npm test` suite.

## E2E Testing

The project includes end-to-end tests across all four Office hosts: ~187 Excel tests (tools, settings persistence, AI round-trips), ~13 PowerPoint tests, ~12 Word tests, and Outlook tests (requiring Exchange sideloading approval).

### How It Works

1. **Mocha runner** (`tests-e2e/runner.test.ts`) starts a local test server on port 4201.
2. A separate **test add-in** is built and served on `https://localhost:3001`.
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
     (test server)                         (Vite dev server)
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
src/server.mjs (Express HTTPS, port 3000)
src/copilotProxy.mjs → @github/copilot-sdk → GitHub Copilot API
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
- **`getToolsForHost(host)`** — returns `Tool[]` (Copilot SDK format) for the current Office host (Excel: ~83 tools, PowerPoint: 24, Word: 35, Outlook: 22)

State is minimal: `useSettingsStore` (Zustand) persists model/agent/skill configuration; chat state is ephemeral.

## UI Layout

The task pane is organized into three areas:

- **ChatHeader** — SkillPicker, Session History picker, Permissions button, and New Conversation action
- **ChatPanel** — thread/message stream, inline thinking indicator, composer, and input toolbar with AgentPicker + ModelPicker
- **App** — root shell that handles Office host detection, theme sync, and connection/session/permission banners

## Authentication

Authentication is handled entirely by the **GitHub Copilot CLI** (`@github/copilot` package). Run `gh auth login` once and the CLI handles OAuth token management. No API keys or Azure AD configuration is needed.

## Tech Stack

- **React 19** — UI framework
- **assistant-ui + Radix UI + Tailwind CSS v4** — task pane UI components and styling
- **GitHub Copilot SDK** (`@github/copilot-sdk`) — session management, streaming events, tool registration
- **WebSocket + JSON-RPC** (`vscode-jsonrpc`, `ws`) — browser-to-proxy transport
- **Express + HTTPS** — local proxy server with Vite dev middleware
- **Zustand 5** — lightweight state management with `OfficeRuntime.storage` persistence
- **Vite 7** — bundling with HMR
- **TypeScript 5** — type safety
- **Vitest** — integration testing
- **Playwright** — browser UI testing for task pane flows
- **Mocha** — E2E testing inside Excel Desktop (~187 tests)
- **Testing Library** — React component testing (`@testing-library/react`, `user-event`)
- **ESLint + Prettier** — code quality

## Project History

This project has gone through two major architectural phases:

### Phase 1 — Vercel AI SDK + Azure AI Foundry (Feb 16 2026)

The initial version of Office Coding Agent was built on the [Vercel AI SDK](https://ai-sdk.dev/) with [Azure AI Foundry](https://ai.azure.com/) as the model backend. It used `@ai-sdk/azure` and `@ai-sdk/react` along with `@assistant-ui/react-ai-sdk` for the chat UI. Users had to configure API endpoints, keys, and model deployments manually through a setup wizard.

### Phase 2 — GitHub Copilot SDK (Feb 20 2026 – present)

Inspired by [patniko/github-copilot-office](https://github.com/patniko/github-copilot-office) — a project by [Patrick Nikoletich](https://github.com/patniko), [Steve Sanderson](https://github.com/SteveSandersonMS), and [contributors](https://github.com/patniko/github-copilot-office/graphs/contributors) — the entire AI backend was replaced with the `@github/copilot-sdk` in [PR #25](https://github.com/sbroenne/office-coding-agent/pull/25). This migration:

- Replaced the Vercel AI SDK and Azure AI Foundry backend with the GitHub Copilot SDK
- Added a Node.js WebSocket proxy server (bridging the browser task pane to the Copilot CLI)
- Removed the setup wizard, API key configuration, and multi-provider endpoint management
- Simplified authentication to a single GitHub account sign-in via `gh auth login`

The proxy server architecture (`server.mjs` → `copilotProxy.mjs` → `@github/copilot-sdk`) and WebSocket-based browser transport were directly adopted from the patterns established in [patniko/github-copilot-office](https://github.com/patniko/github-copilot-office).

## Acknowledgments

- **[patniko/github-copilot-office](https://github.com/patniko/github-copilot-office)** — The proxy server architecture, Copilot SDK integration pattern, and WebSocket transport design used in this project were adopted from this repository by [Patrick Nikoletich](https://github.com/patniko) and [Steve Sanderson](https://github.com/SteveSandersonMS). Their work provided the foundation for the Phase 2 migration.
- **[@trsdn (Torsten)](https://github.com/trsdn)** and **[@urosstojkic](https://github.com/urosstojkic)** — Contributed the Word document orchestrator (planner→worker pattern), 22 Outlook tools, expanded PowerPoint tooling (24 tools), WorkIQ MCP stdio integration, host-specific welcome prompts, improved auto-scroll, and new skills (Outlook email/calendar/drafting, Word formatting/tables/document-builder, PowerPoint content/layout/animation/presentation). Originally submitted as [PR #33](https://github.com/sbroenne/office-coding-agent/pull/33) and merged in [PR #45](https://github.com/sbroenne/office-coding-agent/pull/45).
- **[assistant-ui](https://github.com/assistant-ui/assistant-ui)** — React chat UI components used for the task pane thread and composer.
- **[Vercel AI SDK](https://ai-sdk.dev/)** — Original AI runtime used in Phase 1.

## Community & Security

- [Code of Conduct](./CODE_OF_CONDUCT.md)
- [Security Policy](./SECURITY.md)

## License

MIT
