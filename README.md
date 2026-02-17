# Office Coding Agent

An Office add-in research project that provides an AI-powered coding assistant using Azure AI Foundry models. The runtime is host-routed (tools and prompts per application), with Excel as the first supported host. Built with React, assistant-ui, Radix UI, Tailwind CSS, and the [Vercel AI SDK](https://ai-sdk.dev/). It supports per host Agent Skills and custom agents.

> **Research Project Disclaimer**
>
> This repository is an independent **research project**. It is **not** affiliated with, endorsed by, sponsored by, or otherwise officially related to Microsoft.

## Current Scope

- Fully client-side coding agent runtime (no backend API route)
- Agent skills support (skill context injection + skill toggles)
- Custom agents support via frontmatter
- Host-targeted agents via `hosts` and `defaultForHosts` frontmatter fields
- Host-routed tools and system prompts (Excel implemented first)

## Features

- **AI Chat Panel** — conversational assistant embedded in Excel's task pane
- **83 Excel tools** — the AI can read/write ranges, create charts, manage tables, format cells, add comments, create pivot tables, set data validation, and manipulate sheets
- **Agent system** — split system prompt architecture + custom agents with host targeting (`hosts`, `defaultForHosts`)
- **Skills system** — bundled skill files that inject additional context into the system prompt, toggleable via the SkillPicker (standard Agent Skills layout: one `SKILL.md` plus optional `references/` docs)
- **Custom extension import** — import local ZIP files for custom agents and custom skills from Settings
- **Extension management UX** — manage imported agents/skills in Settings, with bundled content shown as read-only
- **Multi-model support** — add and validate deployed model names from your Azure AI Foundry endpoint (manual entry flow)
- **Model management** — add, rename, and remove models per endpoint; switch models from the input toolbar
- **Streaming responses** — real-time token streaming for fast, responsive interactions
- **Copilot-style UX** — cycling dot progress indicators, collapsible tool progress, choice cards, tool result summaries
- **API key authentication** — simple API key auth per endpoint, no Azure AD app registration required
- **Multiple endpoints** — configure and switch between several Azure AI Foundry resources
- **First-time setup wizard** — guided onboarding flow with auto-discovery and manual model entry

## Agent Skills Format

This project follows the standard Agent Skills model:

- A skill is a folder containing `SKILL.md`.
- Optional supporting docs live under `references/` inside that same skill.

## Prerequisites

- [Node.js](https://nodejs.org/) >= 20
- Microsoft Excel (desktop or Microsoft 365 web)
- An [Azure AI Foundry](https://ai.azure.com/) resource with at least one model deployed

## Getting Started

```bash
# Install dependencies
npm install

# Start the dev server and sideload into Excel Desktop
npm run start:desktop
```

This starts the webpack dev server on `https://localhost:3000` and opens Excel with the add-in sideloaded.

For local shared-folder sideloading and staging manifest workflows, see [docs/SIDELOADING.md](./docs/SIDELOADING.md).

### Environment Variables

Set these before starting the dev server to pre-populate the setup wizard:

| Variable                | Description                                                                  |
| ----------------------- | ---------------------------------------------------------------------------- |
| `AZURE_OPENAI_ENDPOINT` | Azure AI Foundry resource URL (e.g., `https://my-resource.openai.azure.com`) |
| `AZURE_OPENAI_API_KEY`  | API key for the resource (pre-populates the wizard's API key field)          |

```bash
# Example: start with pre-populated endpoint
$env:AZURE_OPENAI_ENDPOINT = "https://my-resource.openai.azure.com"
$env:AZURE_OPENAI_API_KEY = "your-key-here"
npm run start:desktop
```

#### Integration Test Credentials

Integration tests that hit a live Azure AI Foundry endpoint read credentials from a `.env` file in the project root:

| Variable           | Description                                             |
| ------------------ | ------------------------------------------------------- |
| `FOUNDRY_ENDPOINT` | Full resource URL (may include `/api/projects/...`)     |
| `FOUNDRY_API_KEY`  | API key for the resource                                |
| `FOUNDRY_MODEL`    | Model deployment name to test (default: `gpt-5.2-chat`) |

```bash
# .env (gitignored)
FOUNDRY_ENDPOINT=https://your-resource.services.ai.azure.com/api/projects/proj-default
FOUNDRY_API_KEY=your-key-here
FOUNDRY_MODEL=gpt-5.2-chat
```

### First-Time Setup

When you launch the add-in for the first time, a setup wizard guides you through configuration:

#### Step 1 — Connect Your Endpoint

Enter your Azure AI Foundry resource URL (e.g., `https://my-resource.openai.azure.com`) and an optional display name. If the `AZURE_OPENAI_ENDPOINT` environment variable is set, this field is pre-populated.

#### Step 2 — Authentication

Enter your API key from the Azure AI Foundry resource. If the `AZURE_OPENAI_API_KEY` environment variable is set, this field is pre-populated.

#### Step 3 — Model Setup

The wizard validates your endpoint connection, then you add model deployment names manually. The default model (`gpt-5.2-chat`) is auto-validated and pre-added when reachable.

On subsequent launches, the wizard is skipped — you go straight to the chat interface. You can add, remove, or switch endpoints at any time from the Settings dialog (gear icon in the header). If all endpoints or models are removed, the wizard reappears automatically.

## Available Scripts

| Script                           | Description                                        |
| -------------------------------- | -------------------------------------------------- |
| `npm run dev`                    | Start webpack dev server (hot reload)              |
| `npm run build`                  | Production build to `dist/`                        |
| `npm run build:dev`              | Development build to `dist/`                       |
| `npm run start:desktop`          | Build and sideload into Excel Desktop              |
| `npm run stop`                   | Stop debugging / unload the add-in                 |
| `npm run manifest:staging`       | Generate staging manifest pointing to GitHub Pages |
| `npm run extensions:samples`     | Generate sample `agents` and `skills` ZIP files    |
| `npm run sideload:share:setup`   | Create local shared-folder catalog on Windows      |
| `npm run sideload:share:trust`   | Register local share as trusted Office catalog     |
| `npm run sideload:share:publish` | Copy staging manifest into local shared folder     |
| `npm run sideload:share:cleanup` | Remove local share and trusted-catalog setup       |
| `npm run lint`                   | Run ESLint                                         |
| `npm run lint:fix`               | Auto-fix ESLint issues                             |
| `npm run format`                 | Format code with Prettier                          |
| `npm run typecheck`              | Type-check without emitting                        |
| `npm test`                       | Run all Vitest tests                               |
| `npm run test:watch`             | Run tests in watch mode                            |
| `npm run test:coverage`          | Run tests with coverage                            |
| `npm run test:e2e`               | Run E2E tests in Excel Desktop (~187)              |
| `npm run validate`               | Validate `manifest.xml`                            |

## Testing

The project has four layers of tests:

| Layer           | Tool       | Count    | What it covers                                                                  |
| --------------- | ---------- | -------- | ------------------------------------------------------------------------------- |
| **Unit**        | Vitest     | 17 files | Pure functions, Zustand store logic, Zod tool schemas, host/agent parsing       |
| **Integration** | Vitest     | 15 files | Real component wiring + live Azure AI Foundry API + LLM tool calling            |
| **UI**          | Playwright | 14 tests | Browser taskpane flows (chat/settings/wizard)                                   |
| **E2E**         | Mocha      | ~187     | 83 Excel tools + settings persistence + AI round-trip inside real Excel Desktop |

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

# E2E tests (requires Excel Desktop, ~187 tests)
npm run test:e2e

# Validate the Office add-in manifest
npm run validate
```

### Unit Tests

Unit tests in `tests/unit/` (17 files) cover pure functions and store logic that have **no** `Excel.run()` dependency:

| File                                | What it tests                                                                                 |
| ----------------------------------- | --------------------------------------------------------------------------------------------- |
| `agentService.test.ts`              | Agent frontmatter parsing, getAgents, getAgent, getAgentInstructions                          |
| `aiClientFactory.test.ts`           | Provider creation, caching, invalidation, and clear-all behavior                              |
| `buildSkillContext.test.ts`         | `buildSkillContext` and related skill functions with bundled `.md` files                      |
| `chatErrorBoundary.test.tsx`        | Error boundary fallback rendering and recovery flow                                           |
| `chatPanel.test.tsx`                | ChatPanel component logic (mocks assistant-ui components for jsdom)                           |
| `humanizeToolName.test.ts`          | Tool-name formatting for user-facing progress labels                                          |
| `id.test.ts`                        | `generateId` unique ID generation utility                                                     |
| `manifest.test.ts`                  | Manifest and runtime host assumptions used by tests                                           |
| `messagesToCoreMessages.test.ts`    | `messagesToCoreMessages` conversion from ChatMessage[] to core messages                       |
| `modelDiscoveryHelpers.test.ts`     | `inferProvider`, `isEmbeddingOrUtilityModel`, `formatModelName` (table-driven with `it.each`) |
| `normalizeEndpoint.test.ts`         | Endpoint URL normalization (trailing slashes, `/openai` suffixes, Foundry paths)              |
| `officeStorage.test.ts`             | `officeStorage` localStorage fallback (OfficeRuntime undefined in jsdom)                      |
| `officeStorageRuntime.test.ts`      | `officeStorage` behavior when OfficeRuntime is present and throws                             |
| `parseFrontmatter.test.ts`          | YAML frontmatter parsing for skill files (delimiters, multiline scalars, tag arrays)          |
| `settingsStore.test.ts`             | Zustand store: endpoint CRUD, model CRUD, cascade delete, auto-selection, URL dedup           |
| `toolSchemas.test.ts`               | Zod `inputSchema` validation for all 83 tool definitions (valid accepts, invalid rejects)     |
| `useToolInvocations-patch.test.tsx` | assistant-ui patch behavior for tool invocation argument streaming integrity                  |

**Key principle:** unit tests run against the **real** Zustand store with localStorage (jsdom). No mocking.

### Integration Tests

Integration tests in `tests/integration/` (15 files) exercise three categories:

- **Component wiring** (`chat-header-settings-flow.test.tsx`, `wizard-to-chat.test.tsx`, `agent-picker.test.tsx`, `skill-picker.test.tsx`, `chat-panel.test.tsx`, `settings-dialog.test.tsx`, `app-state.test.tsx`, `app-error-boundary.test.tsx`, `model-manager.test.tsx`, `model-picker-interactions.test.tsx`) — renders real components together (no child mocks), verifying cross-component state and interactions.
- **Live API** (`foundry.integration.test.ts`, `chat-pipeline.integration.test.ts`, `multi-turn.integration.test.ts`) — hits a real Azure AI Foundry endpoint to validate client factory, model discovery, streaming, and multi-turn conversations. Requires `FOUNDRY_ENDPOINT` and `FOUNDRY_API_KEY` in `.env`.
- **LLM tool calling** (`llm-tool-calling.integration.test.ts`) — exercises the full ToolLoopAgent pipeline: sends natural-language prompts to a live LLM and verifies it selects the correct Excel tools. Requires live API credentials.
- **Store hydration** (`stale-state.test.tsx`) — tests recovery from stale localStorage data (deleted endpoints, orphaned model IDs).

Integration tests run as part of the default `npm test` suite. Live API tests are skipped automatically when environment variables are not set.

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

The add-in runs a fully **client-side AI agent** — no backend API route is needed. The Vercel AI SDK's `ToolLoopAgent` and `DirectChatTransport` run in-process inside the Excel task pane:

```
┌──────────────────────────────────────────────────────────────┐
│  useOfficeChat(provider, modelId, host)                      │
│                                                              │
│  ┌─────────────────┐    ┌──────────────────────────────────┐ │
│  │  ToolLoopAgent   │    │  DirectChatTransport             │ │
│  │  - model         │───►│  - runs agent in-process         │ │
│  │  - instructions  │    │  - no server / API route needed  │ │
│  │  - tools(host)   │    └──────────────┬───────────────────┘ │
│  │  - stopWhen(10)  │                   │                     │
│  └─────────────────┘                   ▼                     │
│                              useChat (transport)             │
│                              ─────────────────               │
│                              Returns UseChatHelpers          │
│                              (messages, input, handleSubmit) │
└──────────────────────────────────────────────────────────────┘
         │                              │
         ▼                              ▼
  Azure AI Foundry              ChatPanel / ChatHeader
  (streaming LLM)               (React UI via props)
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

- **`useOfficeChat`** — custom hook that creates the host-routed agent, transport, and returns `useChat` helpers
- **`ToolLoopAgent`** — auto-executes Excel tool calls (up to 10 steps) via Zod-typed `execute` handlers
- **`DirectChatTransport`** — routes `useChat` through the agent without an HTTP backend
- **`useChat`** — Vercel AI SDK React hook that manages messages, streaming, and input state

State is minimal: `useSettingsStore` (Zustand) persists endpoint/model/agent configuration; chat state lives entirely in `useChat`.

## UI Layout

The task pane is organized into three areas:

- **ChatHeader** — "AI Chat" title + SkillPicker (icon-only with badge) + New Conversation button + Settings gear
- **ChatPanel** — CopilotChat messages, Copilot-style progress indicators (cycling dots + phase labels), choice cards, error bar, ChatInput, and an **input toolbar** below the text box with AgentPicker + ModelPicker (GitHub Copilot-style)
- **App** — root component that owns settings dialog state, routes between SetupWizard and chat UI, detects system theme

## Authentication

The API key is stored in the browser's `localStorage` as part of the endpoint configuration and sent directly as the `api-key` header to the Azure AI Foundry REST API. No Azure AD app registration is needed.

## Tech Stack

- **React 18** — UI framework
- **assistant-ui + Radix UI + Tailwind CSS v4** — task pane UI components and styling
- **Vercel AI SDK** — `ai` (ToolLoopAgent, DirectChatTransport) + `@ai-sdk/react` (useChat) + `@ai-sdk/azure` for streaming, multi-step tool calling, and client-side agent execution
- **Zustand 5** — lightweight state management with `localStorage` persistence
- **Webpack 5** — bundling with HMR
- **TypeScript 5** — type safety
- **Vitest** — unit, component, and integration testing
- **Playwright** — browser UI testing for task pane flows
- **Mocha** — E2E testing inside Excel Desktop (~187 tests)
- **Testing Library** — React component testing (`@testing-library/react`, `user-event`)
- **dotenv** — environment variable loading for integration tests
- **ESLint + Prettier** — code quality

## Community & Security

- [Code of Conduct](./CODE_OF_CONDUCT.md)
- [Security Policy](./SECURITY.md)

## License

MIT
