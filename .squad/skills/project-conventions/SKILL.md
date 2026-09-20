---
name: "project-conventions"
description: "Use when changing Office tools, Copilot sessions, task-pane UI, or tests in office-coding-agent."
domain: "project-conventions"
confidence: "high"
source: "observed"
---

## Context

This skill guides contributors working on the Office add-in. It is not an Office
task skill loaded into a user's Copilot session. The runtime's Office agents and
skills come from CLI plugins in `sbroenne/office-coding-agent-plugins`; do not
duplicate their instructions here.

Before changing code, identify the affected host (Excel, PowerPoint, Word, or
Outlook), the runtime boundary, and the appropriate test tier. Check current
source and package scripts rather than relying on historical version numbers,
test counts, or architecture summaries.

## Patterns

### Keep browser, proxy, and Office responsibilities separate

- The React task pane communicates over WebSocket/JSON-RPC with the local Node
  proxy, which owns the Copilot SDK client. Keep Node-only APIs and CLI lifecycle
  management out of browser code.
- Trace session changes through `useOfficeChat`, the browser client/transport,
  and `copilotProxy.mjs`. Preserve streaming, cancellation, and session cleanup.
- Route tools through `getToolsForHost`. Unknown hosts receive no tools. Known
  hosts include management tools and are capped at `MAX_TOOLS_PER_REQUEST`;
  check the host tool-count tests when adding tools so none are silently lost.
- For Excel, PowerPoint, and Word, use the existing declarative configs and
  host-specific factories. The factory supplies the Office request context;
  do not nest another host `run` call inside a config's `execute` function.
  Outlook has its own tool implementation; follow that host's existing pattern.

### Keep agents and skills CLI-owned

- Startup bootstrap installs the required Office plugins. The proxy passes
  installed plugin directories to SDK-created sessions so plugin agents resolve.
- The agent picker selects CLI-discovered agents; do not introduce browser-owned
  agent definitions or a separate skills installation UI.
- Slash discovery reads installed plugin `SKILL.md` files and `.prompt.md` files,
  plus workspace prompts. Preserve this distinction when changing suggestions.
- `buildSessionSystemPrompt` combines the base prompt with optional memory.
  Currently `getAppPromptForHost` returns an empty string: do not assume local
  host app prompt files exist. Host-specific instructions belong to the plugins.

### Preserve tool results and settings

- The shared tool factory translates execution errors to SDK failure results;
  do not swallow errors and report success. Preserve image results as
  `binaryResultsForLlm` rather than embedding base64 in model-facing text.
- Persist preferences through the existing Zustand settings store and
  `officeStorage` adapter. Keep discovered model/agent lists out of persisted
  settings; `partialize` explicitly selects the saved fields.
- Read the adapter before changing storage behavior: it prefers
  `OfficeRuntime.storage` and currently falls back to `localStorage` when that
  runtime API is unavailable. Integration tests use `tests/setup.ts`.

### Match VS Code's Copilot Chat

- Reuse the existing chat components and `--vscode-*` theme tokens. Use codicons
  for new icons, compact spacing, and the existing 13px font stack.
- Keep messages full-width, not chat bubbles. Use the existing shimmer for
  thinking/tool progress, not spinners or pulse effects.
- Focus uses a 1px `--vscode-focusBorder` outline, not a box-shadow ring.
- Use `@/` imports for source modules and existing barrel exports where available.
  Follow the repository's Oxlint and Prettier configuration.

### Choose tests by runtime boundary

| Change | Coverage |
| --- | --- |
| React wiring, stores, host routing, tool schemas, proxy/session behavior | Integration tests in `tests/integration/` |
| Task-pane interaction flows | Playwright tests in `tests-ui/`, using the real proxy and Copilot API |
| Real Office API behavior | Mocha E2E in `tests-e2e/`, `tests-e2e-ppt/`, `tests-e2e-word/`, or `tests-e2e-outlook/` |

Do not add isolated unit tests or use fabricated Office contexts as proof that
host operations work. Do not mock network requests, WebSockets, or Copilot
responses in Playwright tests.

For code changes:

1. Run `npm run lint`, `npm run typecheck`, and `npm run build`.
2. Run `npm run test:integration`. Live Copilot tests need the local server and
   working CLI authentication. If the server is missing, start `npm run dev`,
   wait for `https://localhost:3000/api/ping`, and retry.
3. Run `npm run test:ui` for changed task-pane flows. Playwright's configuration
   starts the dev server; Copilot authentication is still required.
4. Run E2E for each affected host: `npm run test:e2e` (Excel),
   `npm run test:e2e:ppt`, `npm run test:e2e:word`, or
   `npm run test:e2e:outlook`. These require real Office Desktop on a supported
   machine; Outlook also requires tenant sideloading approval.

Report missing host/authentication prerequisites and any test failures as
blockers, not passes. `npm test` excludes live-server tests and does not replace
the integration or Office E2E commands. Documentation-only edits do not require
the application test suites unless there is a relevant documentation test.

## Examples and Source References

Paths below are relative to the repository root:

- Host dispatch and tool limits: `src/tools/index.ts`,
  `tests/integration/host-tools-limit.test.ts`.
- Tool execution and SDK results: `src/tools/codegen/factory.ts`.
- Session lifecycle: `src/hooks/useOfficeChat.ts`,
  `src/lib/websocket-client.ts`, `src/lib/websocket-transport.ts`,
  `src/copilotProxy.mjs`.
- CLI plugin loading and slash discovery: `src/plugins/cliPluginBootstrap.mjs`,
  `src/plugins/cliSlashItems.mjs`.
- Prompt composition: `src/services/ai/systemPrompt.ts`.
- Preference persistence: `src/stores/settingsStore.ts`,
  `src/stores/officeStorage.ts`.
- UI patterns: `src/components/chat/`, `src/styles/vscode-theme.css`.
- Test commands and prerequisites: `package.json`, `vitest.config.ts`,
  `playwright.config.ts`, and the affected host's `runner.test.ts`.

## Anti-Patterns

- Calling Copilot directly from the browser or importing Node APIs into the task pane.
- Recreating plugin-owned Office skills or agents in the add-in.
- Treating schema/component tests as validation of actual Office API execution.
- Hardcoding theme colors or introducing another icon library for new UI.
- Updating reusable skill templates with project-specific facts.
- Committing directly to `main`; changes require a feature branch and PR, with
  **Squash and merge** on GitHub.
