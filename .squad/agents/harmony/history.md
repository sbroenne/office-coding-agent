# Project Context

- **Owner:** Stefan Broenner
- **Project:** office-coding-agent — Microsoft Office add-in bringing GitHub Copilot into Excel, PowerPoint, Word, and Outlook with full Copilot CLI plugin support
- **Stack:** React 18, TypeScript, Node.js, Tailwind CSS v4, Copilot SDK, WebSocket + JSON-RPC, Vite 7, Vitest, Playwright, Mocha (E2E)
- **Created:** 2026-03-15

## Learnings

<!-- Append new learnings below. Each entry is something lasting about the project. -->

### 2026-03-16 — Architecture review
- The browser ↔ proxy ↔ Office layering is conceptually clean, but `src/server.mjs` also performs plugin marketplace registration and host plugin installation on startup, so the proxy currently owns product bootstrap side effects as well as transport.
- Host routing is mostly consistent across `detectOfficeHost`, `getToolsForHost`, `systemPrompt.ts`, and `agentService.ts`, but the fallback behavior still defaults missing/unknown Office context to `excel`, which leaks Excel assumptions outside the real host boundary.
- Tool definition architecture is fragmented: Excel uses decomposed config modules + factory, PowerPoint and Word keep very large inline config arrays, and Outlook bypasses the config/factory path entirely with handwritten `Tool[]`. That split is now the biggest source of drift risk.
- The prompt stack is coherent at the top level (BASE + host app prompt + agent instructions), but runtime orchestration behavior for PowerPoint and Word lives in `useOfficeChat.ts` as host-specific heuristics instead of behind a host plugin/registry abstraction.
- Skills are not a first-class architecture layer in the current app code: plugin skills are discovered and shown in the UI, but there is no bundled `src/skills/` layer or local skill-context assembly path matching the documented architecture.
